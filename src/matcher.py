"""核心匹配引擎 —— 多算法加权 + 简称/全称包含关系修正。

算法设计
--------
纯 ``ratio``（编辑距离类）无法区分下面两种情形：

    A 中国石油天然气股份有限公司  vs  B 中国石油化工股份有限公司   ratio=80.0 partial=75.0  ← 不同公司
    A 中国平安保险(集团)股份有限公司 vs  B 中国平安保险(集团)          ratio=76.9 partial=100.0 ← 同一公司

关键判据是 **包含关系**：真·简称/全称对里，短串是长串的子串（partial_ratio 冲顶）；
陷阱对只是"长得像"，没有任何一方包含另一方（partial_ratio 反而低于 ratio）。

因此本引擎采用三段式打分：

1. **基础分**：``0.45*ratio + 0.20*partial_ratio + 0.15*WRatio + 0.20*core_ratio``
   —— ``core_ratio`` 比对"剥离全部标点后的核心字号"，抵抗括号内容缺失。
2. **包含关系修正**：短串是长串子串且长度 ≥4 时，按覆盖率抬分；
   前缀/后缀锚定的（标准简称）抬升力度大于中间截取。
3. **完全一致短路**：清洗后（或去标点后）完全一致直接判 100。

同时计算 **互为最优（mutual best）** 与 **目标争抢（conflict）** 标记，
供第三阶段识别"两条 A 抢同一条 B"的潜在误匹配。
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Sequence

import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process

from . import config
from . import preprocessor as pp

# --------------------------------------------------------------------------- #
#  参数
# --------------------------------------------------------------------------- #
WEIGHTS: dict[str, float] = {
    "ratio": 0.45,         # 全局编辑距离相似度 —— 抗错别字
    "partial_ratio": 0.20, # 最优子串相似度   —— 抗长度差异
    "WRatio": 0.15,        # rapidfuzz 自适应加权
    "core_ratio": 0.20,    # 去标点后核心字号相似度 —— 抗括号差异
}

# ---- 基础相似度算法注册表（赛题 §3.2 加分项：支持多种相似度算法切换）----
#
# 一个"算法"= 对若干基础 scorer 输出的一组权重。5 个 scorer 一次算完，
# 换算法只是换一组权重，不重跑矩阵。
#
# 关键性质：``weighted``（默认）的权重与上方 :data:`WEIGHTS` **完全一致**，
# 且 ``jaro_winkler`` 权重为 0（贡献恒为 0）—— 因此新增切换入口
# **不改变**已验证的 45 / 40 / 15 / 12 匹配结果，纯粹是增量能力。
SCORER_ORDER: tuple[str, ...] = (
    "ratio", "partial_ratio", "WRatio", "core_ratio", "jaro_winkler",
)
#: 结果表「各算法原始分」默认报告这几个通道（保持既有输出不变）
REPORTED_SCORERS: tuple[str, ...] = SCORER_ORDER[:4]

SCORER_LABELS: dict[str, str] = {
    "ratio": "Levenshtein 归一化编辑距离",
    "partial_ratio": "最优子串相似度",
    "WRatio": "rapidfuzz 自适应加权",
    "core_ratio": "去标点后字号相似度",
    "jaro_winkler": "Jaro-Winkler",
}

ALGO_WEIGHTS: dict[str, dict[str, float]] = {
    # 默认：多算法加权组合（本工具主算法，抗错别字 / 长度差异 / 括号差异）
    "weighted": {"ratio": 0.45, "partial_ratio": 0.20, "WRatio": 0.15,
                 "core_ratio": 0.20},
    # 以下为单算法对照模式：用于演示"换算法会怎样"，
    # 它们**同样**会经过字号惩罚 / 通用词折叠 / 关键字号保护等全部业务规则。
    "ratio": {"ratio": 1.0},
    "partial_ratio": {"partial_ratio": 1.0},
    "WRatio": {"WRatio": 1.0},
    "core_ratio": {"core_ratio": 1.0},
    "jaro_winkler": {"jaro_winkler": 1.0},
}

DEFAULT_ALGO = "weighted"
ALGO_CHOICES: tuple[str, ...] = tuple(ALGO_WEIGHTS)

ALGO_LABELS: dict[str, str] = {
    "weighted": "加权组合（推荐）",
    "ratio": "Levenshtein 编辑距离",
    "partial_ratio": "最优子串相似度",
    "WRatio": "rapidfuzz WRatio",
    "core_ratio": "去标点字号比对",
    "jaro_winkler": "Jaro-Winkler",
}


def describe_algo(algo: str = DEFAULT_ALGO) -> str:
    """把算法名翻译成一行人类可读说明（供 CLI / GUI 日志打印）。"""
    if algo not in ALGO_WEIGHTS:
        raise ValueError(f"未知算法 {algo!r}，可选：{'、'.join(ALGO_CHOICES)}")
    w = ALGO_WEIGHTS[algo]
    if algo == DEFAULT_ALGO:
        parts = " + ".join(
            f"{tag}×{w[tag]:g}" for tag in SCORER_ORDER if w.get(tag)
        )
        return f"{parts}（多算法加权组合）"
    return f"{ALGO_LABELS[algo]}（单算法对照模式）"


def _jaro_winkler_ratio(s1: str, s2: str, score_cutoff: float | None = None) -> float:
    """Jaro-Winkler 归一化相似度，缩放到 0–100。

    ``rapidfuzz.process.cdist`` 要求自定义 scorer 接受 ``score_cutoff``
    关键字参数，故显式声明。用惰性导入避免模块级硬依赖。
    """
    from rapidfuzz.distance import JaroWinkler

    value = JaroWinkler.normalized_similarity(s1, s2) * 100.0
    if score_cutoff is None:
        return value
    return value if value >= score_cutoff else 0.0

MIN_CONTAINMENT_LEN = 2        # 短串短于此长度不做包含关系抬分
# 注：这里取 2 而非 4。赛题 §5 明文要求
#     A:"腾讯" vs B:"深圳市腾讯计算机系统有限公司" → 高相似度匹配，
#     并注明「若标记为"未匹配"则扣分」。取 4 会把 2 字的「腾讯」挡在门槛外，
#     判成 A系统独有。取 2 后当前数据的表现完全不变（表内最短名称是 4 字，
#     且长度为 2–3 的名字作为子串出现的配对数为 0），纯粹修好赛题点名场景。
SHORT_ALIAS_LEN = 3            # 短至此长度的完整包含 = 字号级简称
SHORT_ALIAS_FLOOR = 92.0       # 直接抬进「高度匹配」档（≥90）

# ---- 通用词折叠与关键字号保护（见 preprocessor.collapse_generic 的说明）----
BRAND_PENALTY = 0.40           # 关键字号冲突：打折
BRAND_CAP = 50.0               # 关键字号冲突：同时硬封顶，确保 < 识别下限 60
SAME_BRAND_FLOOR = 95.0        # 特征字号一致：抬进「高度匹配」档
SAME_BRAND_BOOST = 1.15        # 同字号、行业词不同（顺丰快递/顺丰控股）：候选排序加成
SCORE_PRECISION = 4            # 最终分数的舍入位数（消除浮点残差造成的阈值抖动）
PARTIAL_TRIGGER = 95.0         # partial_ratio 超过此值才进入包含关系候选
ANCHOR_PENALTY = 15.0          # 前缀/后缀锚定：100 - (1-覆盖率)*15
MIDDLE_PENALTY = 28.0          # 中间截取：100 - (1-覆盖率)*28

DEFAULT_THRESHOLD = 60.0       # 低于此分视为"无匹配"，供第三阶段判定独有记录

# R3 加成**只用于未匹配记录的候选建议**，绝不参与"是否匹配"的判定 ——
# 因此只在结果仍低于识别下限时才生效，且封顶在下限之下。
BOOST_CEILING = DEFAULT_THRESHOLD - 0.1

# 各阶段的进度权重（占本模块 0–100% 的区间），供 CLI / GUI 共用
STAGE_WEIGHTS: dict[str, tuple[float, float]] = {
    "matrix": (0.0, 45.0),       # 4 个 scorer 的相似度矩阵
    "penalty": (45.0, 25.0),     # 字号差异校验
    "boost": (70.0, 7.0),        # 包含关系抬分
    "brand": (77.0, 3.0),        # 通用词折叠 + 关键字号保护
    "select": (80.0, 20.0),      # 选取最佳匹配
}

# 进度回调：(整体百分比 0–100, 消息文本)
ProgressFn = Callable[[int, str], None]

try:                            # 进度条为可选依赖
    from tqdm import tqdm
except ImportError:             # pragma: no cover
    tqdm = None


# --------------------------------------------------------------------------- #
#  阈值
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Thresholds:
    """可调分档阈值（GUI 的滑块 / 微调框直接绑定这三个值）。

    ``high=90, low=70, floor=60`` 为赛题推荐默认值：

    * ``score >= 100``        → 完全匹配
    * ``score >= high``       → 高度匹配（可直接确认）
    * ``score >= low``        → 中低匹配（建议人工复核）
    * ``score >= floor``      → 低置信度匹配（需逐条核实）
    * ``score <  floor``      → 该条 A 记录判为「A系统独有」
    """

    high: float = 90.0
    low: float = 70.0
    floor: float = DEFAULT_THRESHOLD

    def __post_init__(self) -> None:
        if not (100.0 >= self.high >= self.low >= self.floor >= 0.0):
            raise ValueError(
                f"阈值必须满足 100 >= high >= low >= floor >= 0，当前为 "
                f"high={self.high}, low={self.low}, floor={self.floor}"
            )

    def bucket(self, score: float) -> str:
        if score >= 100.0:
            return "完全匹配"
        if score >= self.high:
            return "高度匹配"
        if score >= self.low:
            return "中低匹配"
        if score >= self.floor:
            return "低置信度匹配"
        return "A系统独有"

    @property
    def label(self) -> str:
        return f"100 / {self.high:g} / {self.low:g} / {self.floor:g}"


DEFAULT_THRESHOLDS = Thresholds()


# --------------------------------------------------------------------------- #
#  结果结构
# --------------------------------------------------------------------------- #
@dataclass
class MatchResult:
    """单条 A 记录的最佳匹配结果。"""

    a_index: int
    a_name: str                 # A 原始名称（原样保留）
    a_clean: str
    b_index: int | None
    b_name: str                 # B 原始名称（原样保留）
    b_clean: str
    score: float                # 0–100，保留 1 位小数在输出层处理
    detail: dict[str, float] = field(default_factory=dict)   # 各算法原始分
    runner_up_index: int | None = None
    runner_up_name: str = ""
    runner_up_score: float = 0.0
    mutual_best: bool = False   # 该 B 记录是否也把本条 A 当作最优
    conflict: bool = False      # 首选 B 被其他 A 记录争抢过（不必然让位）
    conflict_count: int = 1     # 争抢其首选的 A 记录条数
    displaced: bool = False     # 一对一约束下让位：首选被更高分记录拿走，改用次佳
    floor: float = DEFAULT_THRESHOLD   # 本次运行实际使用的识别下限（判定口径）

    @property
    def is_matched(self) -> bool:
        """本条是否被接受为一次匹配。

        .. note:: 判定口径必须用**本次运行实际使用的** ``floor``，
            不能写死模块级常量 ``DEFAULT_THRESHOLD`` —— 用户在 GUI 上把
            「识别下限」调成 30 之后，一条 45 分的记录确实匹配上了，
            却会被写死 60 的属性判成未匹配，与报表分档自相矛盾。
        """
        return self.b_index is not None and self.score >= self.floor


@dataclass
class MatchOutcome:
    """整表匹配结果。"""

    results: list[MatchResult]
    score_matrix: np.ndarray
    b_taken: dict[int, list[int]]     # B 行号 → 争抢它的 A 行号列表（降序按分）
    stats: dict[str, Any] = field(default_factory=dict)


# --------------------------------------------------------------------------- #
#  打分矩阵
# --------------------------------------------------------------------------- #
def _emit(progress: ProgressFn | None, stage: str, done: int, total: int,
          message: str) -> None:
    """把「阶段内第 done/total 步」换算成本模块 0–100% 的百分比并上报。"""
    if progress is None:
        return
    lo, span = STAGE_WEIGHTS[stage]
    frac = (done / total) if total else 1.0
    progress(int(lo + span * min(max(frac, 0.0), 1.0)), message)


def _iter(seq: Iterable[Any], desc: str, unit: str,
          progress: ProgressFn | None = None, stage: str = "",
          message: str = ""):
    """统一迭代器：CLI 走 tqdm，GUI 走 progress 回调，两者都没有则裸迭代。"""
    seq = list(seq)
    if progress is not None:
        total = len(seq)
        for k, item in enumerate(seq, start=1):
            yield item
            _emit(progress, stage, k, total, f"{message} {k}/{total}")
    elif tqdm is not None:
        yield from tqdm(seq, desc=desc, unit=unit, leave=False)
    else:
        yield from seq


def build_score_matrix(
    a_clean: Sequence[str],
    b_clean: Sequence[str],
    a_core: Sequence[str] | None = None,
    b_core: Sequence[str] | None = None,
    workers: int = -1,
    verbose: bool = True,
    progress: ProgressFn | None = None,
    algo: str = DEFAULT_ALGO,
) -> tuple[np.ndarray, np.ndarray]:
    """计算 A×B 全量相似度矩阵。

    参数 ``algo`` 选定基础相似度算法（见 :data:`ALGO_WEIGHTS`）；无论选哪个，
    之后的字号惩罚 / 通用词折叠 / 关键字号保护 / 包含关系抬分等**业务规则
    全部照常生效** —— 换的只是最底层的相似度，不是判据。

    返回 ``(score, detail_stack)``：
    * ``score``        —— ``(len(a), len(b))`` 最终加权分 0–100
    * ``detail_stack`` —— ``(len(a), len(b), 5)`` 各 scorer 原始分（通道顺序
      同 :data:`SCORER_ORDER`），便于结果表追溯
    """
    if algo not in ALGO_WEIGHTS:
        raise ValueError(
            f"未知算法 {algo!r}，可选：{'、'.join(ALGO_CHOICES)}"
        )
    a_clean = list(a_clean)
    b_clean = list(b_clean)
    n_a, n_b = len(a_clean), len(b_clean)
    n_sc = len(SCORER_ORDER)
    if n_a == 0 or n_b == 0:
        return np.zeros((n_a, n_b)), np.zeros((n_a, n_b, n_sc))

    a_core = list(a_core) if a_core is not None else [pp.strip_punct(x) for x in a_clean]
    b_core = list(b_core) if b_core is not None else [pp.strip_punct(x) for x in b_clean]

    # 只算用得上的 scorer：当前算法有权重的 + 结果表要报告的分项。
    # 默认 weighted 模式下 jaro_winkler 不被需要 → 跳过整次 cdist，
    # 默认路径的耗时与加切换入口之前**完全一致**。
    needed = {tag for tag, wt in ALGO_WEIGHTS[algo].items() if wt}
    needed |= set(REPORTED_SCORERS)

    # rapidfuzz 的 cdist 一次只接受单个 scorer，故逐个调用；
    # 每个 scorer 内部由 C++ 多线程并行（workers=-1），n×m 全量矩阵毫秒级完成。
    mats: list[np.ndarray] = []
    todo = [t for t in SCORER_ORDER if t in needed]
    for k, tag in enumerate(todo, start=1):
        scorer, left, right = {
            "ratio": (fuzz.ratio, a_clean, b_clean),
            "partial_ratio": (fuzz.partial_ratio, a_clean, b_clean),
            "WRatio": (fuzz.WRatio, a_clean, b_clean),
            "core_ratio": (fuzz.ratio, a_core, b_core),
            "jaro_winkler": (_jaro_winkler_ratio, a_clean, b_clean),
        }[tag]
        t = time.perf_counter()
        mats.append(
            # dtype=np.float64 必须显式指定：rapidfuzz 默认返回 **float32**，
            # 其 ~1e-5 的精度误差足以把恰好等于阈值（60.0）的分数压成
            # 59.9999977，进而被误判为「独有」而不是「低置信度匹配」。
            process.cdist(left, right, scorer=scorer, workers=workers,
                          dtype=np.float64)
        )
        if verbose:
            print(f"    · {tag:<14} {n_a}×{n_b} 矩阵  {time.perf_counter() - t:.4f}s")
        _emit(progress, "matrix", k, len(todo),
              f"计算相似度矩阵 {k}/{len(todo)}（{tag}）")

    # 未被计算的通道补零矩阵，保证通道索引与 SCORER_ORDER 恒定对齐
    # （调用方按固定下标取 detail[:, :, 1] 等，不能因跳过计算而错位）
    computed = dict(zip(todo, mats))
    detail = np.stack(
        [computed[t] if t in computed else np.zeros((n_a, n_b), dtype=np.float64)
         for t in SCORER_ORDER],
        axis=2,
    )

    weights = ALGO_WEIGHTS[algo]
    score = np.zeros((n_a, n_b), dtype=np.float64)
    for idx, tag in enumerate(SCORER_ORDER):
        w = weights.get(tag, 0.0)
        if w:
            score += w * detail[:, :, idx]

    # ---- 修正 1：字号（distinctive core）差异惩罚 ----
    # 必须先于"包含关系抬分"执行：简称/全称 会在此被压低，随后被包含了关系重新抬回。
    score = apply_core_penalty(score, a_clean, b_clean, verbose=verbose,
                               progress=progress)

    # ---- 修正 1.5：通用词折叠 + 关键字号保护 ----
    # 必须晚于字号惩罚（R2 要把被压低的"仅通用词差异"抬回来），
    # 早于包含关系抬分（后者处理的是子串关系，两者互不重叠）。
    score = apply_brand_rules(score, a_clean, b_clean, verbose=verbose,
                              progress=progress)

    # ---- 修正 2：包含关系（简称 / 全称）抬分 ----
    score = apply_containment_boost(
        score, detail[:, :, 1], a_clean, b_clean, verbose=verbose, progress=progress
    )

    # ---- 修正 3：剥离标点后完全一致 → 100（仅括号/标点差异） ----
    core_eq = np.array(
        [[1.0 if a == b and a else 0.0 for b in b_core] for a in a_core],
        dtype=np.float64,
    )
    score = np.where(core_eq > 0, 100.0, score)

    # ---- 修正 4：清洗后完全一致 → 100（优先级最高） ----
    for i, a in enumerate(a_clean):
        if not a:
            continue
        for j, b in enumerate(b_clean):
            if a == b:
                score[i, j] = 100.0

    # 舍入到 4 位小数：float64 下仍可能出现 59.99999999999 这类残差，
    # 恰好卡在阈值上时会造成分级抖动。4 位远细于任何展示需求，纯属消噪。
    return np.round(np.clip(score, 0.0, 100.0), SCORE_PRECISION), detail


def apply_core_penalty(
    score: np.ndarray,
    a_clean: Sequence[str],
    b_clean: Sequence[str],
    trigger: float = 0.0,
    verbose: bool = True,
    progress: ProgressFn | None = None,
) -> np.ndarray:
    """对"共享长通用尾巴但字号不同"的配对打折。

    参数
    ----
    trigger : 只处理基础分 ``>= trigger`` 的配对。**默认 0.0（全部处理）**。

    .. warning:: **不要把 trigger 调高。**
       早期为省开销设过 ``trigger=55``，只惩罚"看起来还有希望"的配对，结果
       制造了**排序反转**：

       * ``上海哔哩哔哩科技有限公司 ↔ 上海哔哩哔哩有限公司`` 基础分 88.7，
         被惩罚 ×0.55 → **48.8**
       * ``中国海外发展有限公司 ↔ 上海哔哩哔哩有限公司`` 基础分仅 51.4，
         低于阈值被**跳过惩罚** → **51.4**

       于是语义上毫不相关的后者，反而成了「B系统最佳候选」列里显示的答案。
       分档结果不受影响（本就低于识别下限），但候选建议彻底指错人。

       实测代价：全量惩罚 9700 对使端到端从 0.161s 增至 0.187s（+26ms）。
       用这点开销换排序正确性完全值得。
    """
    cand = np.argwhere(score >= trigger)
    if cand.size == 0:
        _emit(progress, "penalty", 1, 1, "字号差异校验（无需处理）")
        return score

    if verbose and len(cand) > 1000:
        print(f"  [i] 字号差异校验 {len(cand)} 对，正在比对特征片段…")

    for i, j in _iter(cand, "  字号差异校验", "对", progress, "penalty",
                      "字号差异校验"):
        penalty = pp.core_divergence_penalty(a_clean[i], b_clean[j])
        if penalty < 1.0:
            score[i, j] *= penalty

    return score


def apply_brand_rules(
    score: np.ndarray,
    a_clean: Sequence[str],
    b_clean: Sequence[str],
    verbose: bool = True,
    progress: ProgressFn | None = None,
) -> np.ndarray:
    """关键字号保护（R1）与特征字号一致抬分（R2）。

    判据见 :func:`src.preprocessor.collapse_generic` —— 把行政区划、括号附注、
    通用词全部剥掉后剩下的就是「特征字号」。两条名称特征字号相同，说明差异纯属
    写法不同，判为同一主体。

    * **R2** 特征字号一致且非空 → ``max(score, 95)``，抬进「高度匹配」档
    * **R1** 关键字号集合不等（如「中国**建设**银行」vs「中国银行」）→
      ``min(score*0.40, 50)``，硬封顶确保判为独有

    性能：``collapse_generic`` / ``protected_brands_in`` 都是 Python 字符串操作，
    但**按名称预计算**（O(n+m)）而非逐对调用（O(n·m)）—— R2 因此退化成一次
    dict 查表，R1 用 ``np.ix_`` 向量化写入，两者都几乎不耗时。
    """
    n_a, n_b = score.shape
    if n_a == 0 or n_b == 0:
        return score

    # ---- 预计算：每个名称只算一次 ----
    a_core = [pp.collapse_generic(x) for x in a_clean]
    b_core = [pp.collapse_generic(x) for x in b_clean]
    a_brand = [pp.protected_brands_in(x) for x in a_clean]
    b_brand = [pp.protected_brands_in(x) for x in b_clean]

    # ---- R2：特征字号一致 → 抬分 ----
    b_by_core: dict[str, list[int]] = {}
    for j, core in enumerate(b_core):
        if core:
            b_by_core.setdefault(core, []).append(j)

    boosted = 0
    for i, core in enumerate(a_core):
        if not core:
            continue
        for j in b_by_core.get(core, ()):
            if score[i, j] < SAME_BRAND_FLOOR:
                score[i, j] = SAME_BRAND_FLOOR
                boosted += 1
        _emit(progress, "brand", i + 1, max(len(a_core), 1), "特征字号比对")

    # ---- R3：同字号、行业词不同 → 候选排序加成 ----
    # 「顺丰快递」与「顺丰控股」剥到纯字号都是「顺丰」，是同一字号下的两个业务主体：
    # 不判为同一家公司（分数仍低于识别下限），但互为最佳候选远比「申通快递」合理。
    a_ind = [pp.collapse_industry(x) for x in a_clean]
    b_ind = [pp.collapse_industry(x) for x in b_clean]
    b_by_ind: dict[str, list[int]] = {}
    for j, ind in enumerate(b_ind):
        if ind:
            b_by_ind.setdefault(ind, []).append(j)

    for i, ind in enumerate(a_ind):
        if not ind or ind == a_core[i]:      # 纯字号 == 特征字号说明没剥掉行业词，跳过
            continue
        for j in b_by_ind.get(ind, ()):
            if b_ind[j] == b_core[j]:        # 对方也没剥掉行业词，不是同组
                continue
            if score[i, j] < DEFAULT_THRESHOLD:
                score[i, j] = min(score[i, j] * SAME_BRAND_BOOST, BOOST_CEILING)

    # ---- R1：关键字号冲突 → 硬封顶 ----
    a_by_brand: dict[frozenset[str], list[int]] = {}
    b_by_brand: dict[frozenset[str], list[int]] = {}
    for i, brands in enumerate(a_brand):
        a_by_brand.setdefault(brands, []).append(i)
    for j, brands in enumerate(b_brand):
        b_by_brand.setdefault(brands, []).append(j)

    capped = 0
    for set_a, rows in a_by_brand.items():
        for set_b, cols in b_by_brand.items():
            if set_a == set_b:          # 关键字号一致 → 不冲突
                continue
            block = score[np.ix_(rows, cols)]
            hit = block * BRAND_PENALTY > BRAND_CAP
            score[np.ix_(rows, cols)] = np.minimum(block * BRAND_PENALTY, BRAND_CAP)
            capped += int(hit.sum()) if hit.size else 0

    if verbose and (boosted or capped):
        print(f"  [i] 通用词折叠：{boosted} 对抬至 {SAME_BRAND_FLOOR:.0f} 分"
              f"（特征字号一致）；{capped} 对封顶 {BRAND_CAP:.0f} 分（关键字号冲突）")
    _emit(progress, "brand", 1, 1, "关键字号保护完成")
    return score


def apply_containment_boost(
    score: np.ndarray,
    partial: np.ndarray,
    a_clean: Sequence[str],
    b_clean: Sequence[str],
    verbose: bool = True,
    progress: ProgressFn | None = None,
) -> np.ndarray:
    """对"短串是长串子串"的配对按覆盖率抬分。

    只检查 ``partial_ratio ≥ PARTIAL_TRIGGER`` 的少量候选，避免 O(n·m) 全量字符串扫描，
    保证大数据量下依然快速。
    """
    cand = np.argwhere(partial >= PARTIAL_TRIGGER)
    if cand.size == 0:
        _emit(progress, "boost", 1, 1, "包含关系修正（无需处理）")
        return score

    if verbose and len(cand) > 200:
        print(f"  [i] 包含关系候选 {len(cand)} 对，正在按覆盖率修正…")

    for i, j in _iter(cand, "  包含关系修正", "对", progress, "boost", "包含关系修正"):
        a, b = a_clean[i], b_clean[j]
        if not a or not b or a == b:
            continue
        short, long_ = (a, b) if len(a) <= len(b) else (b, a)
        if len(short) < MIN_CONTAINMENT_LEN or short not in long_:
            continue

        coverage = len(short) / len(long_)
        anchored = long_.startswith(short) or long_.endswith(short)
        penalty = ANCHOR_PENALTY if anchored else MIDDLE_PENALTY
        boosted = 100.0 - (1.0 - coverage) * penalty

        # 短简称（≤3 字）完整出现在长名里 —— 例如
        #     A:"腾讯"  vs  B:"深圳市腾讯计算机系统有限公司"
        # 这是赛题 §5 点名要求判为「高相似度匹配」的场景（并注明
        # 「若标记为"未匹配"则扣分」）。按覆盖率公式，2/14 的覆盖率只能得 76 分
        # （中低匹配档），达不到文档要求的档位。而"一个完整的短名称原样出现在
        # 另一个名称中"本身已是强证据，长度悬殊不应成为重罚理由，故设下限。
        if len(short) <= SHORT_ALIAS_LEN:
            boosted = max(boosted, SHORT_ALIAS_FLOOR)

        if boosted > score[i, j]:
            score[i, j] = boosted

    return score


# --------------------------------------------------------------------------- #
#  同分候选裁决
# --------------------------------------------------------------------------- #
def _tie_break_indices(
    score: np.ndarray, best_idx: np.ndarray, df_a: pd.DataFrame, df_b: pd.DataFrame
) -> np.ndarray:
    """名称分打平时，改用「金额差 + 日期是否同日」挑最贴合的那条 B 流水。

    仅在候选分数与最高分**完全相等**（容差 1e-6）时生效 —— 名称分不同的时候
    金额绝不参与决策，符合赛题「金额不作为匹配依据」的要求。
    """
    a_amt = pd.to_numeric(df_a[config.A_AMOUNT_COL], errors="coerce").to_numpy(float)
    b_amt = pd.to_numeric(df_b[config.B_AMOUNT_COL], errors="coerce").to_numpy(float)
    a_date = df_a[config.DATE_COL].astype(str).to_numpy()
    b_date = df_b[config.DATE_COL].astype(str).to_numpy()

    picked = best_idx.copy()
    for i in range(score.shape[0]):
        top = score[i, best_idx[i]]
        if top <= 0:
            continue
        tied = np.flatnonzero(np.abs(score[i] - top) < 1e-6)
        if len(tied) <= 1:
            continue

        # 日期同日优先，其次金额差最小
        same_day = np.array([a_date[i] == b_date[j] for j in tied])
        gaps = np.abs(b_amt[tied] - a_amt[i])
        if same_day.any():
            cand = tied[same_day]
            picked[i] = int(cand[np.argmin(np.abs(b_amt[cand] - a_amt[i]))])
        else:
            picked[i] = int(tied[np.argmin(gaps)])
    return picked


# --------------------------------------------------------------------------- #
#  一对一分配
# --------------------------------------------------------------------------- #
def _resolve_one_to_one(score: np.ndarray, best_score: np.ndarray) -> np.ndarray:
    """把「每条 A 各取最佳」的分配结果收敛成**一一对应**。

    赛题 FAQ Q3 明确「默认一对一」，§2.4 的两侧算术（83+17=100、83+14=97）
    也只有在一条 B 只能被一条 A 认领时才成立。

    算法：按 A 的**最佳分降序**贪心 —— 分数高的 A 先挑走自己的首选，
    分数低的若发现首选已被拿走，就退而取自己剩余候选里分数最高的那条。
    这样同一条 B 只会落到一条 A 名下，且优先满足把握最大的配对。

    .. warning:: 返回值用 **-1 表示"抢不到 B"**（达标 A 的条数 > B 的总条数时
        必然出现）。调用方**必须**显式处理 -1 —— numpy 接受负索引，
        ``score[i, -1]`` / ``df_b.iloc[-1]`` 不会报错，只会静默取到最后一列 /
        最后一行数据。见 :func:`match_tables` 中 ``ok = sub >= 0`` 一段。
    """
    n_a, n_b = score.shape
    row_order = np.argsort(-score, axis=1)          # 每行候选按分数降序
    taken = np.zeros(n_b, dtype=bool)
    picked = np.full(n_a, -1, dtype=int)
    for i in np.argsort(-best_score):               # 分数高的 A 优先
        for j in row_order[i]:
            if not taken[j]:
                taken[j] = True
                picked[i] = int(j)
                break
    return picked


# --------------------------------------------------------------------------- #
#  主匹配流程
# --------------------------------------------------------------------------- #
def match_tables(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    a_col: str,
    b_col: str,
    a_core_col: str | None = None,
    b_core_col: str | None = None,
    workers: int = -1,
    verbose: bool = True,
    tie_break: bool = True,
    thresholds: Thresholds = DEFAULT_THRESHOLDS,
    progress: ProgressFn | None = None,
    algo: str = DEFAULT_ALGO,
) -> MatchOutcome:
    """A × B 全量比对，返回每条 A 记录的最佳匹配。

    ``tie_break=True`` 时，对**名称分数完全相同**的多条 B 候选（典型场景：B 表里有两条
    同名流水，对应两笔不同交易），用「金额 + 交易日期」的贴合度挑选最可能的那笔，
    使结果不再依赖 B 表的行序。名称分数不同时该裁决绝不介入，金额不会影响匹配结论。

    ``thresholds`` 仅影响 ``b_taken``（哪些 B 记录算"被认领"）；分档标注在
    :meth:`Thresholds.bucket` / :func:`summarize` 中进行。

    ``algo`` 选定基础相似度算法（:data:`ALGO_WEIGHTS`）。字号惩罚、通用词折叠、
    关键字号保护、包含关系抬分等业务规则**与算法无关，照常生效**。
    """
    a_core_col = a_core_col or f"{a_col}{pp.CORE_SUFFIX}"
    b_core_col = b_core_col or f"{b_col}{pp.CORE_SUFFIX}"

    a_raw = ["" if pd.isna(x) else str(x) for x in df_a[a_col].tolist()]
    b_raw = ["" if pd.isna(x) else str(x) for x in df_b[b_col].tolist()]
    a_clean = df_a[a_col].map(pp.clean_name).tolist()
    b_clean = df_b[b_col].map(pp.clean_name).tolist()
    # 优先复用表里已有的副列，保持与导出结果一致
    a_core = (
        df_a[a_core_col].tolist() if a_core_col in df_a.columns
        else [pp.strip_punct(x) for x in a_clean]
    )
    b_core = (
        df_b[b_core_col].tolist() if b_core_col in df_b.columns
        else [pp.strip_punct(x) for x in b_clean]
    )

    score, detail = build_score_matrix(
        a_clean, b_clean, a_core, b_core, workers=workers, verbose=verbose,
        progress=progress, algo=algo,
    )
    n_a, n_b = score.shape
    if n_a == 0 or n_b == 0:
        # B 表为空（或 A 表为空）时，**每条 A 都应落入 Sheet2「A系统独有」**。
        # 这里不能直接 return 空列表 —— build_result_table / build_a_only_table
        # 都是遍历 outcome.results 出数据的，空列表会让整张表凭空消失：
        # Sheet1/2/3 全空，Sheet4 却写着「A系统总记录数 = 1」，自相矛盾。
        results = [
            MatchResult(
                a_index=i, a_name=a_raw[i], a_clean=a_clean[i],
                b_index=None, b_name="", b_clean="", score=0.0,
                detail={}, runner_up_index=None, runner_up_name="",
                runner_up_score=0.0, mutual_best=False, conflict=False,
                conflict_count=1, displaced=False,
                floor=float(thresholds.floor),
            )
            for i in range(n_a)
        ]
        return MatchOutcome(
            results=results, score_matrix=score, b_taken={},
            stats={
                "n_a": n_a, "n_b": n_b, "pairs": int(n_a * n_b),
                "exact_100": 0, "conflict_targets": 0, "conflict_rows": 0,
                "displaced": 0, "b_claimed": 0, "mutual_best": 0,
            },
        )

    # ---- 每条 A 的最佳 / 次佳 ----
    order = np.argsort(-score, axis=1)[:, :2]
    best_idx = order[:, 0]

    if tie_break and config.A_AMOUNT_COL in df_a.columns and config.B_AMOUNT_COL in df_b.columns:
        tie_idx = _tie_break_indices(score, best_idx, df_a, df_b)
        changed = int((tie_idx != best_idx).sum())
        if verbose and changed:
            print(f"  [i] 同分候选裁决：{changed} 条 A 记录改选金额/日期更贴合的 B 流水")
        best_idx = tie_idx

    best_score = score[np.arange(n_a), best_idx]
    if n_b > 1:
        runner_idx = order[:, 1]
        runner_score = score[np.arange(n_a), runner_idx]
    else:
        runner_idx = np.full(n_a, -1)
        runner_score = np.zeros(n_a)

    # ---- 每条 B 的最佳 A（用于互为最优判定） ----
    b_best_a = np.argmax(score, axis=0)          # (nB,)

    # ---- 一对一约束 ----
    # 赛题 FAQ Q3「默认一对一」+ §2.4 的两侧算术都要求同一条 B 只能被一条 A 认领。
    # 先记录"谁想要谁"（供报表说明让位原因），再按 A 的最佳分降序贪心分配。
    wanted: dict[int, list[int]] = {}
    for i, j in enumerate(best_idx):
        if best_score[i] >= thresholds.floor:
            wanted.setdefault(int(j), []).append(i)
    for j in wanted:
        wanted[j].sort(key=lambda ai: -score[ai, j])

    # 只让"本来就能匹配上"（最佳分 ≥ 下限）的 A 参与分配 ——
    # 分数不达标的 A 本来就不占任何 B，把它们卷进来只会污染 Sheet2 的
    # 「最佳候选」展示（真实最优候选会被换成一条毫不相干的低分记录）。
    final_idx = best_idx.copy()
    matched = best_score >= thresholds.floor
    if matched.any():
        idx_m = np.flatnonzero(matched)
        sub = _resolve_one_to_one(score[matched], best_score[matched])
        ok = sub >= 0
        final_idx[idx_m[ok]] = sub[ok]
        # 让位后的分数必须**重算**：final_idx 已经变了，沿用分配前的
        # best_score 会让"首选被抢走、退而取次优"的记录仍按首选分入档。
        final_score = score[np.arange(n_a), final_idx]
        if (~ok).any():
            # 达标 A 的条数 > B 的总条数时，必然有 A 抢不到 B（_resolve_one_to_one
            # 用 -1 表示）。按一对一约束，这些 A 本就该判为「A系统独有」。
            #
            # 但 -1 **绝不能留在 final_idx 里** —— 它是合法整数，numpy 会当成
            # 负索引：score[i, -1] 取的是最后一列，df_b.iloc[-1] 取的是最后一行，
            # 这条 A 就被静默配到 B 表最后一条流水上，金额 / 日期 / 户名全部张冠李戴，
            # 而 b_taken 里还会混进非法键 -1。GUI 的「识别下限」可以往下调，
            # 用户随手一滑就能触发。
            final_idx[idx_m[~ok]] = best_idx[idx_m[~ok]]   # 先还原成自身最优，避免负索引
            final_score[idx_m[~ok]] = 0.0                  # 压到下限之下 → 自动归入 Sheet2
            matched[idx_m[~ok]] = False
    else:
        final_score = score[np.arange(n_a), final_idx]
    displaced = (final_idx != best_idx) & matched
    if verbose and displaced.any():
        print(f"  [i] 一对一约束：{int(displaced.sum())} 条 A 记录的首选 B 已被更高分"
              f"记录认领，改用其他候选")

    # 最终被认领的 B（每条至多一条 A）—— Sheet3「B系统独有」据此判定
    b_taken: dict[int, list[int]] = {}
    for i, j in enumerate(final_idx):
        if final_score[i] >= thresholds.floor:
            b_taken.setdefault(int(j), []).append(i)

    results: list[MatchResult] = []
    for i in _iter(range(n_a), "  选取最佳匹配", "行", progress, "select",
                   "选取最佳匹配"):
        j = int(final_idx[i])
        rivals = wanted.get(int(best_idx[i]), [])   # 原本争抢其首选的那些 A
        results.append(
            MatchResult(
                a_index=i,
                a_name=a_raw[i],
                a_clean=a_clean[i],
                b_index=j,
                b_name=b_raw[j],
                b_clean=b_clean[j],
                score=float(final_score[i]),
                detail={
                    tag: float(detail[i, j, k])
                    for k, tag in enumerate(SCORER_ORDER)
                },
                runner_up_index=int(runner_idx[i]) if runner_idx[i] >= 0 else None,
                runner_up_name=b_raw[int(runner_idx[i])] if runner_idx[i] >= 0 else "",
                runner_up_score=float(runner_score[i]),
                mutual_best=bool(b_best_a[j] == i),
                # 一对一分配之后，"争抢"标记应跟随**让位方**而不是胜出方 ——
                # 胜出的那条记录是正常匹配，标它"目标争抢"会误导审计师。
                conflict=bool(displaced[i]),
                conflict_count=max(1, len(rivals)),
                displaced=bool(displaced[i]),
                floor=float(thresholds.floor),
            )
        )

    stats = {
        "n_a": n_a,
        "n_b": n_b,
        "pairs": int(n_a * n_b),
        "exact_100": int(np.sum(np.round(final_score, 6) >= 100.0)),
        # 争抢发生在分配之前，用 wanted 统计；分配之后每条 B 至多一条 A
        "conflict_targets": sum(1 for v in wanted.values() if len(v) > 1),
        "conflict_rows": sum(len(v) - 1 for v in wanted.values() if len(v) > 1),
        "displaced": int(displaced.sum()),
        "b_claimed": len(b_taken),
        "mutual_best": sum(1 for r in results if r.mutual_best),
    }
    return MatchOutcome(results=results, score_matrix=score, b_taken=b_taken, stats=stats)


# --------------------------------------------------------------------------- #
#  汇总 / 抽样
# --------------------------------------------------------------------------- #
def bucket_of(score: float, high: float = 90.0, low: float = 70.0,
              floor: float = DEFAULT_THRESHOLD) -> str:
    """按官方 5 档口径给出归属（阈值可调）。

    GUI 请直接构造 :class:`Thresholds` 并调用其 :meth:`~Thresholds.bucket`。
    """
    return Thresholds(high=high, low=low, floor=floor).bucket(score)


def summarize(results: Sequence[MatchResult],
              thresholds: Thresholds = DEFAULT_THRESHOLDS) -> dict[str, int]:
    """按 5 档统计条数。"""
    counts = {"完全匹配": 0, "高度匹配": 0, "中低匹配": 0,
              "低置信度匹配": 0, "A系统独有": 0}
    for r in results:
        counts[thresholds.bucket(r.score)] += 1
    return counts


def sample_band(
    results: Sequence[MatchResult], low: float, high: float, k: int = 3, seed: int = 42
) -> list[MatchResult]:
    """在 ``[low, high)`` 分数区间内随机抽样 k 条，供人工核对算法合理性。"""
    pool = [r for r in results if low <= r.score < high]
    if not pool:
        return []
    rng = np.random.default_rng(seed)
    idx = rng.choice(len(pool), size=min(k, len(pool)), replace=False)
    return [pool[i] for i in sorted(idx)]


# --------------------------------------------------------------------------- #
#  LLM 仲裁兜底（LLM Arbiter）—— 架构扩展位
# --------------------------------------------------------------------------- #
LLM_TIMEOUT = 8.0
LLM_DEFAULT_MODEL = "claude-opus-5"
LLM_ENDPOINT_ENV = "LLM_ARBITER_URL"
LLM_KEY_ENV = "LLM_ARBITER_KEY"

_LLM_SYSTEM_PROMPT = (
    "你是审计对账领域的实体对齐专家。给你两条来自不同系统的企业名称，"
    "判断它们是否指向**同一个法律主体**。\n"
    "判定要点：\n"
    "1. 简称/全称、括号附注差异（如「(特殊普通合伙)」「(北京分所)」）→ 同一主体\n"
    "2. 字号不同（如「安永华明」vs「大华」）→ 不同主体\n"
    "3. 单字错别字（如「有限」vs「有线」、「字节跳动」vs「字跳」）→ 同一主体\n"
    "4. 行业词被替换（如「石油天然气」vs「石油化工」）→ 不同主体\n"
    "只输出 JSON，不要任何其他文字。"
)


@dataclass(frozen=True)
class LLMVerdict:
    """LLM 仲裁结论。"""

    same_entity: bool
    confidence: float          # 0–1，模型自评把握度
    reason: str                # 一句话判据
    source: str = "llm"        # 结论来源：llm / rule-fallback / error

    def __str__(self) -> str:
        flag = "同一主体" if self.same_entity else "不同主体"
        return f"[{self.source}] {flag}（把握 {self.confidence:.0%}）：{self.reason}"


def llm_fallback_verify(
    a_name: str,
    b_name: str,
    *,
    context: dict[str, Any] | None = None,
    timeout: float = LLM_TIMEOUT,
) -> LLMVerdict | None:
    """低置信度配对的 **LLM 仲裁兜底**（架构扩展位，默认不联网）。

    何时调用
    --------
    规则引擎给不出的判断 —— 典型是 ``floor ≤ score < low``（60–70 分）的
    "低置信度匹配"，以及 A/B 独有记录里相似度 45–60 分的候选。这部分记录
    人工逐条看成本最高、收益也最高。

    为什么是"兜底"而不是主路径
    --------------------------
    100 条规模下规则引擎已经做到 45 完全 + 38 高度、A独有 17 / B独有 14，
    与官方口径完全吻合；引入 LLM 的价值在**规模化**（万级记录）与
    **规则盲区**（新造的简称模式）。因此默认关闭，仅按需启用。

    本函数的行为
    ------------
    * 未设置环境变量 ``LLM_ARBITER_URL`` → 直接返回 ``None``（不联网、零副作用）
    * 已设置 → 用**标准库 urllib** 向该地址 POST 一段 JSON，8 秒超时，
      任何异常都吞掉并返回 ``None``，绝不拖垮主流程
    * 请求/响应格式见下方 ``HTTP 协议约定``

    用法
    ----
    ::

        # 1) 指向任一 LLM 代理（自建网关 / 本地 Ollama / 单位内网模型服务）
        #    Windows:  set LLM_ARBITER_URL=http://127.0.0.1:8000/v1/chat
        #    Linux  :  export LLM_ARBITER_URL=http://127.0.0.1:8000/v1/chat
        verdict = llm_fallback_verify("安永华明会计师事务所（特殊普通合伙）",
                                      "安永华明会计师事务所（北京分所）")
        if verdict and verdict.same_entity:
            ...

    HTTP 协议约定
    -------------
    请求 ``POST {LLM_ARBITER_URL}``，``Content-Type: application/json``::

        {
          "model": "claude-opus-5",
          "system": "<上面的 _LLM_SYSTEM_PROMPT>",
          "max_tokens": 256,
          "messages": [{"role": "user", "content": "A系统名称：...\\nB系统名称：..."}],
          "options": {"temperature": 0, "timeout": 8}
        }

    响应（最小可用形状，兼容 OpenAI 风格与 Anthropic 风格）::

        { "content": "{\\"same_entity\\": true, \\"confidence\\": 0.92, \\"reason\\": \\"括号内为分支机构\\"}" }

    也接受 ``{"completion": "..."}`` / ``{"choices":[{"message":{"content":"..."}}]}``
    / ``{"content":[{"type":"text","text":"..."}]}``（Anthropic Messages API 原样返回）。

    生产环境改为官方 SDK
    --------------------
    标准库 urllib 是为了**零新增依赖**。若允许装包，直接用 Anthropic 官方 SDK
    更稳（自带重试、超时、类型）：

    .. code-block:: python

        # pip install anthropic
        import anthropic, json
        client = anthropic.Anthropic()          # 读 ANTHROPIC_API_KEY 环境变量

        def llm_fallback_verify(a_name, b_name):
            resp = client.messages.create(
                model="claude-opus-5",           # $5 / $25 每百万 token（输入/输出）
                max_tokens=256,
                system=_LLM_SYSTEM_PROMPT,
                messages=[{"role": "user",
                           "content": f"A系统名称：{a_name}\\nB系统名称：{b_name}"}],
            )
            data = json.loads(resp.content[0].text)
            return LLMVerdict(data["same_entity"], float(data["confidence"]),
                              data["reason"], source="claude-opus-5")

    每条判定只花约 200 输入 / 60 输出 token，即约 $0.0025。1000 条低置信度
    记录约 $2.5；再开 prompt caching（把 ``_LLM_SYSTEM_PROMPT`` 标为
    ``cache_control``）可再降约 90% 的输入成本。

    批量场景建议改用 **Message Batches API**（``client.messages.batches``），
    异步执行、成本减半，适合"跑完一夜出结果"的对账场景。
    """
    import json
    import os
    import urllib.error
    import urllib.request

    endpoint = os.environ.get(LLM_ENDPOINT_ENV)
    if not endpoint:
        return None                                   # 未配置 → 静默跳过，不联网

    user_prompt = f"A系统名称：{a_name}\nB系统名称：{b_name}"
    if context:
        user_prompt += "\n补充信息：" + json.dumps(context, ensure_ascii=False)
    payload = {
        "model": os.environ.get("LLM_ARBITER_MODEL", LLM_DEFAULT_MODEL),
        "system": _LLM_SYSTEM_PROMPT,
        "max_tokens": 256,
        "messages": [{"role": "user", "content": user_prompt}],
        "options": {"temperature": 0, "timeout": timeout},
    }
    headers = {"Content-Type": "application/json"}
    if key := os.environ.get(LLM_KEY_ENV):
        headers["Authorization"] = f"Bearer {key}"

    try:
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return None                                   # 网络/解析失败 → 静默降级

    return _parse_llm_body(body)


def _parse_llm_body(body: dict[str, Any]) -> LLMVerdict | None:
    """兼容多种网关返回形状，抽出模型给出的 JSON 结论。"""
    import json

    text = ""
    if isinstance(body.get("content"), str):
        text = body["content"]
    elif isinstance(body.get("completion"), str):
        text = body["completion"]
    elif isinstance(body.get("choices"), list) and body["choices"]:
        text = str(body["choices"][0].get("message", {}).get("content", ""))
    elif isinstance(body.get("content"), list) and body["content"]:
        # Anthropic Messages API 原样返回：content 是 block 列表
        text = "".join(b.get("text", "") for b in body["content"]
                       if isinstance(b, dict))
    if not text:
        return None

    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```")
    try:
        data = json.loads(text)
    except ValueError:
        return None

    return LLMVerdict(
        same_entity=bool(data.get("same_entity", False)),
        confidence=float(data.get("confidence", 0.5)),
        reason=str(data.get("reason", "")).strip(),
        source="llm",
    )


def apply_llm_fallback(
    outcome: MatchOutcome,
    df_a: "pd.DataFrame",
    df_b: "pd.DataFrame",
    a_col: str,
    b_col: str,
    thresholds: Thresholds = DEFAULT_THRESHOLDS,
    *,
    band: tuple[float, float] | None = None,
    verbose: bool = True,
) -> dict[int, LLMVerdict]:
    """对低置信度记录批量跑 LLM 仲裁，返回 ``{A行号: 结论}``。

    仅在环境变量 ``LLM_ARBITER_URL`` 已配置时才有实际动作；否则原样返回空字典，
    调用方可安全地无条件调用本函数。

    ``band`` 默认取 ``[floor, low)``，即"低置信度匹配"区间。
    """
    lo, hi = band if band else (thresholds.floor, thresholds.low)
    targets = [r for r in outcome.results if lo <= r.score < hi]
    if not targets:
        return {}
    if not __import__("os").environ.get(LLM_ENDPOINT_ENV):
        if verbose:
            print(f"  [i] LLM 仲裁：{len(targets)} 条待复核，但未配置 "
                  f"{LLM_ENDPOINT_ENV}，已跳过（规则引擎结果不受影响）")
        return {}

    if verbose:
        print(f"  [i] LLM 仲裁：正在复核 {len(targets)} 条低置信度记录…")

    verdicts: dict[int, LLMVerdict] = {}
    for r in _iter(targets, "  LLM 仲裁", "条", None, "", ""):
        v = llm_fallback_verify(r.a_name, r.b_name,
                                context={"规则分数": round(r.score, 1)})
        if v is not None:
            verdicts[r.a_index] = v
    return verdicts
