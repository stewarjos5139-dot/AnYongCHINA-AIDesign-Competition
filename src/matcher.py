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

MIN_CONTAINMENT_LEN = 4        # 短串短于此长度不做包含关系抬分（防"华为"类过度匹配）
PARTIAL_TRIGGER = 95.0         # partial_ratio 超过此值才进入包含关系候选
ANCHOR_PENALTY = 15.0          # 前缀/后缀锚定：100 - (1-覆盖率)*15
MIDDLE_PENALTY = 28.0          # 中间截取：100 - (1-覆盖率)*28

DEFAULT_THRESHOLD = 60.0       # 低于此分视为"无匹配"，供第三阶段判定独有记录

# 各阶段的进度权重（占本模块 0–100% 的区间），供 CLI / GUI 共用
STAGE_WEIGHTS: dict[str, tuple[float, float]] = {
    "matrix": (0.0, 45.0),       # 4 个 scorer 的相似度矩阵
    "penalty": (45.0, 25.0),     # 字号差异校验
    "boost": (70.0, 10.0),       # 包含关系抬分
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
    conflict: bool = False      # 该 B 记录是否被其他 A 记录抢为首选
    conflict_count: int = 1

    @property
    def is_matched(self) -> bool:
        return self.b_index is not None and self.score >= DEFAULT_THRESHOLD


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
) -> tuple[np.ndarray, np.ndarray]:
    """计算 A×B 全量相似度矩阵。

    返回 ``(score, detail_stack)``：
    * ``score``        —— ``(len(a), len(b))`` 最终加权分 0–100
    * ``detail_stack`` —— ``(len(a), len(b), 4)`` 各算法原始分，便于结果表追溯
    """
    a_clean = list(a_clean)
    b_clean = list(b_clean)
    n_a, n_b = len(a_clean), len(b_clean)
    if n_a == 0 or n_b == 0:
        return np.zeros((n_a, n_b)), np.zeros((n_a, n_b, 4))

    a_core = list(a_core) if a_core is not None else [pp.strip_punct(x) for x in a_clean]
    b_core = list(b_core) if b_core is not None else [pp.strip_punct(x) for x in b_clean]

    # rapidfuzz 的 cdist 一次只接受单个 scorer，故逐个调用；
    # 每个 scorer 内部由 C++ 多线程并行（workers=-1），n×m 全量矩阵毫秒级完成。
    mats: list[np.ndarray] = []
    for k, (tag, (scorer, left, right)) in enumerate({
        "ratio": (fuzz.ratio, a_clean, b_clean),
        "partial_ratio": (fuzz.partial_ratio, a_clean, b_clean),
        "WRatio": (fuzz.WRatio, a_clean, b_clean),
        "core_ratio": (fuzz.ratio, a_core, b_core),
    }.items(), start=1):
        t = time.perf_counter()
        mats.append(
            np.asarray(
                process.cdist(left, right, scorer=scorer, workers=workers),
                dtype=np.float64,
            )
        )
        if verbose:
            print(f"    · {tag:<14} {n_a}×{n_b} 矩阵  {time.perf_counter() - t:.4f}s")
        _emit(progress, "matrix", k, 4, f"计算相似度矩阵 {k}/4（{tag}）")

    detail = np.stack(mats, axis=2)                         # (nA, nB, 4)

    score = (
        WEIGHTS["ratio"] * detail[:, :, 0]
        + WEIGHTS["partial_ratio"] * detail[:, :, 1]
        + WEIGHTS["WRatio"] * detail[:, :, 2]
        + WEIGHTS["core_ratio"] * detail[:, :, 3]
    )

    # ---- 修正 1：字号（distinctive core）差异惩罚 ----
    # 必须先于"包含关系抬分"执行：简称/全称 会在此被压低，随后被包含了关系重新抬回。
    score = apply_core_penalty(score, a_clean, b_clean, verbose=verbose,
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

    return np.clip(score, 0.0, 100.0), detail


def apply_core_penalty(
    score: np.ndarray,
    a_clean: Sequence[str],
    b_clean: Sequence[str],
    trigger: float = 55.0,
    verbose: bool = True,
    progress: ProgressFn | None = None,
) -> np.ndarray:
    """对"共享长通用尾巴但字号不同"的配对打折。

    仅处理基础分 ``≥ trigger`` 的配对 —— 分数本就很低的配对无论如何都进不了匹配档，
    无需消耗字符串分析开销，保证大数据量下的效率。
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
) -> MatchOutcome:
    """A × B 全量比对，返回每条 A 记录的最佳匹配。

    ``tie_break=True`` 时，对**名称分数完全相同**的多条 B 候选（典型场景：B 表里有两条
    同名流水，对应两笔不同交易），用「金额 + 交易日期」的贴合度挑选最可能的那笔，
    使结果不再依赖 B 表的行序。名称分数不同时该裁决绝不介入，金额不会影响匹配结论。

    ``thresholds`` 仅影响 ``b_taken``（哪些 B 记录算"被认领"）；分档标注在
    :meth:`Thresholds.bucket` / :func:`summarize` 中进行。
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
        progress=progress,
    )
    n_a, n_b = score.shape
    if n_a == 0 or n_b == 0:
        return MatchOutcome(results=[], score_matrix=score, b_taken={},
                            stats={"n_a": n_a, "n_b": n_b})

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

    # ---- B 记录争抢情况（只统计达到匹配阈值的 A 记录） ----
    b_taken: dict[int, list[int]] = {}
    for i, j in enumerate(best_idx):
        if best_score[i] >= thresholds.floor:
            b_taken.setdefault(int(j), []).append(i)
    for j in b_taken:
        b_taken[j].sort(key=lambda ai: -score[ai, j])

    results: list[MatchResult] = []
    for i in _iter(range(n_a), "  选取最佳匹配", "行", progress, "select",
                   "选取最佳匹配"):
        j = int(best_idx[i])
        rivals = b_taken.get(j, [])
        results.append(
            MatchResult(
                a_index=i,
                a_name=a_raw[i],
                a_clean=a_clean[i],
                b_index=j,
                b_name=b_raw[j],
                b_clean=b_clean[j],
                score=float(best_score[i]),
                detail={
                    "ratio": float(detail[i, j, 0]),
                    "partial_ratio": float(detail[i, j, 1]),
                    "WRatio": float(detail[i, j, 2]),
                    "core_ratio": float(detail[i, j, 3]),
                },
                runner_up_index=int(runner_idx[i]) if runner_idx[i] >= 0 else None,
                runner_up_name=b_raw[int(runner_idx[i])] if runner_idx[i] >= 0 else "",
                runner_up_score=float(runner_score[i]),
                mutual_best=bool(b_best_a[j] == i),
                conflict=len(rivals) > 1,
                conflict_count=max(1, len(rivals)),
            )
        )

    stats = {
        "n_a": n_a,
        "n_b": n_b,
        "pairs": int(n_a * n_b),
        "exact_100": int(np.sum(np.round(best_score, 6) >= 100.0)),
        "conflict_targets": sum(1 for v in b_taken.values() if len(v) > 1),
        "conflict_rows": sum(len(v) - 1 for v in b_taken.values() if len(v) > 1),
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
