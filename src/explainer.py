"""匹配结果可解释性模块（Explainable AI 标签）。

把"为什么这两条被判为同一主体 / 为什么不是"翻译成审计师能直接看懂的规则标签，
写进成果表 Sheet1 的「备注」列。

标签分两层
----------
1. **清洗层** —— 原始串上发生了什么（括号全半角、空格、隐形字符、大小写）
2. **主体层** —— 清洗后两条名称是什么关系（包含 / 字号差异 / 错别字 / 括号内容）

第三层「算法层」由匹配引擎的运行时信息补充（非互为最优、目标争抢、得分偏低）。

标签词表见 :data:`TAG_VOCAB`，全部为**受控词表**，不会出现自由文本标签，
便于审计师在 Excel 里按标签筛选。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable

from . import preprocessor as pp

# --------------------------------------------------------------------------- #
#  受控标签词表
# --------------------------------------------------------------------------- #
TAG_VOCAB: tuple[str, ...] = (
    # --- 清洗层 ---
    "名称完全一致",
    "清洗后一致",
    "括号全/半角差异",
    "多余空格",
    "不可见字符",
    "英文大小写差异",
    # --- 主体层：包含关系 ---
    "包含关系（丢弃公司后缀）",
    "包含关系（附加地区）",
    "包含关系（附加分支机构）",
    "包含关系（附加组织形式）",
    "包含关系（简称 vs 全称）",
    "字号一致",
    # --- 主体层：差异 ---
    "疑似错别字",
    "字号不同",
    "括号内信息不同",
    "含中英文字符",
    "数字不同",
    "长度差异",
    # --- 算法层 ---
    "非互为最优",
    "目标争抢",
    "低置信度（建议逐条核实）",
)

# 地区词（用于判定"附加地区"）
REGION_WORDS: tuple[str, ...] = (
    "中国", "中华", "北京", "上海", "天津", "重庆", "广州", "深圳", "广东", "江苏",
    "浙江", "山东", "河南", "河北", "四川", "湖北", "湖南", "福建", "安徽", "陕西",
    "辽宁", "江西", "云南", "广西", "山西", "内蒙古", "新疆", "贵州", "甘肃", "海南",
    "宁夏", "青海", "西藏", "吉林", "黑龙江", "香港", "澳门", "台湾", "苏州", "杭州",
    "南京", "武汉", "成都", "西安", "青岛", "大连", "宁波", "厦门", "无锡", "佛山",
    "东莞", "郑州", "长沙", "合肥", "福州", "济南", "沈阳", "哈尔滨", "乌鲁木齐",
)

# 分支机构词
BRANCH_WORDS: tuple[str, ...] = (
    "分公司", "分行", "分所", "办事处", "营业部", "支公司", "分厂", "分部", "分局",
    "分店", "连锁店", "门店",
)

# 组织形式 / 集团标记词
ORG_WORDS: tuple[str, ...] = (
    "集团", "特殊普通合伙", "有限合伙", "普通合伙", "合伙企业", "控股",
    "股份有限公司", "有限责任公司", "有限公司", "公司", "事务所", "研究院", "研究所",
    "中心", "厂", "商行", "银行", "保险", "证券", "基金", "信托",
)

_LATIN_RE = re.compile(r"[A-Za-z]")
_DIGIT_RE = re.compile(r"\d")


# --------------------------------------------------------------------------- #
#  结构
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Reason:
    """一条可解释性标签。``detail`` 为可选的补充说明（不参与筛选）。"""

    tag: str
    detail: str = ""

    def __str__(self) -> str:
        return f"{self.tag}（{self.detail}）" if self.detail else self.tag


# --------------------------------------------------------------------------- #
#  辅助判定
# --------------------------------------------------------------------------- #
def _classify_extra(extra: str) -> str:
    """把"两条名称的差集片段"归类到具体的包含关系标签。"""
    if not extra:
        return "字号一致"
    if any(w in extra for w in BRANCH_WORDS):
        return "包含关系（附加分支机构）"
    if any(w in extra for w in ORG_WORDS):
        return "包含关系（丢弃公司后缀）"
    if any(w in extra for w in REGION_WORDS):
        return "包含关系（附加地区）"
    # 纯地区短片段（如「(广州)」去掉括号后只剩两个地区字）
    if any(extra == w or extra.endswith(w) for w in REGION_WORDS):
        return "包含关系（附加地区）"
    return "包含关系（简称 vs 全称）"


def _edit_distance_one(s1: str, s2: str) -> bool:
    """两串是否只差一个字符（插入 / 删除 / 替换）。"""
    if abs(len(s1) - len(s2)) > 1:
        return False
    if len(s1) == len(s2):
        return sum(a != b for a, b in zip(s1, s2)) == 1
    short, long_ = (s1, s2) if len(s1) < len(s2) else (s2, s1)
    for i in range(len(long_)):
        if long_[:i] + long_[i + 1:] == short:
            return True
    return False


# --------------------------------------------------------------------------- #
#  主入口
# --------------------------------------------------------------------------- #
def explain(
    a_name: Any,
    b_name: Any,
    result: Any | None = None,
    thresholds: Any | None = None,
) -> list[Reason]:
    """产出一组受控标签，解释这条匹配的成因。

    参数
    ----
    a_name / b_name : 两侧的**原始**名称
    result          : 可选，``matcher.MatchResult``；提供时追加算法层标签
    thresholds      : 可选，``matcher.Thresholds``；用于判定"低置信度"
    """
    raw_a, raw_b = str(a_name), str(b_name)
    clean_a, clean_b = pp.clean_name(raw_a), pp.clean_name(raw_b)
    reasons: list[Reason] = []

    # ---------------- 第一层：清洗 ----------------
    if raw_a == raw_b:
        reasons.append(Reason("名称完全一致"))
        return _with_runtime(reasons, result, thresholds)

    _append_cleaning_tags(reasons, raw_a, raw_b)

    if clean_a == clean_b:
        reasons.append(Reason("清洗后一致"))
        return _with_runtime(reasons, result, thresholds)

    # ---------------- 第二层：主体关系 ----------------
    core_a, core_b = pp.distinctive_core(clean_a, clean_b)
    penalty = pp.core_divergence_penalty(clean_a, clean_b)
    contained = False

    if clean_a in clean_b:
        contained = True
        extra = clean_b.replace(clean_a, "", 1)
        reasons.append(Reason(_classify_extra(extra), f"B多出「{extra}」"))
    elif clean_b in clean_a:
        contained = True
        extra = clean_a.replace(clean_b, "", 1)
        reasons.append(Reason(_classify_extra(extra), f"A多出「{extra}」"))
    elif core_a or core_b:
        # 非包含关系 —— 按"差异片段的形态"区分错别字与真正不同的字号
        if not core_a or not core_b:
            # 一侧为空 = 纯插入/删除：1 字算漏字，≥2 字属于字号不同
            tag = "疑似错别字" if max(len(core_a), len(core_b)) <= 1 else "字号不同"
        elif _edit_distance_one(core_a, core_b) or min(len(core_a), len(core_b)) <= 1:
            tag = "疑似错别字"
        else:
            tag = "字号不同"
        reasons.append(
            Reason(tag, f"「{core_a or '（无）'}」vs「{core_b or '（无）'}」")
        )
        if re.findall(r"\(([^)]*)\)", clean_a) != re.findall(r"\(([^)]*)\)", clean_b):
            reasons.append(Reason("括号内信息不同"))

    # 数字差异
    if _DIGIT_RE.search(clean_a) or _DIGIT_RE.search(clean_b):
        if re.findall(r"\d+", clean_a) != re.findall(r"\d+", clean_b):
            reasons.append(Reason("数字不同"))

    # 长度差异（包含关系本身已隐含长度差，不重复标注）
    if not contained and abs(len(clean_a) - len(clean_b)) >= 4:
        reasons.append(Reason("长度差异", f"{len(clean_a)}字 vs {len(clean_b)}字"))

    return _with_runtime(reasons, result, thresholds)


def _append_cleaning_tags(reasons: list[Reason], raw_a: str, raw_b: str) -> None:
    """清洗层的标签（顺序固定，便于在 Excel 里按标签筛选）。"""
    if (
        raw_a.replace("（", "(").replace("）", ")")
        == raw_b.replace("（", "(").replace("）", ")")
    ):
        reasons.append(Reason("括号全/半角差异"))
    if pp.has_invisible(raw_a) or pp.has_invisible(raw_b):
        reasons.append(Reason("不可见字符"))
    elif re.search(r"\s", raw_a) or re.search(r"\s", raw_b):
        reasons.append(Reason("多余空格"))
    if raw_a.upper() != raw_a or raw_b.upper() != raw_b:
        reasons.append(Reason("英文大小写差异"))
    # 中英文混杂：一侧含拉丁字母而另一侧不含
    latin_a, latin_b = (
        bool(_LATIN_RE.search(pp.clean_name(raw_a))),
        bool(_LATIN_RE.search(pp.clean_name(raw_b))),
    )
    if latin_a != latin_b:
        reasons.append(Reason("含中英文字符", "一侧含英文，另一侧不含"))


def _with_runtime(reasons: list[Reason], result: Any | None,
                  thresholds: Any | None) -> list[Reason]:
    """追加算法层标签。"""
    if result is None:
        return reasons
    if getattr(result, "conflict", False):
        reasons.append(Reason("目标争抢", f"被{result.conflict_count}条A记录同时选中"))
    elif not getattr(result, "mutual_best", True):
        reasons.append(Reason("非互为最优", "建议复核"))
    # 仅对真正落在"低置信度匹配"档（floor ≤ score < low）的记录标注，
    # 低于 floor 的属于「A系统独有」，不进 Sheet1，标在这里会误导。
    if thresholds is not None:
        score = getattr(result, "score", 100.0)
        lo, hi = getattr(thresholds, "floor", 60.0), getattr(thresholds, "low", 70.0)
        if lo <= score < hi:
            reasons.append(Reason("低置信度（建议逐条核实）"))
    return reasons


def explain_text(
    a_name: Any,
    b_name: Any,
    result: Any | None = None,
    thresholds: Any | None = None,
    sep: str = "；",
    limit: int = 5,
) -> str:
    """标签串（供 Excel「备注」列）。最多保留 ``limit`` 条，避免备注过长。"""
    items = explain(a_name, b_name, result, thresholds)
    if not items:
        return ""
    text = sep.join(str(r) for r in items[:limit])
    if len(items) > limit:
        text += f"{sep}…（共{len(items)}项）"
    return text


def tag_names(reasons: Iterable[Reason]) -> list[str]:
    """只取标签名，便于统计各标签出现频次。"""
    return [r.tag for r in reasons]


def tag_histogram(rows: Iterable[tuple[Any, Any]]) -> dict[str, int]:
    """统计一批名称对的标签分布（供控制台/文档展示可解释性覆盖率）。"""
    hist: dict[str, int] = {}
    for a, b in rows:
        for tag in tag_names(explain(a, b)):
            hist[tag] = hist.get(tag, 0) + 1
    return dict(sorted(hist.items(), key=lambda kv: -kv[1]))
