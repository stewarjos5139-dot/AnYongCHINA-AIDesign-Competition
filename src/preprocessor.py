"""文本清洗模块 —— 生成辅助列，**绝不改动原始名称列**。

清洗流水线（顺序不可调换）
--------------------------
1. 剔隐形字符：TAB/换行、NBSP、全角空格、零宽字符、BOM、C0 控制符
2. 全角 → 半角：括号 ``（）``→``()``，以及全部全角 ASCII(U+FF01–U+FF5E) 与常见中文标点
3. 去空格：删除串内与首尾**所有**空格（中文企业名去空格最稳妥）
4. 统一大小写：英文字母转大写

输出两列（均为**新增**，原始列不动）：
* ``<名称列>_清洗后``        —— 主比对列（保留括号等标点）
* ``<名称列>_清洗后_去标点`` —— 副比对列（剥离全部标点，用于"核心字号"比对）

实现备注：所有不可见字符一律以码点区间声明后**程序化拼装**正则，
源码中不出现任何真实控制字符，避免编辑器/复制粘贴环节损坏。
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

# --------------------------------------------------------------------------- #
#  规则表
# --------------------------------------------------------------------------- #
# 全角 ASCII → 半角 ASCII：U+FF01–U+FF5E 与 U+0021–U+007E 相差固定 0xFEE0。
# 该区间天然覆盖全角括号 （ U+FF08 → ( 、 ） U+FF09 → ) 以及全角字母/数字。
_FULLWIDTH_ASCII: dict[int, int] = {c: c - 0xFEE0 for c in range(0xFF01, 0xFF5F)}

# 全角 ASCII 区间之外、但企业名中常见的中文标点 → 半角等价物
_EXTRA_PUNCT: dict[int, int] = {
    0x3001: 0x2C,   # 、 → ,
    0x3002: 0x2E,   # 。 → .
    0x3008: 0x3C,   # 〈 → <
    0x3009: 0x3E,   # 〉 → >
    0x300A: 0x3C,   # 《 → <
    0x300B: 0x3E,   # 》 → >
    0x3010: 0x5B,   # 【 → [
    0x3011: 0x5D,   # 】 → ]
    0x2018: 0x27,   # ‘ → '
    0x2019: 0x27,   # ’ → '
    0x201C: 0x22,   # “ → "
    0x201D: 0x22,   # ” → "
    0x2013: 0x2D,   # – → -
    0x2014: 0x2D,   # — → -
    0x2015: 0x2D,   # ― → -
    0x2212: 0x2D,   # − → -
    0x00B7: 0x2E,   # · → .
}

_TRANSLATE_TABLE: dict[int, int] = {**_FULLWIDTH_ASCII, **_EXTRA_PUNCT}

# 不可见 / 格式控制字符的码点区间（含空格类、零宽类、控制类）
_INVISIBLE_RANGES: tuple[tuple[int, int], ...] = (
    (0x0009, 0x000D),   # TAB / LF / VT / FF / CR
    (0x0000, 0x0008),   # C0 起始段
    (0x000E, 0x001F),   # C0 剩余段
    (0x007F, 0x007F),   # DEL
    (0x00A0, 0x00A0),   # NBSP 不换行空格
    (0x1680, 0x1680),   # OGHAM SPACE MARK
    (0x2000, 0x200F),   # EN QUAD … RLM（含 ZWSP / ZWNJ / ZWJ / LRM / RLM）
    (0x2028, 0x2029),   # LINE / PARAGRAPH SEPARATOR
    (0x202F, 0x202F),   # NARROW NBSP
    (0x205F, 0x205F),   # MEDIUM MATHEMATICAL SPACE
    (0x2060, 0x2060),   # WORD JOINER
    (0x3000, 0x3000),   # IDEOGRAPHIC SPACE 全角空格
    (0xFEFF, 0xFEFF),   # BOM / ZWNBSP
)


def _build_char_class(ranges: tuple[tuple[int, int], ...]) -> str:
    """把码点区间拼成 ``[\\u0009-\\u000d...]`` 形式的字符类文本。"""
    parts = [
        f"\\u{lo:04x}" if lo == hi else f"\\u{lo:04x}-\\u{hi:04x}"
        for lo, hi in ranges
    ]
    return "[" + "".join(parts) + "]"


_INVISIBLE_RE = re.compile(_build_char_class(_INVISIBLE_RANGES))

# 需要剥离的标点（用于"核心字号"副列）
_PUNCT_CHARS = (
    "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
    "、。〈〉《》「」『』【】〔〕（）［］｛｝，．；：？！…—–―·～￥"
)
_PUNCT_RE = re.compile("[" + re.escape(_PUNCT_CHARS) + "]")

# 常见公司后缀 —— 供"缺后缀"类差异判定复用（按长度降序，先匹配长后缀）
COMPANY_SUFFIXES: tuple[str, ...] = (
    "股份有限公司", "有限责任公司", "集团有限公司", "股份公司",
    "有限公司", "有限公司", "集团", "公司", "事务所", "合伙", "中心", "厂", "店",
)

# 主比对列 / 副比对列 后缀
CLEAN_SUFFIX = "_清洗后"
CORE_SUFFIX = "_清洗后_去标点"


# --------------------------------------------------------------------------- #
#  单条字符串清洗
# --------------------------------------------------------------------------- #
def strip_invisible(text: str) -> str:
    """剔除不可见字符（各类空格 / 零宽 / 控制符）。"""
    return _INVISIBLE_RE.sub("", text)


def has_invisible(text: str) -> bool:
    """是否含不可见字符。"""
    return _INVISIBLE_RE.search(text) is not None


def to_halfwidth(text: str) -> str:
    """全角 → 半角（含括号、字母、数字、常见中文标点）。"""
    return text.translate(_TRANSLATE_TABLE)


def upper_case(text: str) -> str:
    """英文字母统一大写。"""
    return text.upper()


def remove_spaces(text: str) -> str:
    """删除串内与首尾全部空格。"""
    return strip_invisible(text).replace(" ", "")


def clean_name(value: object) -> str:
    """完整清洗流水线：隐形字符 → 全半角 → 去空格 → 大写。

    ``None`` / ``NaN`` → 空串。原始字符串不做任何回写。
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value)
    text = strip_invisible(text)   # 1 隐形字符
    text = to_halfwidth(text)      # 2 全角 → 半角（含括号）
    text = text.replace(" ", "")   # 3 去空格
    return text.upper()            # 4 统一大写


def strip_punct(text: str) -> str:
    """剥离全部标点，得到"核心字号"字符串（供副比对列使用）。"""
    return _PUNCT_RE.sub("", text)


_PAREN_RE = re.compile(r"\([^()]*\)")


def strip_parens(text: str) -> str:
    """删除所有成对括号及其内容（``(集团)``、``(杭州)``、``(特殊普通合伙)``）。

    中文企业名里的括号内容基本是**附注性**的（分支机构 / 组织形式 / 集团标记），
    去掉后再比对一次，可以识别出"正文相同、只有括号附注不同"的同一主体，
    也能避免尾部括号破坏前后缀对齐而放走误匹配。
    """
    return _PAREN_RE.sub("", text)


def core_name(value: object) -> str:
    """清洗 + 去标点。"""
    return strip_punct(clean_name(value))


# --------------------------------------------------------------------------- #
#  字号（distinctive core）分析 —— 抵抗"长通用尾巴"造成的误匹配
# --------------------------------------------------------------------------- #
def common_affix_len(s1: str, s2: str) -> tuple[int, int]:
    """返回两串的 **(公共前缀长度, 公共后缀长度)**，两段互不重叠。"""
    n = min(len(s1), len(s2))
    prefix = 0
    while prefix < n and s1[prefix] == s2[prefix]:
        prefix += 1
    suffix = 0
    while suffix < n - prefix and s1[len(s1) - 1 - suffix] == s2[len(s2) - 1 - suffix]:
        suffix += 1
    return prefix, suffix


def distinctive_core(s1: str, s2: str) -> tuple[str, str]:
    """剥离公共前后缀后剩下的"A 特征片段 / B 特征片段"。

    中文企业名结构为 ``[地区][字号][行业][组织类型]``。两条名称共享
    ``事务所(特殊普通合伙)`` / ``股份有限公司`` 这类长通用尾巴时，
    ``ratio`` 会被尾巴推高，但真正决定是否同一主体的**字号**恰恰在差异片段里。

    >>> distinctive_core("安永华明会计师事务所(特殊普通合伙)", "大华会计师事务所(特殊普通合伙)")
    ('安永华明会计', '大华')
    >>> distinctive_core("中国建设银行股份有限公司", "中国银行股份有限公司")
    ('建设', '')
    >>> distinctive_core("中国石油天然气股份有限公司", "中国石油化工股份有限公司")
    ('天然气', '化工')
    """
    if not s1 or not s2 or s1 == s2:
        return "", ""
    prefix, suffix = common_affix_len(s1, s2)
    return s1[prefix: len(s1) - suffix], s2[prefix: len(s2) - suffix]


# 差异片段惩罚系数
PENALTY_NONE = 1.00        # 错别字级差异，不惩罚
PENALTY_MILD = 0.90        # 差异片段较相似
PENALTY_INSERT_2 = 0.55    # 纯插入/删除 2 字（如 中国[建设]银行）—— 字号不同
PENALTY_HEAVY = 0.50       # 两侧字号都不同（如 安永华明 vs 大华）
PENALTY_LONG_INSERT = 0.40 # 纯插入/删除 ≥3 字（如 平安保险(集团)[股份有限公司]）

TYPO_CORE_LEN = 1          # 差异片段 ≤1 字 → 单字错别字 / 漏字，不惩罚
SIM_KEEP = 70.0            # 差异片段相似度 ≥70 → 视为同一字号
SIM_MILD = 45.0            # ≥45 → 轻度惩罚


def core_divergence_penalty(s1: str, s2: str) -> float:
    """字号差异惩罚系数（0–1），乘到基础相似度上。

    取两个视角中**较严**的一个：

    * 视角 A —— 原样比对（``客户名称_清洗后``）
    * 视角 B —— 剥掉所有括号附注后比对（``strip_parens``）

    视角 B 专门解决"尾部/中部括号附注破坏前后缀对齐"导致的漏判，例如::

        安永华明会计师事务所(特殊普通合伙)  vs  华兴会计师事务所(特殊普通合伙)(杭州)
        视角A：差异片段相似度 66.7 → 仅轻度惩罚（放行 81 分，误匹配）
        视角B：去括号后 安永华明会计师事务所 vs 华兴会计师事务所(特殊普通合伙)
               差异片段 安永华明 vs 华兴会 相似度 33.3 → 重度惩罚（降至 45 分）

    单条规则（按优先级）：

    1. 两串完全相同 → 1.0
    2. 一侧差异片段为空 = **纯插入 / 删除**：≤1 字 → 1.0；2 字 → 0.55；≥3 字 → 0.40
    3. 两侧都非空但最短一侧 ≤1 字 → **单字错别字 / 漏字** → 1.0
    4. 其余按差异片段相似度给 1.0 / 0.90 / 0.50

    典型效果（均为中文企业名的真实陷阱）：

    ==============================================================  ======
    名称对                                                          系数
    ==============================================================  ======
    北京字节跳动网络技术有限公司 / 北京字跳网络技术有限公司            1.00
    云南白药集团股份有限公司 / 云南白药集团股份有线公司                1.00
    中国建设银行股份有限公司 / 中国银行股份有限公司                    0.55
    安永华明会计师事务所(特殊普通合伙) / 大华会计师事务所(特殊普通合伙)   0.50
    中国石油天然气股份有限公司 / 中国石油化工股份有限公司              0.50
    中国平安保险(集团)股份有限公司 / 中国平安保险(集团)                  0.40
    ==============================================================  ======

    .. note:: 末行为纯后缀缺失，这里先压低；随后匹配引擎的"包含关系抬分"
       会把它重新抬回高度匹配档（简称 / 全称）。
    """
    return min(_penalty_one_view(s1, s2), _penalty_one_view(strip_parens(s1),
                                                            strip_parens(s2)))


def _penalty_one_view(s1: str, s2: str) -> float:
    """单视角（不做去括号处理）的字号差异惩罚。"""
    if not s1 or not s2 or s1 == s2:
        return PENALTY_NONE

    core_a, core_b = distinctive_core(s1, s2)
    if not core_a and not core_b:
        return PENALTY_NONE

    longest = max(len(core_a), len(core_b))

    # 一侧为空 = 纯插入 / 删除（如 平安保险(集团)[股份有限公司]）
    if not core_a or not core_b:
        if longest <= TYPO_CORE_LEN:
            return PENALTY_NONE
        return PENALTY_INSERT_2 if longest == 2 else PENALTY_LONG_INSERT

    # 两侧都非空 —— 区分"错别字 / 漏字"与"字号真的不同"
    # （a）两侧都只剩 1 字 → 单字替换：云[南]白药 vs 云[蓝]白药
    # （b）短片段（≤2 字）是长片段的子串 → 漏字：北京字[节]跳动 vs 北京字跳
    core_short, core_long = (
        (core_a, core_b) if len(core_a) <= len(core_b) else (core_b, core_a)
    )
    if len(core_a) <= TYPO_CORE_LEN and len(core_b) <= TYPO_CORE_LEN:
        return PENALTY_NONE
    if len(core_short) <= 2 and core_short in core_long:
        return PENALTY_NONE

    # 其余：字号真的不同（如 建设银 vs 筑、天然气 vs 化工、安永华明 vs 大华）
    sim = _ratio(core_a, core_b)
    if sim >= SIM_KEEP:
        return PENALTY_NONE
    if sim >= SIM_MILD:
        return PENALTY_MILD
    return PENALTY_HEAVY


def _ratio(s1: str, s2: str) -> float:
    """惰性导入 rapidfuzz，避免模块级硬依赖（缺失时退化到 difflib）。"""
    try:
        from rapidfuzz import fuzz

        return float(fuzz.ratio(s1, s2))
    except ImportError:              # pragma: no cover
        from difflib import SequenceMatcher

        return SequenceMatcher(None, s1, s2).ratio() * 100.0


# --------------------------------------------------------------------------- #
#  DataFrame 级接口
# --------------------------------------------------------------------------- #
def clean_series(series: pd.Series) -> pd.Series:
    """对整个 Series 做清洗，返回新的 ``string`` Series（不改原列）。"""
    return series.map(clean_name).astype("string")


def add_clean_columns(
    df: pd.DataFrame, name_col: str, clean_col: str | None = None
) -> pd.DataFrame:
    """**原地新增**辅助列：``<name_col>_清洗后`` 与 ``<name_col>_清洗后_去标点``。

    原始 ``name_col`` 保持原样，供最终报表原样输出。
    """
    clean_col = clean_col or f"{name_col}{CLEAN_SUFFIX}"
    core_col = f"{name_col}{CORE_SUFFIX}"

    df[clean_col] = clean_series(df[name_col])
    df[core_col] = df[clean_col].map(strip_punct).astype("string")
    return df


def diff_profile(raw_a: str, raw_b: str) -> list[str]:
    """对比两条名称，产出人类可读的差异原因标签（用于结果表「备注」列）。

    分两层判定，互不重复：

    * **清洗层**：原始串里存在哪些脏数据（括号全半角 / 空格 / 隐形字符 / 大小写）
    * **主体层**：清洗后两条名称是什么关系（包含 / 字号不同 / 括号内容不同 / 单字差异）
    """
    raw_a, raw_b = str(raw_a), str(raw_b)
    reasons: list[str] = []

    if raw_a == raw_b:
        return ["名称完全一致"]

    # ---------------- 清洗层 ----------------
    if (
        raw_a.replace("（", "(").replace("）", ")")
        == raw_b.replace("（", "(").replace("）", ")")
    ):
        reasons.append("括号全/半角格式不同")
    if has_invisible(raw_a) or has_invisible(raw_b):
        reasons.append("含不可见字符")
    elif re.search(r"\s", raw_a) or re.search(r"\s", raw_b):
        reasons.append("存在多余空格")
    if raw_a.upper() != raw_a or raw_b.upper() != raw_b:
        reasons.append("含英文大小写差异")

    clean_a, clean_b = clean_name(raw_a), clean_name(raw_b)
    if clean_a == clean_b:
        return reasons or ["清洗后完全一致"]

    # ---------------- 主体层 ----------------
    core_a, core_b = distinctive_core(clean_a, clean_b)
    penalty = core_divergence_penalty(clean_a, clean_b)

    if clean_a in clean_b:
        extra = clean_b.replace(clean_a, "", 1)
        reasons.append(f"B在A基础上附加「{extra}」")
    elif clean_b in clean_a:
        extra = clean_a.replace(clean_b, "", 1)
        reasons.append(f"A在B基础上附加「{extra}」")
    else:
        # 非包含关系：先看字号是否真的不同
        if penalty <= PENALTY_INSERT_2:
            reasons.append(f"字号不同「{core_a or '（无）'}」vs「{core_b or '（无）'}」")
        elif penalty < PENALTY_NONE:
            reasons.append(f"字号近似「{core_a}」vs「{core_b}」")

        if re.findall(r"\(([^)]*)\)", clean_a) != re.findall(r"\(([^)]*)\)", clean_b):
            reasons.append("括号内附加信息不同")

        if not reasons:
            prefix, _ = common_affix_len(clean_a, clean_b)
            reasons.append(
                f"第{prefix + 1}字起不同「{core_a or '（无）'}」vs「{core_b or '（无）'}」"
            )

    if abs(len(clean_a) - len(clean_b)) >= 2:
        reasons.append(f"长度 {len(clean_a)}字 vs {len(clean_b)}字")

    return reasons or ["字符存在差异"]


# --------------------------------------------------------------------------- #
#  Unicode 诊断
# --------------------------------------------------------------------------- #
def invisible_hits(text: str) -> list[tuple[int, str]]:
    """返回 ``(位置, 字符repr)`` 列表，用于控制台诊断打印。"""
    return [
        (i, repr(ch))
        for i, ch in enumerate(text)
        if has_invisible(ch) or unicodedata.category(ch) in ("Cf", "Cc")
    ]
