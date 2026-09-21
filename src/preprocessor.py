"""文本清洗模块 —— 生成辅助列，**绝不改动原始名称列**。

清洗流水线（顺序不可调换）
--------------------------
1. 剔隐形字符：C0/C1 控制符、各类空格、`Cf` 格式控制符全类、TAG 字符等
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

# 不可见 / 格式控制字符的码点区间。
#
# 覆盖口径是**按 Unicode 类别穷举**，而不是凭印象挑几个最常见的：
#
# * ``Cc`` 控制符（C0 + C1）
# * ``Zs`` / ``Zl`` / ``Zp`` 各类空格分隔符
# * ``Cf`` 格式控制符 —— **全部**（软连字符、双向控制、TAG 字符…）
# * 少数类别不属于 Cf 但渲染为空白、且在企业名里纯属噪声的字符
#   （谚文填充符、变体选择符）
#
# 序号不可乱动：区间之间**不允许重叠**，也不允许跨 U+FFFF 边界
# （见 :func:`_build_char_class` 的转义宽度选择）。
_INVISIBLE_RANGES: tuple[tuple[int, int], ...] = (
    # ---- Cc：C0 / C1 控制符 ----
    (0x0000, 0x0008),   # NUL … BS
    (0x0009, 0x000D),   # TAB / LF / VT / FF / CR
    (0x000E, 0x001F),   # C0 剩余段
    (0x007F, 0x009F),   # DEL + C1 整段
    # ---- Zs / Zl / Zp：各类空格分隔符 ----
    (0x00A0, 0x00A0),   # NBSP 不换行空格
    (0x1680, 0x1680),   # OGHAM SPACE MARK
    (0x2000, 0x200A),   # EN QUAD … HAIR SPACE
    (0x2028, 0x2029),   # LINE / PARAGRAPH SEPARATOR
    (0x202F, 0x202F),   # NARROW NBSP
    (0x205F, 0x205F),   # MEDIUM MATHEMATICAL SPACE
    (0x3000, 0x3000),   # IDEOGRAPHIC SPACE 全角空格
    # ---- Cf：格式控制符（另见下方「不可见但非 Cf」组）----
    (0x00AD, 0x00AD),   # SOFT HYPHEN 软连字符
    (0x0600, 0x0605),   # 阿拉伯数字符号
    (0x061C, 0x061C),   # ARABIC LETTER MARK
    (0x06DD, 0x06DD),   # 阿拉伯文结束符
    (0x070F, 0x070F),   # 叙利亚文缩写符
    (0x0890, 0x0891),   # 阿拉伯文符号
    (0x08E2, 0x08E2),   # 阿拉伯文分歧结束符
    (0x180E, 0x180E),   # MONGOLIAN VOWEL SEPARATOR
    (0x200B, 0x200F),   # ZWSP / ZWNJ / ZWJ / LRM / RLM
    (0x202A, 0x202E),   # 双向嵌入 / 覆盖
    (0x2060, 0x2064),   # WORD JOINER + 不可见运算符
    (0x2066, 0x206F),   # 双向隔离符 + 数字形状
    (0xFEFF, 0xFEFF),   # BOM / ZWNBSP
    (0xFFF9, 0xFFFB),   # 注释锚点
    (0x110BD, 0x110BD), # KAITHI NUMBER SIGN
    (0x110CD, 0x110CD), # KAITHI NUMBER SIGN ABOVE
    (0x13430, 0x1343F), # 埃及圣书体格式控制符
    (0x1BCA0, 0x1BCA3), # 速记格式控制符
    (0x1D173, 0x1D17A), # 音乐符号
    (0xE0001, 0xE0001), # LANGUAGE TAG
    (0xE0020, 0xE007F), # TAG 字符（可用来隐藏任意文本）
    # ---- 类别非 Cf，但同样不可见、在企业名里纯属噪声 ----
    (0x115F, 0x1160),   # 谚文初声/中声填充符
    (0x17B4, 0x17B5),   # 高棉语固有元音
    (0x180B, 0x180D),   # 蒙古文自由变体选择符
    (0x3164, 0x3164),   # 谚文填充符
    (0xFE00, 0xFE0F),   # 变体选择符 VS1–VS16
    (0xFFA0, 0xFFA0),   # 半角谚文填充符
    (0xE0100, 0xE01EF), # 变体选择符补充区 VS17–VS256
)


def _build_char_class(ranges: tuple[tuple[int, int], ...]) -> str:
    """把码点区间拼成 ``[\\u0009-\\u000d...]`` 形式的字符类文本。

    .. warning:: 转义宽度必须按码点大小选：``\\u`` 只吃**恰好 4 位**十六进制，
        而 ``U+E0001``（TAG 字符）是 5 位 —— 若一律用 ``\\u`` 拼装，
        ``\\ue0001`` 会被正则解析成 ``\\ue000`` **加上一个字面量字符** ``1``，
        不可见字符清不掉，还会顺手把真正的数字 ``1`` 一起吃掉。
        因此 U+FFFF 以上的码点一律用 ``\\U`` + 8 位。
    """
    def esc(cp: int) -> str:
        return f"\\u{cp:04x}" if cp <= 0xFFFF else f"\\U{cp:08x}"

    parts = [
        esc(lo) if lo == hi else f"{esc(lo)}-{esc(hi)}"
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

    缺失值一律 → 空串。原始字符串不做任何回写。

    .. note:: 缺失值判定必须走 :func:`pandas.isna` 而不是 ``isinstance(x, float)``
        —— :func:`~src.data_loader.load_table` 会把名称列 ``astype("string")``，
        空单元格是 ``pd.NA``（类型 ``NAType``），日期列缺失值是 ``pd.NaT``
        （类型 ``NaTType``），两者都**不是** ``float``。只认 float 的话,
        它们会掉进 ``str(value)``，空名称被清洗成字面量字符串 ``'<NA>'`` /
        ``'NAT'`` —— 两条空记录于是互相判「完全匹配 100%」，正是赛题 §5
        明文重罚的误匹配。
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):             # 一次覆盖 None / nan / pd.NA / pd.NaT
            return ""
    except (TypeError, ValueError):     # 非标量（数组 / 列表）→ 交给 str()
        pass
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
#  通用词折叠 —— 判断两条名称的差异是否"纯属表述格式"
# --------------------------------------------------------------------------- #
# 中文企业名可拆成 ``[行政区划][特征字号][行业实词][组织类型]`` 四段。
# 其中**只有「特征字号 + 行业实词」能标识主体**，行政区划与组织类型是
# 任何公司都会带的"包装"。两条名称若把包装全部剥掉后仍然相同，说明它们
# 之间的差异只是写法不同，应当判为同一主体。
#
# 归类原则（唯一的设计开关，增补时照此判断）：
#   * 剥离 = 单独一个词不足以标识主体、且在大量公司名里反复出现
#   * 保留 = 能单独区分主体的实词
# 特别注意：「控股」「快递」「天然气」「化工」「医药」「地产」「汽车」
# 「钢铁」等一律**不进**通用表 —— 它们正是用来区分主体的。
# 「银行」「保险」进通用表，因为真正区分银行的是前置的字号
# （建设 / 工商 / 农业…），该部分由 PROTECTED_BRANDS 单独保护。

# 行政区划（前缀）
REGION_WORDS: tuple[str, ...] = (
    "中国", "中华", "北京", "上海", "天津", "重庆", "广州", "深圳", "广东",
    "江苏", "浙江", "山东", "河南", "河北", "四川", "湖北", "湖南", "福建",
    "安徽", "陕西", "辽宁", "江西", "云南", "广西", "山西", "内蒙古", "新疆",
    "贵州", "甘肃", "海南", "宁夏", "青海", "西藏", "吉林", "黑龙江",
    "香港", "澳门", "台湾", "苏州", "杭州", "南京", "武汉", "成都", "西安",
    "青岛", "大连", "宁波", "厦门", "无锡", "佛山", "东莞", "郑州", "长沙",
    "合肥", "福州", "济南", "沈阳", "哈尔滨",
)

# 通用词（后缀 / 泛行业词）
GENERIC_WORDS: tuple[str, ...] = (
    # 组织类型
    "股份有限公司", "有限责任公司", "集团有限公司", "有限公司", "股份公司",
    "集团", "公司", "事务所", "会计师事务所", "研究院", "研究所", "中心",
    "合伙企业", "厂", "店",
    # 泛行业词（不足以标识主体）
    "科技", "网络", "信息", "技术", "服务", "咨询", "管理", "在线",
    "计算机", "系统", "电子", "实业", "发展", "投资", "国际", "贸易",
    "银行", "保险", "证券", "基金", "信托", "租赁", "石油",
)

# 关键字号保护清单：一方含、另一方不含 → 强惩罚，永不判为同一主体
PROTECTED_BRANDS: tuple[str, ...] = (
    "建设", "工商", "农业", "交通", "招商", "民生", "光大",
    "平安", "浦发", "中信", "兴业", "华夏",
)

# 行业实词：与 GENERIC_WORDS 相反，这些词**能区分业务类型**，因而保留在特征字号里。
# 但两组名称若剥到这个层次后字号仍然相同（如「顺丰快递」/「顺丰控股」都剩「顺丰」），
# 说明它们是**同一字号下的不同业务主体** —— 不判为同一家公司，但互为最佳候选。
INDUSTRY_TOKENS: tuple[str, ...] = (
    "控股", "快递", "物流", "速递", "重工", "乳业", "制药", "新药开发",
    "天然气", "化工", "医药", "地产", "置业", "能源", "汽车", "钢铁",
    "水泥", "航空", "电力", "机械", "工程", "建筑", "环保", "通信",
    "半导体", "新能源", "生物", "材料", "食品", "饮料", "传媒", "教育",
    "旅游", "酒店", "农业", "证券", "基金", "信托", "租赁",
)

# 按长度降序，保证"股份有限公司"先于"有限公司"匹配
_GENERIC_DESC: tuple[str, ...] = tuple(sorted(GENERIC_WORDS, key=len, reverse=True))
_REGION_DESC: tuple[str, ...] = tuple(sorted(REGION_WORDS, key=len, reverse=True))
_GENERIC_SET: frozenset[str] = frozenset(GENERIC_WORDS)
_REGION_SET: frozenset[str] = frozenset(REGION_WORDS)
_INDUSTRY_DESC: tuple[str, ...] = tuple(
    sorted(INDUSTRY_TOKENS, key=len, reverse=True)
)


def collapse_generic(value: object) -> str:
    """剥掉行政区划、括号附注、通用词后剩下的「特征字号」。

    这是本模块最核心的判据 —— **两条名称折叠后相同，差异就纯属表述格式**。

    >>> collapse_generic("上海哔哩哔哩科技有限公司")
    '哔哩哔哩'
    >>> collapse_generic("上海哔哩哔哩有限公司")
    '哔哩哔哩'
    >>> collapse_generic("中国建设银行股份有限公司")     # 保留「建设」
    '建设'
    >>> collapse_generic("中国银行股份有限公司")         # 剥完只剩空串
    ''
    >>> collapse_generic("顺丰控股股份有限公司")         # 「控股」是特征词，保留
    '顺丰控股'

    实现要点：**从尾部迭代剥后缀**，而不是对全文做 ``replace`` ——
    盲替换会误伤字号本身（例如把「中国银行」里的「银行」连同别的词一起吃掉、
    留下 ``（）`` 残渣），也无法处理"叠后缀"（``阿里巴巴网络技术有限公司``
    需要连剥 ``有限公司`` → ``技术`` → ``网络``）。
    """
    text = strip_parens(clean_name(value))
    if not text:
        return ""

    # 1) 反复剥尾部通用词，直到剥不动（支持叠后缀）
    changed = True
    while changed and text:
        changed = False
        for word in _GENERIC_DESC:
            if text.endswith(word) and len(text) > len(word):
                text = text[: len(text) - len(word)]
                changed = True
                break

    # 2) 剩余物若本身就是通用词 / 行政区划（如「中国银行」剥完只剩「中国」），
    #    说明这条名称**没有可标识主体的字号**，折叠结果视为空 ——
    #    否则「中国银行」与「中国石油」都会折叠成「中国」而被误判为同一主体。
    if text in _REGION_SET or text in _GENERIC_SET:
        return ""
    return text


def collapse_industry(value: object) -> str:
    """在 :func:`collapse_generic` 基础上**再剥掉尾部行业实词**，得到"纯字号"。

    用于识别「同一字号下的不同业务主体」：

    >>> collapse_industry("顺丰快递股份有限公司")
    '顺丰'
    >>> collapse_industry("顺丰控股股份有限公司")
    '顺丰'
    >>> collapse_industry("申通快递股份有限公司")   # 字号不同，不会与前两者同组
    '申通'

    末尾同样有"剩余物是通用词/行政区则视为空"的守卫 —— 否则
    ``中国生物制药`` / ``中国建筑`` / ``中国农业银行`` 剥完都会只剩「中国」，
    被凑成一组虚假的"同字号"。
    """
    text = collapse_generic(value)
    if not text:
        return ""
    changed = True
    while changed and text:
        changed = False
        for token in _INDUSTRY_DESC:
            if text.endswith(token) and len(text) > len(token):
                text = text[: len(text) - len(token)]
                changed = True
                break
    if text in _REGION_SET or text in _GENERIC_SET:
        return ""
    return text


def protected_brands_in(value: object) -> frozenset[str]:
    """返回名称中含有的关键字号集合（供 :func:`brand_conflict` 比对）。"""
    text = clean_name(value)
    return frozenset(b for b in PROTECTED_BRANDS if b in text)


def brand_conflict(name_a: object, name_b: object) -> bool:
    """两侧的关键字号集合是否不等。

    典型触发：「中国**建设**银行」含 ``建设``，而「中国银行」不含 ——
    这是两家不同银行，必须拦下，哪怕它们共享 ``中国`` 前缀与 ``银行`` 后缀。
    """
    return protected_brands_in(name_a) != protected_brands_in(name_b)


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
    # ---- 前置豁免：差异**完全落在括号附注内** ----
    # 赛题 §1.1 把「附加备注」列为需要模糊匹配解决的差异类型，其原始示例正是
    #     安永华明会计师事务所（特殊普通合伙） vs 安永华明会计师事务所（北京分所）
    # 这类配对的特征是：**剥掉括号后两条名称完全相同**，差异纯属附注。
    # 若不做豁免，"（北京分所）"会被当成"纯插入 5 字"打成 ×0.40，把本该匹配的
    # 一对压到 60 分以下误判为独有。而下方"取两视角较严值"的机制只会更严，
    # 救不回来 —— 必须在这里显式放行。
    #
    # 注意与"真字号差异"的区别：中国石油天然气 vs 中国石油化工 不含括号，
    # strip_parens 前后相同，因此**不受本豁免影响**，仍按字号差异重罚。
    if strip_parens(s1) == strip_parens(s2):
        return PENALTY_NONE

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
