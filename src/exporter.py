"""成果报表导出模块 —— 严格对照《【考题材料】输出成果模板.xlsx》生成 4 个 Sheet。

产出文件：``output/模糊匹配结果_<队名>_<YYYYMMDD>.xlsx``

Sheet 布局（与官方模板逐列一致）
--------------------------------
====================  ==============================================================
Sheet                 列
====================  ==============================================================
模糊匹配结果          匹配序号 / A系统-客户编码 / A系统-客户名称 / B系统-对方户名 /
                      相似度(%) / 匹配状态 / A系统-交易金额 / B系统-交易金额 /
                      金额差异 / 交易日期 / 备注
A系统独有记录         序号 / 客户编码 / 客户名称 / 交易金额（元） / 交易日期 /
                      业务类型 / 部门 / 最高相似度(%) / B系统最佳候选 / 处理建议
B系统独有记录         序号 / 对方户名 / 对方账号 / 交易金额（元） / 交易日期 /
                      流水号 / 摘要 / 最高相似度(%) / A系统最佳候选 / 处理建议
匹配统计汇总          指标 / 数量 / 占比
====================  ==============================================================

样式规范（对应评分表"成果输出规范性 25 分"）
----------------------------------------------
* 表头：加粗 / 居中 / 蓝底 ``4472C4``（官方模板实测值）/ 白字 / 微软雅黑 11pt
* 数据区：微软雅黑 10pt，细边框
* 冻结首行 + 自动筛选（4 个 Sheet 全部启用，Sheet4 冻结到表头行）
* 金额字段：``#,##0.00``；相似度：``0.0``；占比：``0.0%``
* 完全匹配 / 高度匹配 → 整行浅绿底；中低匹配 → 整行浅黄底；低置信度 / 未匹配 → 红色字体
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import config as C
from . import explainer
from . import matcher
from . import preprocessor as pp

# --------------------------------------------------------------------------- #
#  样式常量
# --------------------------------------------------------------------------- #
FONT_NAME = "微软雅黑"
BODY_SIZE = 10          # §4.3「数据区：微软雅黑 10pt」
HEADER_SIZE = 11        # 表头字号跟随官方模板（模板表头为 11pt，§4.3 只约束数据区）

# 表头配色取官方模板的实测值 4472C4（§4.3 只写"蓝色底白字"未给 RGB，
# 但评分表"成果输出规范性 25 分"第 ① 条是"输出表格与官方模板一致"）。
HEADER_FILL = PatternFill("solid", fgColor="4472C4")            # 官方模板蓝
HEADER_FONT = Font(name=FONT_NAME, size=HEADER_SIZE, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

BODY_FONT = Font(name=FONT_NAME, size=BODY_SIZE)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
RIGHT = Alignment(horizontal="right", vertical="center")

GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")             # 浅绿：完全/高度匹配
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")            # 浅黄：中低匹配
RED_FONT = Font(name=FONT_NAME, size=BODY_SIZE, color="C00000")  # 红字：低置信度/未匹配
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="4472C4")   # 与表头同色，四表视觉统一
NOTE_FONT = Font(name=FONT_NAME, size=9, italic=True, color="595959")

_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

MONEY_FMT = "#,##0.00"
SCORE_FMT = "0.0"      # Sheet1 相似度：赛题要求保留 1 位小数
SCORE2_FMT = "0.00"    # 独有记录表：2 位小数，避免 59.96 被显示成 60.0
PCT_FMT = "0.0%"

# 列宽（沿用官方模板）
WIDTHS_RESULT = [10, 18, 46, 46, 13, 14, 18, 18, 14, 14, 24]
WIDTHS_A_ONLY = [8, 16, 46, 18, 14, 16, 14, 15, 46, 20]
WIDTHS_B_ONLY = [8, 46, 22, 18, 14, 22, 14, 15, 46, 20]
WIDTHS_SUMMARY = [36, 20, 14]

SHEET_RESULT = "模糊匹配结果"
SHEET_A_ONLY = "A系统独有记录"
SHEET_B_ONLY = "B系统独有记录"
SHEET_SUMMARY = "匹配统计汇总"

MONEY_COLS = ("A系统-交易金额", "B系统-交易金额", "金额差异", "交易金额（元）")


# --------------------------------------------------------------------------- #
#  小工具
# --------------------------------------------------------------------------- #
def _py(value: Any) -> Any:
    """把 numpy / pandas 标量转成 openpyxl 能写的原生 Python 类型。"""
    if value is None or value is pd.NA:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if hasattr(value, "item"):          # numpy 标量
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    return value


def _fmt_score(value: float) -> float:
    """Sheet1 相似度：保留 1 位小数（赛题 §4.1 明文要求）。"""
    return round(float(value), 1)


def _fmt_score2(value: float) -> float:
    """独有记录表的「最高相似度」：保留 2 位小数。

    这些表的候选分数**必然低于匹配下限**，用 1 位小数时 59.96 会被显示成
    「60.0」，与同行的「候选相似度仍低于匹配阈值」文案看起来自相矛盾。
    多留一位小数即可消除这种四舍五入造成的错觉。
    """
    return round(float(value), 2)


def _cell(row: pd.Series, col: str, default: Any = "") -> Any:
    """安全取单元格值 —— 列不存在时返回 ``default``，不抛 ``KeyError``。

    赛题只规定了"A 系统 / B 系统"两个工作簿的**必备列**，用户换一份自己的
    数据时少一列（例如没有「业务类型」）是常态。原始实现用 ``arow[col]``
    裸取，缺列直接 ``KeyError`` 冒到顶层，整条流水线崩掉 —— 命中评分表
    "大数据量稳定性 15 分"的"无闪退、报错"项。缺列不该让导出失败。
    """
    if col not in row.index:
        return default
    value = row[col]
    return default if _py(value) is None else value


def _num(row: pd.Series, col: str) -> float:
    """安全取数值：缺列 / 空值 / 非数字文本一律退化为 ``0.0``。

    赛题 FAQ Q4 明确「金额不作为匹配依据」，所以读不出金额时退化为 0
    不影响**任何匹配结论**，只影响展示 —— 但绝不该让整个导出崩掉。
    """
    raw = _cell(row, col, None)
    if raw is None:
        return 0.0
    try:
        return float(str(raw).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


# --------------------------------------------------------------------------- #
#  Sheet 1：模糊匹配结果
# --------------------------------------------------------------------------- #
def build_result_table(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    outcome: matcher.MatchOutcome,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
) -> pd.DataFrame:
    """Sheet1 —— 只放"匹配上"的记录（完全 / 高度 / 中低 / 低置信度）。

    未达 ``thresholds.floor`` 的 A 记录不进入本表，归入 Sheet2「A系统独有」。
    """
    rows: list[dict[str, Any]] = []
    seq = 0
    for r in outcome.results:
        if r.score < thresholds.floor:
            continue                        # 低于地板分 → 归入 Sheet2「A系统独有」
        seq += 1
        arow, brow = df_a.iloc[r.a_index], df_b.iloc[r.b_index]
        amt_a = _num(arow, C.A_AMOUNT_COL)
        amt_b = _num(brow, C.B_AMOUNT_COL)
        rows.append(
            {
                "匹配序号": seq,
                "A系统-客户编码": _cell(arow, C.A_KEY_COL),
                "A系统-客户名称": r.a_name,
                "B系统-对方户名": r.b_name,
                "相似度(%)": _fmt_score(r.score),
                "匹配状态": thresholds.bucket(r.score),
                "A系统-交易金额": amt_a,
                "B系统-交易金额": amt_b,
                "金额差异": round(amt_a - amt_b, 2),
                "交易日期": _cell(arow, C.DATE_COL),
                "备注": _remark(r, thresholds),
            }
        )
    cols = list(RESULT_COLUMNS)
    return pd.DataFrame(rows, columns=cols)


RESULT_COLUMNS = (
    "匹配序号", "A系统-客户编码", "A系统-客户名称", "B系统-对方户名", "相似度(%)",
    "匹配状态", "A系统-交易金额", "B系统-交易金额", "金额差异", "交易日期", "备注",
)


def _remark(r: matcher.MatchResult,
            thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS) -> str:
    """「备注」列 —— 由可解释性模块产出的规则标签串。

    标签来自受控词表（:data:`explainer.TAG_VOCAB`），例如
    「包含关系（丢弃公司后缀）（A多出「股份有限公司」）」、
    「疑似错别字（「限」vs「线」）」、「字号不同」+「非互为最优」。
    审计师可直接在 Excel 里按标签筛选复核。
    """
    return explainer.explain_text(r.a_name, r.b_name, result=r, thresholds=thresholds)


# --------------------------------------------------------------------------- #
#  Sheet 2：A 系统独有记录
# --------------------------------------------------------------------------- #
def build_a_only_table(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    outcome: matcher.MatchOutcome,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
) -> pd.DataFrame:
    """Sheet2 —— B 系统中未找到达标匹配的 A 记录，附最高候选与处理建议。"""
    rows: list[dict[str, Any]] = []
    score_matrix = outcome.score_matrix
    seq = 0
    for r in outcome.results:
        if r.score >= thresholds.floor:
            continue
        seq += 1
        arow = df_a.iloc[r.a_index]
        # 真实最佳候选直接取相似度矩阵的**列最大值**，不能用 r.b_index ——
        # 那是经一对一约束调整后的最终分配，会把这条记录的候选换成一条
        # 毫不相干的低分记录，审计师就看不到"最像的那条"到底是什么了。
        if score_matrix.shape[1]:
            best_j = int(score_matrix[r.a_index].argmax())
            best_name = str(df_b.iloc[best_j][C.B_NAME_COL])
            best_score = float(score_matrix[r.a_index, best_j])
        else:
            best_j, best_name, best_score = -1, "", 0.0
        # 该候选分数达标、却已被别的 A 记录认领（本记录必然未被认领）
        occupied = bool(best_score >= thresholds.floor and best_j in outcome.b_taken)
        rows.append(
            {
                "序号": seq,
                C.A_KEY_COL: _cell(arow, C.A_KEY_COL),
                C.A_NAME_COL: r.a_name,
                C.A_AMOUNT_COL: _num(arow, C.A_AMOUNT_COL),
                C.DATE_COL: _cell(arow, C.DATE_COL),
                "业务类型": _cell(arow, "业务类型"),
                "部门": _cell(arow, "部门"),
                "最高相似度(%)": _fmt_score2(best_score),
                "B系统最佳候选": best_name,
                "处理建议": _advice_a_only(best_score, occupied=occupied),
            }
        )
    return pd.DataFrame(rows, columns=list(A_ONLY_COLUMNS))


A_ONLY_COLUMNS = (
    "序号", "客户编码", "客户名称", "交易金额（元）", "交易日期", "业务类型", "部门",
    "最高相似度(%)", "B系统最佳候选", "处理建议",
)


def _advice_a_only(score: float, occupied: bool = False) -> str:
    """A 系统独有记录的复核建议（阈值 60 = 匹配地板分，低于它的候选均为噪声级）。

    ``occupied=True`` 表示最佳候选分数已达标，但那条 B 流水已被另一条 A 记录
    认领（一对一约束）。此时若照常输出「候选相似度仍低于匹配阈值」，
    就与「最高相似度」列的数值自相矛盾。
    """
    if occupied:
        return "最佳候选流水已被另一条客户记录匹配（一对一约束），建议核对是否一对多"
    if score < 40:
        return "B系统无可信候选，建议核实是否为未达账项或单据缺失"
    if score < 55:
        return "B系统无可信候选（最高<55%），建议核实该笔款项是否已入账"
    return "候选相似度仍低于匹配阈值，建议人工核实是否同一主体"


# --------------------------------------------------------------------------- #
#  Sheet 3：B 系统独有记录
# --------------------------------------------------------------------------- #
def build_b_only_table(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    outcome: matcher.MatchOutcome,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
) -> pd.DataFrame:
    """Sheet3 —— A 系统中无任何记录认领的 B 记录，附 A 侧最佳候选。"""
    claimed = set(outcome.b_taken.keys())          # 已被 A 记录认领（且达标）的 B 行
    score = outcome.score_matrix
    rows: list[dict[str, Any]] = []
    seq = 0
    for j in range(len(df_b)):
        if j in claimed:
            continue
        seq += 1
        brow = df_b.iloc[j]
        occupied = False
        if score.shape[0]:
            best_i = int(score[:, j].argmax())
            best_name = str(df_a.iloc[best_i][C.A_NAME_COL])
            best_score = float(score[best_i, j])
            # 一对一约束：**本笔的候选分数已达标**，但那条 A 记录已被另一笔流水认领，
            # 所以本笔用不上它。必须把这件事告诉审计师，否则「最高相似度 95% 却判独有」
            # 与「候选相似度仍低于匹配阈值」的文案自相矛盾。
            # 门槛必须同时卡住本笔分数 —— 否则 30 分级别的噪声候选也会命中，
            # 把更有用的「无可信客户」文案顶掉。
            cand = outcome.results[best_i]
            occupied = bool(
                best_score >= thresholds.floor
                and cand.b_index is not None
                and cand.b_index != j
            )
        else:
            best_name, best_score = "", 0.0
        rows.append(
            {
                "序号": seq,
                C.B_NAME_COL: _cell(brow, C.B_NAME_COL),
                C.B_ACCOUNT_COL: _cell(brow, C.B_ACCOUNT_COL),
                C.B_AMOUNT_COL: _num(brow, C.B_AMOUNT_COL),
                C.DATE_COL: _cell(brow, C.DATE_COL),
                C.B_KEY_COL: _cell(brow, C.B_KEY_COL),
                "摘要": _cell(brow, "摘要"),
                "最高相似度(%)": _fmt_score2(best_score),
                "A系统最佳候选": best_name,
                "处理建议": _advice_b_only(best_score, occupied=occupied),
            }
        )
    return pd.DataFrame(rows, columns=list(B_ONLY_COLUMNS))


B_ONLY_COLUMNS = (
    "序号", "对方户名", "对方账号", "交易金额（元）", "交易日期", "流水号", "摘要",
    "最高相似度(%)", "A系统最佳候选", "处理建议",
)


def _advice_b_only(score: float, occupied: bool = False) -> str:
    """B 系统独有记录的处理建议。

    ``occupied=True`` 表示这条记录的最佳 A 候选**分数已达匹配线，但那条 A 记录
    已经被另一笔流水占走**（一对一约束下不可重复使用）。此时若照常输出
    「候选相似度仍低于匹配阈值」，就与「最高相似度」列显示的 95% 自相矛盾 ——
    必须点明是被占用，而不是分数不够。
    """
    if occupied:
        return "最佳候选客户已被另一条流水匹配（一对一约束），本笔可能为重复入账，建议核对是否一对多"
    if score < 40:
        return "A系统无可信客户，可能为新客户，建议补充客户档案"
    if score < 55:
        return "A系统无可信客户（最高<55%），建议核实是否为销售未开票或预收款"
    return "候选相似度仍低于匹配阈值，建议人工核实是否同一主体"


# --------------------------------------------------------------------------- #
#  Sheet 4：匹配统计汇总
# --------------------------------------------------------------------------- #
def build_summary_rows(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    outcome: matcher.MatchOutcome,
    sheet1: pd.DataFrame,
    sheet2: pd.DataFrame,
    sheet3: pd.DataFrame,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
) -> list[tuple[str, int, float]]:
    """返回 ``[(指标, 数量, 占比), ...]``，指标文案随阈值动态生成。"""
    n_a, n_b = len(df_a), len(df_b)
    counts = matcher.summarize(outcome.results, thresholds)     # 按 A 记录分档
    matched = len(sheet1)
    t = thresholds

    def pct(num: int, den: int) -> float:
        return round(num / den, 4) if den else 0.0

    # 指标名**逐字对齐赛题 §4.1「Sheet 4：匹配统计汇总」的字段表**。
    # 说明：
    # * 阈值仍用 f-string 插值而非写死 90/70/60 —— 默认阈值下与 §4.1 一字不差，
    #   用户在 GUI 上调过阈值后，标签会跟着变，不会出现"标 90 实际按 95 算"。
    # * 「A系统独有记录」不带括号阈值：原先写「A系统独有（<60%）」，但同一工作簿
    #   Sheet2 里存在最高相似度恰为 60.00 的记录（60 并不 < 60），标签与自家数据打架。
    return [
        ("A系统总记录数", n_a, 1.0),
        ("B系统总记录数", n_b, 1.0),
        ("完全匹配（100%）", counts["完全匹配"], pct(counts["完全匹配"], n_a)),
        (f"高度匹配（≥{t.high:g}%）", counts["高度匹配"],
         pct(counts["高度匹配"], n_a)),
        (f"中低匹配（{t.low:g}% ≤ score < {t.high:g}%）", counts["中低匹配"],
         pct(counts["中低匹配"], n_a)),
        (f"低置信度匹配（<{t.low:g}%）", counts["低置信度匹配"],
         pct(counts["低置信度匹配"], n_a)),
        ("A系统独有记录", counts["A系统独有"], pct(counts["A系统独有"], n_a)),
        ("B系统独有记录", len(sheet3), pct(len(sheet3), n_b)),
        ("总匹配成功数", matched, pct(matched, n_a)),
    ]


def summary_note(thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS) -> str:
    t = thresholds
    return (
        "说明：完全匹配 / 高度匹配 / 中低匹配 / 低置信度匹配 / A系统独有 的占比以「A系统总记录数」"
        "为分母；B系统独有的占比以「B系统总记录数」为分母。相似度按清洗（去空格、统一全半角括号、"
        f"统一大小写）后的名称计算，识别阈值 {t.floor:g}%，分档阈值 {t.low:g}% / {t.high:g}%。"
    )


# --------------------------------------------------------------------------- #
#  通用写入器
# --------------------------------------------------------------------------- #
def _write_table(
    ws: Worksheet,
    df: pd.DataFrame,
    widths: Sequence[float],
    *,
    money_cols: Iterable[str] = (),
    score_cols: Iterable[str] = (),
    score_fmt: str = SCORE_FMT,
    center_cols: Iterable[str] = (),
    row_style: Callable[[dict[str, Any]], tuple[PatternFill | None,
                                                Font | None]] | None = None,
    start_row: int = 1,
) -> int:
    """把 DataFrame 写入工作表并套用统一格式，返回最后一行的行号。"""
    money_set, score_set, center_set = set(money_cols), set(score_cols), set(center_cols)

    # 表头
    for c, name in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=c, value=str(name))
        cell.fill, cell.font, cell.alignment, cell.border = (
            HEADER_FILL, HEADER_FONT, HEADER_ALIGN, BORDER,
        )
    ws.row_dimensions[start_row].height = 30

    # 数据区
    for r, (_, record) in enumerate(df.iterrows(), start=start_row + 1):
        fill, font = (None, None)
        if row_style is not None:
            fill, font = row_style(record)
        for c, name in enumerate(df.columns, start=1):
            cell = ws.cell(row=r, column=c, value=_py(record[name]))
            cell.font = font or BODY_FONT
            cell.border = BORDER
            if fill is not None:
                cell.fill = fill
            if name in money_set:
                cell.number_format, cell.alignment = MONEY_FMT, RIGHT
            elif name in score_set:
                cell.number_format, cell.alignment = score_fmt, CENTER
            elif name in center_set:
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
    last_row = start_row + len(df)

    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    if len(df):
        ws.auto_filter.ref = (
            f"A{start_row}:{get_column_letter(len(df.columns))}{last_row}"
        )
    return last_row


def _result_row_style(record: dict[str, Any]) -> tuple[PatternFill | None, Font | None]:
    """Sheet1 条件格式：完全/高度 → 浅绿；中低 → 浅黄；低置信度 → 红字。"""
    status = record["匹配状态"]
    if status in ("完全匹配", "高度匹配"):
        return GREEN_FILL, None
    if status == "中低匹配":
        return YELLOW_FILL, None
    return None, RED_FONT


def _unmatched_row_style(record: dict[str, Any]) -> tuple[None, Font | None]:
    """Sheet2 / Sheet3：未匹配行 → 红色字体（相似度回升到中高时仍提示红字，便于定位）。"""
    return None, RED_FONT


# --------------------------------------------------------------------------- #
#  主入口
# --------------------------------------------------------------------------- #
def default_filename(team: str = C.DEFAULT_TEAM, when: _dt.date | None = None) -> str:
    """``模糊匹配结果_<队名>_<YYYYMMDD>.xlsx``"""
    when = when or _dt.date.today()
    return f"模糊匹配结果_{team}_{when:%Y%m%d}.xlsx"


def export_report(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    outcome: matcher.MatchOutcome,
    *,
    team: str = C.DEFAULT_TEAM,
    out_dir: Path | None = None,
    filename: str | None = None,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
) -> dict[str, Any]:
    """生成 4-Sheet 规范成果报表，返回 ``{path, sheets: {名: 行数}}``。"""
    out_dir = Path(out_dir) if out_dir else C.OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / (filename or default_filename(team))

    sheet1 = build_result_table(df_a, df_b, outcome, thresholds=thresholds)
    sheet2 = build_a_only_table(df_a, df_b, outcome, thresholds=thresholds)
    sheet3 = build_b_only_table(df_a, df_b, outcome, thresholds=thresholds)
    summary = build_summary_rows(df_a, df_b, outcome, sheet1, sheet2, sheet3,
                                 thresholds=thresholds)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = SHEET_RESULT
    _write_table(
        ws1, sheet1, WIDTHS_RESULT,
        money_cols=MONEY_COLS, score_cols=("相似度(%)",),
        center_cols=("匹配序号", "匹配状态", "交易日期"),
        row_style=_result_row_style,
    )

    ws2 = wb.create_sheet(SHEET_A_ONLY)
    _write_table(
        ws2, sheet2, WIDTHS_A_ONLY,
        money_cols=("交易金额（元）",), score_cols=("最高相似度(%)",),
        score_fmt=SCORE2_FMT,
        center_cols=("序号", "交易日期", "业务类型", "部门"),
        row_style=_unmatched_row_style,
    )

    ws3 = wb.create_sheet(SHEET_B_ONLY)
    _write_table(
        ws3, sheet3, WIDTHS_B_ONLY,
        money_cols=("交易金额（元）",), score_cols=("最高相似度(%)",),
        score_fmt=SCORE2_FMT,
        center_cols=("序号", "交易日期", "摘要"),
        row_style=_unmatched_row_style,
    )

    ws4 = wb.create_sheet(SHEET_SUMMARY)
    _write_summary(ws4, summary, thresholds)

    path = _save_with_fallback(wb, path)
    return {
        "path": path,
        "sheets": {
            SHEET_RESULT: len(sheet1),
            SHEET_A_ONLY: len(sheet2),
            SHEET_B_ONLY: len(sheet3),
            SHEET_SUMMARY: len(summary),
        },
        "frames": {"sheet1": sheet1, "sheet2": sheet2, "sheet3": sheet3,
                   "summary": summary},
    }


def _save_with_fallback(wb: Workbook, path: Path, tries: int = 9) -> Path:
    """保存工作簿；文件被占用（Excel 正打开着）时自动改用带序号的新文件名。

    这是评测现场最常见的意外 —— 上一轮生成的文件还开着，直接覆盖会抛
    ``PermissionError`` 让整个程序崩掉。这里改为降级另存，保证流程不中断，
    并把实际落盘的路径返回给调用方展示。
    """
    try:
        wb.save(path)
        return path
    except PermissionError:
        pass
    for k in range(2, tries + 2):
        alt = path.with_name(f"{path.stem}({k}){path.suffix}")
        try:
            wb.save(alt)
            return alt
        except PermissionError:
            continue
    raise PermissionError(
        f"目标文件被占用且无法另存：{path}\n请关闭 Excel 中打开的该文件后重试。"
    )


def _write_summary(ws: Worksheet, summary: list[tuple[str, int, float]],
                   thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS) -> None:
    """Sheet4 自定义布局：合并标题 → 空行 → 表头 → 指标 → 空行 → 合并说明。"""
    n_cols = len(WIDTHS_SUMMARY)
    last_col = get_column_letter(n_cols)

    # 标题
    ws.merge_cells(f"A1:{last_col}1")
    title = ws["A1"]
    title.value = SHEET_SUMMARY
    title.font, title.fill, title.alignment = TITLE_FONT, TITLE_FILL, HEADER_ALIGN
    ws.row_dimensions[1].height = 28
    for c in range(1, n_cols + 1):
        ws.cell(row=1, column=c).fill = TITLE_FILL

    # 表头
    for c, name in enumerate(("指标", "数量", "占比"), start=1):
        cell = ws.cell(row=3, column=c, value=name)
        cell.fill, cell.font, cell.alignment, cell.border = (
            HEADER_FILL, HEADER_FONT, HEADER_ALIGN, BORDER,
        )
    ws.row_dimensions[3].height = 24

    # 指标
    for i, (name, count, ratio) in enumerate(summary):
        r = 4 + i
        is_total = name.endswith("总记录数")
        a = ws.cell(row=r, column=1, value=name)
        b = ws.cell(row=r, column=2, value=int(count))
        c = ws.cell(row=r, column=3, value=float(ratio))
        for cell in (a, b, c):
            cell.font = Font(name=FONT_NAME, size=BODY_SIZE, bold=is_total)
            cell.border = BORDER
        a.alignment = LEFT
        b.alignment, b.number_format = CENTER, "#,##0"
        c.alignment, c.number_format = CENTER, PCT_FMT

    # 说明
    note_row = 4 + len(summary) + 1
    ws.merge_cells(f"A{note_row}:{last_col}{note_row}")
    note = ws.cell(row=note_row, column=1, value=summary_note(thresholds))
    note.font, note.alignment = NOTE_FONT, Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[note_row].height = 42

    for i, w in enumerate(WIDTHS_SUMMARY, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # §4.3「冻结首行，启用自动筛选」—— Sheet4 的表头在第 3 行（1 标题 / 2 空行），
    # 所以冻结到 A4（锁住标题 + 表头），筛选区从表头行到最后一个指标行。
    # 原先只设了列宽、漏了这两项，四个 Sheet 里唯独这张没有，属明确不合规。
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{last_col}{3 + len(summary)}"
