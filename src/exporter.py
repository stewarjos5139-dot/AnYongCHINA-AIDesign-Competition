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
* 表头：加粗 / 居中 / 蓝底 ``004080`` / 白字 / 微软雅黑 10pt
* 数据区：微软雅黑 10pt，细边框
* 冻结首行 + 自动筛选
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
BODY_SIZE = 10

HEADER_FILL = PatternFill("solid", fgColor="004080")            # 深蓝底
HEADER_FONT = Font(name=FONT_NAME, size=BODY_SIZE, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

BODY_FONT = Font(name=FONT_NAME, size=BODY_SIZE)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
RIGHT = Alignment(horizontal="right", vertical="center")

GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")             # 浅绿：完全/高度匹配
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")            # 浅黄：中低匹配
RED_FONT = Font(name=FONT_NAME, size=BODY_SIZE, color="C00000")  # 红字：低置信度/未匹配
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="004080")
NOTE_FONT = Font(name=FONT_NAME, size=9, italic=True, color="595959")

_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

MONEY_FMT = "#,##0.00"
SCORE_FMT = "0.0"
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
    return round(float(value), 1)


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
        amt_a = float(arow[C.A_AMOUNT_COL])
        amt_b = float(brow[C.B_AMOUNT_COL])
        rows.append(
            {
                "匹配序号": seq,
                "A系统-客户编码": arow[C.A_KEY_COL],
                "A系统-客户名称": r.a_name,
                "B系统-对方户名": r.b_name,
                "相似度(%)": _fmt_score(r.score),
                "匹配状态": thresholds.bucket(r.score),
                "A系统-交易金额": amt_a,
                "B系统-交易金额": amt_b,
                "金额差异": round(amt_a - amt_b, 2),
                "交易日期": arow[C.DATE_COL],
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
    seq = 0
    for r in outcome.results:
        if r.score >= thresholds.floor:
            continue
        seq += 1
        arow = df_a.iloc[r.a_index]
        best_j = r.b_index if r.b_index is not None else -1
        best_name = r.b_name if best_j >= 0 else ""
        best_score = r.score
        rows.append(
            {
                "序号": seq,
                C.A_KEY_COL: arow[C.A_KEY_COL],
                C.A_NAME_COL: r.a_name,
                C.A_AMOUNT_COL: float(arow[C.A_AMOUNT_COL]),
                C.DATE_COL: arow[C.DATE_COL],
                "业务类型": arow.get("业务类型", ""),
                "部门": arow.get("部门", ""),
                "最高相似度(%)": _fmt_score(best_score),
                "B系统最佳候选": best_name,
                "处理建议": _advice_a_only(best_score),
            }
        )
    return pd.DataFrame(rows, columns=list(A_ONLY_COLUMNS))


A_ONLY_COLUMNS = (
    "序号", "客户编码", "客户名称", "交易金额（元）", "交易日期", "业务类型", "部门",
    "最高相似度(%)", "B系统最佳候选", "处理建议",
)


def _advice_a_only(score: float) -> str:
    """A 系统独有记录的复核建议（阈值 60 = 匹配地板分，低于它的候选均为噪声级）。"""
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
        if score.shape[0]:
            best_i = int(score[:, j].argmax())
            best_name = str(df_a.iloc[best_i][C.A_NAME_COL])
            best_score = float(score[best_i, j])
        else:
            best_name, best_score = "", 0.0
        rows.append(
            {
                "序号": seq,
                C.B_NAME_COL: brow[C.B_NAME_COL],
                C.B_ACCOUNT_COL: brow.get(C.B_ACCOUNT_COL, ""),
                C.B_AMOUNT_COL: float(brow[C.B_AMOUNT_COL]),
                C.DATE_COL: brow[C.DATE_COL],
                C.B_KEY_COL: brow.get(C.B_KEY_COL, ""),
                "摘要": brow.get("摘要", ""),
                "最高相似度(%)": _fmt_score(best_score),
                "A系统最佳候选": best_name,
                "处理建议": _advice_b_only(best_score),
            }
        )
    return pd.DataFrame(rows, columns=list(B_ONLY_COLUMNS))


B_ONLY_COLUMNS = (
    "序号", "对方户名", "对方账号", "交易金额（元）", "交易日期", "流水号", "摘要",
    "最高相似度(%)", "A系统最佳候选", "处理建议",
)


def _advice_b_only(score: float) -> str:
    """B 系统独有记录的处理建议。"""
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

    return [
        ("A系统总记录数", n_a, 1.0),
        ("B系统总记录数", n_b, 1.0),
        ("完全匹配（100%）", counts["完全匹配"], pct(counts["完全匹配"], n_a)),
        (f"高度匹配（{t.high:g}%–<100%）", counts["高度匹配"],
         pct(counts["高度匹配"], n_a)),
        (f"中低匹配（{t.low:g}%–<{t.high:g}%）", counts["中低匹配"],
         pct(counts["中低匹配"], n_a)),
        (f"低置信度匹配（{t.floor:g}%–<{t.low:g}%）", counts["低置信度匹配"],
         pct(counts["低置信度匹配"], n_a)),
        (f"A系统独有（<{t.floor:g}%）", counts["A系统独有"],
         pct(counts["A系统独有"], n_a)),
        ("B系统独有", len(sheet3), pct(len(sheet3), n_b)),
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
                cell.number_format, cell.alignment = SCORE_FMT, CENTER
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
#  批量：总览 + 合并明细
# --------------------------------------------------------------------------- #
SOURCE_COLS = ("数据来源文件（A）", "数据来源文件（B）")

SHEET_BATCH_OVERVIEW = "批量总览"
SHEET_BATCH_RESULT = "全部匹配明细"
SHEET_BATCH_A_ONLY = "全部A系统独有"
SHEET_BATCH_B_ONLY = "全部B系统独有"

BATCH_OVERVIEW_COLUMNS = (
    "任务序号", "数据来源文件（A）", "数据来源文件（B）",
    "A系统记录数", "B系统记录数",
    "完全匹配", "高度匹配", "中低匹配", "低置信度匹配",
    "A系统独有", "B系统独有", "匹配成功率",
)
WIDTHS_BATCH = [8, 34, 34, 12, 12, 10, 10, 10, 12, 11, 11, 12]
WIDTHS_BATCH_DETAIL = [34, 34] + WIDTHS_RESULT


@dataclass
class BatchEntry:
    """一个 (A文件, B文件) 配对跑完后的产物，供批量总览汇总。"""

    label: str
    a_file: Path
    b_file: Path
    n_a: int
    n_b: int
    counts: dict[str, int]
    sheet1: pd.DataFrame
    sheet2: pd.DataFrame
    sheet3: pd.DataFrame
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS

    @property
    def matched(self) -> int:
        return len(self.sheet1)

    @property
    def hit_rate(self) -> float:
        return round(self.matched / self.n_a, 4) if self.n_a else 0.0


def batch_overview_table(entries: Sequence[BatchEntry]) -> pd.DataFrame:
    """批量总览表：一行一个配对。"""
    rows = []
    for i, e in enumerate(entries, start=1):
        c = e.counts
        rows.append({
            "任务序号": i,
            "数据来源文件（A）": e.a_file.name,
            "数据来源文件（B）": e.b_file.name,
            "A系统记录数": e.n_a,
            "B系统记录数": e.n_b,
            "完全匹配": c.get("完全匹配", 0),
            "高度匹配": c.get("高度匹配", 0),
            "中低匹配": c.get("中低匹配", 0),
            "低置信度匹配": c.get("低置信度匹配", 0),
            "A系统独有": c.get("A系统独有", 0),
            "B系统独有": len(e.sheet3),
            "匹配成功率": e.hit_rate,
        })
    return pd.DataFrame(rows, columns=list(BATCH_OVERVIEW_COLUMNS))


def stack_detail(entries: Sequence[BatchEntry], which: str) -> pd.DataFrame:
    """把各配对的 Sheet1/2/3 纵向拼接，并在最前面加上「数据来源文件」列。"""
    parts: list[pd.DataFrame] = []
    for e in entries:
        frame = getattr(e, which)
        if frame.empty:
            continue
        block = frame.copy()
        block.insert(0, SOURCE_COLS[0], e.a_file.name)
        block.insert(1, SOURCE_COLS[1], e.b_file.name)
        parts.append(block)
    if not parts:
        base = {"sheet1": RESULT_COLUMNS, "sheet2": A_ONLY_COLUMNS,
                "sheet3": B_ONLY_COLUMNS}[which]
        return pd.DataFrame(columns=list(SOURCE_COLS) + list(base))
    return pd.concat(parts, ignore_index=True)


def export_batch_overview(
    entries: Sequence[BatchEntry],
    *,
    team: str = C.DEFAULT_TEAM,
    out_dir: Path | None = None,
    filename: str | None = None,
    when: _dt.date | None = None,
) -> Path:
    """生成批量总览工作簿（总览 + 三张合并明细）。"""
    out_dir = Path(out_dir) if out_dir else C.OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    when = when or _dt.date.today()
    path = out_dir / (filename or f"批量汇总_{team}_{when:%Y%m%d}.xlsx")

    overview = batch_overview_table(entries)
    detail = stack_detail(entries, "sheet1")
    a_only = stack_detail(entries, "sheet2")
    b_only = stack_detail(entries, "sheet3")

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_BATCH_OVERVIEW
    _write_table(
        ws, overview, WIDTHS_BATCH,
        score_cols=("匹配成功率",),
        center_cols=("任务序号",) + tuple(c for c in BATCH_OVERVIEW_COLUMNS[3:]),
        row_style=None,
    )
    # 合计行
    last = 1 + len(overview)
    if len(overview):
        total = ws.cell(row=last + 1, column=1, value="合计")
        total.font = Font(name=FONT_NAME, size=BODY_SIZE, bold=True)
        total.border = BORDER
        for col, key in ((4, "A系统记录数"), (5, "B系统记录数"), (6, "完全匹配"),
                         (7, "高度匹配"), (8, "中低匹配"), (9, "低置信度匹配"),
                         (10, "A系统独有"), (11, "B系统独有")):
            cell = ws.cell(row=last + 1, column=col,
                           value=int(overview[key].sum()))
            cell.font = Font(name=FONT_NAME, size=BODY_SIZE, bold=True)
            cell.border, cell.alignment = BORDER, CENTER
        n_a_total = int(overview["A系统记录数"].sum())
        rate = ws.cell(row=last + 1, column=12,
                       value=round(int(overview["完全匹配"].sum()
                                       + overview["高度匹配"].sum()
                                       + overview["中低匹配"].sum()
                                       + overview["低置信度匹配"].sum())
                                   / n_a_total, 4) if n_a_total else 0.0)
        rate.font = Font(name=FONT_NAME, size=BODY_SIZE, bold=True)
        rate.border, rate.alignment, rate.number_format = BORDER, CENTER, PCT_FMT

    _write_table(
        wb.create_sheet(SHEET_BATCH_RESULT), detail, WIDTHS_BATCH_DETAIL,
        money_cols=MONEY_COLS, score_cols=("相似度(%)",),
        center_cols=("匹配序号", "匹配状态", "交易日期"),
        row_style=_result_row_style,
    )
    _write_table(
        wb.create_sheet(SHEET_BATCH_A_ONLY), a_only, WIDTHS_BATCH_DETAIL,
        money_cols=("交易金额（元）",), score_cols=("最高相似度(%)",),
        center_cols=("序号", "交易日期", "业务类型", "部门"),
        row_style=_unmatched_row_style,
    )
    _write_table(
        wb.create_sheet(SHEET_BATCH_B_ONLY), b_only, WIDTHS_BATCH_DETAIL,
        money_cols=("交易金额（元）",), score_cols=("最高相似度(%)",),
        center_cols=("序号", "交易日期", "摘要"),
        row_style=_unmatched_row_style,
    )

    return _save_with_fallback(wb, path)


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
        center_cols=("序号", "交易日期", "业务类型", "部门"),
        row_style=_unmatched_row_style,
    )

    ws3 = wb.create_sheet(SHEET_B_ONLY)
    _write_table(
        ws3, sheet3, WIDTHS_B_ONLY,
        money_cols=("交易金额（元）",), score_cols=("最高相似度(%)",),
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
