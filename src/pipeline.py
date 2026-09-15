"""端到端流水线：读取 → 清洗 → 匹配 → 导出。

CLI（``main.py``）与 GUI（``gui_app.py``）共用这一份编排逻辑，保证两个入口
产出的成果文件逐字节一致，也避免"界面版"和"命令行版"各写一套而逐渐跑偏。

进度与日志通过两个回调外泄，本模块自身不做任何打印：

* ``progress(pct: int, message: str)`` —— 整体进度 0–100%
* ``log(text: str)``                   —— 一行运行日志
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

import pandas as pd

from . import config as C
from . import exporter
from . import matcher
from . import preprocessor as pp
from .data_loader import list_sheets, load_a_system, load_b_system

ProgressFn = Callable[[int, str], None]
LogFn = Callable[[str], None]

# 各阶段在整体进度条上占据的区间
SPAN_LOAD = (0.0, 10.0)         # 读取 + 清洗
SPAN_MATCH = (10.0, 85.0)       # 相似度矩阵 + 修正 + 选优
SPAN_EXPORT = (85.0, 100.0)     # 分档 + 写入 Excel

STAGE_TITLES = ("读取与清洗", "相似度比对", "成果导出")


# --------------------------------------------------------------------------- #
#  结果容器
# --------------------------------------------------------------------------- #
@dataclass
class PipelineResult:
    """一次完整运行的产物。"""

    df_a: pd.DataFrame
    df_b: pd.DataFrame
    outcome: matcher.MatchOutcome
    thresholds: matcher.Thresholds
    report: dict[str, Any] | None = None
    elapsed: float = 0.0
    clean_seconds: float = 0.0
    match_seconds: float = 0.0

    @property
    def path(self) -> Path | None:
        return self.report["path"] if self.report else None

    @property
    def sheets(self) -> dict[str, int]:
        return self.report["sheets"] if self.report else {}

    @property
    def frame(self) -> pd.DataFrame:
        """Sheet1「模糊匹配结果」的 DataFrame（供 GUI 表格预览）。"""
        return self.report["frames"]["sheet1"] if self.report else pd.DataFrame()

    @property
    def counts(self) -> dict[str, int]:
        return matcher.summarize(self.outcome.results, self.thresholds)


def _noop_progress(pct: int, message: str) -> None:      # pragma: no cover
    pass


def _noop_log(text: str) -> None:                        # pragma: no cover
    pass


def _remap(pct: float, span: tuple[float, float]) -> int:
    """把 0–100 的阶段内百分比映射到整体区间的百分比。"""
    lo, hi = span
    return int(round(lo + (hi - lo) * min(max(pct, 0.0), 100.0) / 100.0))


# --------------------------------------------------------------------------- #
#  三个阶段，可单独调用
# --------------------------------------------------------------------------- #
def load_and_clean(
    a_file: Path | str | None = None,
    b_file: Path | str | None = None,
    progress: ProgressFn = _noop_progress,
    log: LogFn = _noop_log,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """读取 A / B 两表并新增清洗辅助列（原始名称列保持不动）。"""
    progress(_remap(0, SPAN_LOAD), "读取 A 系统数据…")
    df_a = load_a_system(a_file)
    log(f"[读取] A系统 {len(df_a)} 行 × {len(df_a.columns)} 列  "
        f"Sheet={df_a.attrs.get('sheet')}")
    progress(_remap(35, SPAN_LOAD), "读取 B 系统数据…")

    df_b = load_b_system(b_file)
    log(f"[读取] B系统 {len(df_b)} 行 × {len(df_b.columns)} 列  "
        f"Sheet={df_b.attrs.get('sheet')}")

    progress(_remap(55, SPAN_LOAD), "清洗名称字段…")
    pp.add_clean_columns(df_a, C.A_NAME_COL)
    pp.add_clean_columns(df_b, C.B_NAME_COL)

    changed_a = int((df_a[C.A_NAME_COL].astype(str)
                     != df_a[f"{C.A_NAME_COL}{pp.CLEAN_SUFFIX}"].astype(str)).sum())
    changed_b = int((df_b[C.B_NAME_COL].astype(str)
                     != df_b[f"{C.B_NAME_COL}{pp.CLEAN_SUFFIX}"].astype(str)).sum())
    log(f"[清洗] 全角括号→半角 / 去空格 / 去隐形字符 / 统一大写；"
        f"发生变化的记录 A={changed_a}/{len(df_a)}，B={changed_b}/{len(df_b)}")
    log(f"[清洗] 新增辅助列 {C.A_NAME_COL}{pp.CLEAN_SUFFIX}、"
        f"{C.B_NAME_COL}{pp.CLEAN_SUFFIX}（原始列未被修改）")

    shared = len(set(df_a[f"{C.A_NAME_COL}{pp.CLEAN_SUFFIX}"].tolist())
                 & set(df_b[f"{C.B_NAME_COL}{pp.CLEAN_SUFFIX}"].tolist()))
    log(f"[清洗] 清洗后两表名称完全一致（100% 精确匹配）：{shared} 条")

    progress(_remap(100, SPAN_LOAD), "读取与清洗完成")
    return df_a, df_b


def run_match(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
    workers: int = -1,
    tie_break: bool = True,
    progress: ProgressFn = _noop_progress,
    log: LogFn = _noop_log,
) -> matcher.MatchOutcome:
    """执行相似度比对，进度回调会被换算到整体 10–85% 区间。

    调用方不关心进度（传入默认空回调，如 CLI 场景）时，进度条交还给 matcher 内部
    的 tqdm 处理，命令行观感与直接调用 matcher 完全一致。
    """
    if progress is _noop_progress:
        inner = None
    else:
        def inner(pct: int, message: str) -> None:
            progress(_remap(pct, SPAN_MATCH), message)

    log(f"[匹配] 算法：ratio×{matcher.WEIGHTS['ratio']} + "
        f"partial×{matcher.WEIGHTS['partial_ratio']} + "
        f"WRatio×{matcher.WEIGHTS['WRatio']} + "
        f"core×{matcher.WEIGHTS['core_ratio']}，"
        f"另加字号差异惩罚与简称/全称包含关系修正")
    log(f"[匹配] 阈值：{thresholds.label}（完全 / 高度 / 中低 / 识别下限）")

    outcome = matcher.match_tables(
        df_a, df_b, C.A_NAME_COL, C.B_NAME_COL,
        workers=workers, tie_break=tie_break,
        thresholds=thresholds, progress=inner, verbose=False,
    )
    if inner is not None:
        progress(_remap(100, SPAN_MATCH), "相似度比对完成")
    return outcome


def run_pipeline(
    a_file: Path | str | None = None,
    b_file: Path | str | None = None,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
    *,
    team: str = C.DEFAULT_TEAM,
    out_dir: Path | str | None = None,
    filename: str | None = None,
    export: bool = True,
    workers: int = -1,
    tie_break: bool = True,
    progress: ProgressFn = _noop_progress,
    log: LogFn = _noop_log,
) -> PipelineResult:
    """一次跑完 读取 → 清洗 → 匹配 → 导出。"""
    t0 = time.perf_counter()

    df_a, df_b = load_and_clean(a_file, b_file, progress, log)
    t_clean = time.perf_counter()

    if df_a.empty or df_b.empty:
        raise ValueError(
            f"数据表为空（A={len(df_a)} 行，B={len(df_b)} 行），无法比对"
        )

    outcome = run_match(df_a, df_b, thresholds, workers, tie_break, progress, log)
    t_match = time.perf_counter()

    counts = matcher.summarize(outcome.results, thresholds)
    log("[匹配] " + "  ".join(f"{k}={v}" for k, v in counts.items()))
    log(f"[匹配] 互为最优 {outcome.stats['mutual_best']} 条；"
        f"B 记录被争抢 {outcome.stats['conflict_targets']} 条")

    report: dict[str, Any] | None = None
    if export:
        progress(_remap(0, SPAN_EXPORT), "生成成果报表…")
        report = exporter.export_report(
            df_a, df_b, outcome,
            team=team, out_dir=Path(out_dir) if out_dir else None,
            filename=filename, thresholds=thresholds,
        )
        for name, rows in report["sheets"].items():
            log(f"[导出] {name}：{rows} 行")
        progress(_remap(100, SPAN_EXPORT), "成果报表生成完成")

    total = time.perf_counter() - t0
    log(f"[完成] 总耗时 {total:.2f}s（清洗 {t_clean - t0:.2f}s，"
        f"匹配 {t_match - t_clean:.2f}s）")

    return PipelineResult(
        df_a=df_a, df_b=df_b, outcome=outcome, thresholds=thresholds,
        report=report, elapsed=total,
        clean_seconds=t_clean - t0, match_seconds=t_match - t_clean,
    )


# --------------------------------------------------------------------------- #
#  辅助
# --------------------------------------------------------------------------- #
def available_sheets(path: Path | str) -> Iterable[str]:
    """列出工作簿 Sheet 名（供 GUI 选择 / 排错）。"""
    return list_sheets(path)
