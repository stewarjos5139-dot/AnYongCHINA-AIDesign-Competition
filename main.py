"""Topic03 多源运营数据"相似度"模糊匹配工具 —— 主执行文件。

用法：
    python main.py                 # 全流程：环境自检 → 读取 → 清洗 → 匹配 → 校验打印
    python main.py -n 5            # 预览前 5 条原始名称记录
    python main.py --no-progress   # 关闭进度条（适合重定向到文件）
"""

from __future__ import annotations

import argparse
import platform
import sys
import time
from pathlib import Path
from typing import Any

# Windows 控制台默认可能是 GBK，强制 UTF-8 输出，避免中文/特殊符号乱码
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
    except (AttributeError, ValueError):
        pass

import pandas as pd

from src import config, matcher
from src import exporter
from src import pipeline
from src import preprocessor as pp
from src.data_loader import DataLoadError, list_sheets, raw_names, unicode_report

LINE = "=" * 78
SUB = "-" * 78
COL_A = config.A_NAME_COL
COL_B = config.B_NAME_COL


# --------------------------------------------------------------------------- #
#  打印工具
# --------------------------------------------------------------------------- #
def banner(title: str) -> None:
    print(f"\n{LINE}\n  {title}\n{LINE}")


def section(title: str) -> None:
    print(f"\n{SUB}\n  {title}\n{SUB}")


def console_log(text: str) -> None:
    """pipeline 的运行日志 → 控制台（GUI 版改成写日志框）。"""
    print(f"  {text}")


def print_versions() -> None:
    banner("阶段一 · 环境自检")
    print(f"  Python      : {platform.python_version()}  ({sys.executable})")
    print(f"  操作系统    : {platform.system()} {platform.release()}")
    print(f"  pandas      : {pd.__version__}")
    for mod, hint in (("openpyxl", "pip install openpyxl"),
                      ("rapidfuzz", "pip install rapidfuzz"),
                      ("tqdm", "pip install tqdm")):
        try:
            m = __import__(mod)
            print(f"  {mod:<12}: {getattr(m, '__version__', '已安装')}")
        except ImportError:
            print(f"  {mod:<12}: 【未安装】{hint}")


def print_table_info(df: pd.DataFrame, name_col: str) -> None:
    label = df.attrs.get("label", "表")
    section(f"{label} · 基础信息")
    print(f"  文件        : {df.attrs.get('path')}")
    print(f"  Sheet       : {df.attrs.get('sheet')}")
    print(f"  总行数      : {len(df)}      总列数 : {len(df.columns)}")
    print(f"  核心对比列  : {name_col}")
    print(f"  列名        : {list(df.columns)}")
    missing = df.attrs.get("missing_cols") or []
    print(f"  {'[!] 缺少期望字段 : ' + str(missing) if missing else '[OK] 期望字段校验通过，无缺失。'}")


def print_raw_names(df: pd.DataFrame, name_col: str, n: int) -> None:
    label = df.attrs.get("label", "表")
    section(f"{label} · 前 {n} 条「{name_col}」原始字符串对照表")
    for i, raw in enumerate(raw_names(df, name_col)[:n], start=1):
        info = unicode_report(raw)
        print(f"\n  [{label} #{i}] 原始值 : {info['text']}")
        print(f"        repr   : {info['repr']}")
        print(f"        码点   : {info['codepoints']}")
        flag = "首尾存在空白" if info["has_leading_or_trailing_space"] else ""
        if info["invisibles"]:
            flag += "；" + "、".join(
                f"位置{p['pos']}→{p['meaning']}" for p in info["invisibles"]
            )
        print(f"        [!] 清洗关注点 : {flag}" if flag else "        [OK] 无可疑隐形字符")


def print_sheet_inventory() -> None:
    section("工作簿 Sheet 清单")
    for path in (config.A_FILE, config.B_FILE):
        try:
            print(f"  {path.name}\n      -> {list_sheets(path)}")
        except DataLoadError as exc:
            print(f"  {path.name}\n      -> 读取失败：{exc}")


# --------------------------------------------------------------------------- #
#  第二阶段：清洗与匹配
# --------------------------------------------------------------------------- #
def print_clean_effect(df: pd.DataFrame, name_col: str, n: int = 4) -> None:
    """打印若干"原始 → 清洗后"对照，证明清洗生效且原始列未被改动。"""
    clean_col = f"{name_col}{pp.CLEAN_SUFFIX}"
    core_col = f"{name_col}{pp.CORE_SUFFIX}"
    label = df.attrs.get("label", "表")
    section(f"{label} ·「{name_col}」清洗对照（原始列未被修改）")

    changed = df[df[name_col].astype(str) != df[clean_col].astype(str)]
    print(f"  原始与清洗后存在差异的记录：{len(changed)} / {len(df)} 条")
    print(f"  新增辅助列：{clean_col}（主比对）、{core_col}（去标点核心字号）")

    for _, row in changed.head(n).iterrows():
        print(f"\n    原始   : {row[name_col]!r}")
        print(f"    清洗后 : {row[clean_col]!r}")
        print(f"    去标点 : {row[core_col]!r}")
    if changed.empty:
        print("    （本表无需清洗的差异项）")


def print_match_summary(outcome: matcher.MatchOutcome, elapsed: float,
                        thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
                        algo: str = matcher.DEFAULT_ALGO) -> None:
    st = outcome.stats
    banner("阶段二 · 清洗与匹配引擎 运行结果")
    print(f"  比对规模    : A {st['n_a']} 行 × B {st['n_b']} 行 = {st['pairs']:,} 对")
    print(f"  耗时        : {elapsed:.2f} 秒")
    print(f"  算法        : {matcher.describe_algo(algo)}")
    print(f"                （另加字号差异惩罚、通用词折叠、关键字号保护、"
          f"简称/全称包含关系修正 —— 换算法不换判据）")
    print(f"  相似度 100% : {st['exact_100']} 条")
    print(f"  互为最优    : {st['mutual_best']} 条")
    print(f"  目标争抢    : {st['conflict_targets']} 条 B 记录被多条 A 同时选为最佳"
          f"（涉及 {st['conflict_rows']} 条 A 记录）")
    print(f"  一对一让位  : {st.get('displaced', 0)} 条 A 记录的首选被更高分记录认领，"
          f"已改用其他候选；最终认领 B 记录 {st.get('b_claimed', 0)} 条")

    counts = matcher.summarize(outcome.results, thresholds)
    total = sum(v for k, v in counts.items() if k != "A系统独有")
    section(f"五档分布（阈值：{thresholds.label}）")
    for k, v in counts.items():
        pct = v / st["n_a"] * 100 if st["n_a"] else 0
        bar = "█" * int(round(pct / 2.5))
        print(f"    {k:<10} {v:>4} 条  {pct:>5.1f}%  {bar}")
    print(f"    {'—'*46}")
    print(f"    {'参与匹配合计':<10} {total:>4} 条  "
          f"（A系统独有 {counts['A系统独有']} 条需第三阶段单列）")


def print_samples(results: list[matcher.MatchResult]) -> None:
    section("人工核对样本 · 相似度 70%–95% 区间随机 3 条")
    print("  格式：A表原名 -> B表原名 (得分)  ← 请核对是否确实是同一家主体\n")
    for k, r in enumerate(results, start=1):
        print(f"  [{k}] {r.a_name}")
        print(f"      -> {r.b_name}   ({r.score:.1f}%)")
        print(f"      差异原因 : {'；'.join(pp.diff_profile(r.a_name, r.b_name))}")
        print(f"      分项得分 : ratio={r.detail['ratio']:.1f} "
              f"partial={r.detail['partial_ratio']:.1f} "
              f"WRatio={r.detail['WRatio']:.1f} "
              f"core={r.detail['core_ratio']:.1f}")
        print(f"      次佳候选 : {r.runner_up_name} ({r.runner_up_score:.1f}%)"
              f"   互为最优={r.mutual_best}   被{r.conflict_count}条A争抢")
        print()


def print_report_summary(report: dict) -> None:
    """第三阶段：打印成果文件位置与 4 个 Sheet 的行数校验。"""
    banner("第三阶段 · 成果报表导出完成")
    path: Path = report["path"]

    rel = path.relative_to(config.BASE_DIR) if path.is_relative_to(config.BASE_DIR) else path
    print(f"  报表已成功生成至 {rel.as_posix()}")
    print(f"  绝对路径 : {path}")
    print(f"  文件大小 : {path.stat().st_size:,} 字节")

    section("Sheet 行数校验")
    expect = {
        exporter.SHEET_RESULT: "匹配上的记录（完全/高度/中低/低置信度）",
        exporter.SHEET_A_ONLY: "B 系统无达标匹配的 A 记录",
        exporter.SHEET_B_ONLY: "A 系统未认领的 B 记录",
        exporter.SHEET_SUMMARY: "指标行数",
    }
    for name, rows in report["sheets"].items():
        unit = "行数据" if name != exporter.SHEET_SUMMARY else "项指标"
        print(f"    {name:<14} {rows:>4} {unit}   （{expect.get(name, '')}）")

    s1 = report["frames"]["sheet1"]
    if not s1.empty:
        print(f"\n  Sheet1 匹配状态分布：")
        for k, v in s1["匹配状态"].value_counts().items():
            print(f"      {k:<12} {v:>4} 条")
    # 表头配色 / 字号直接从 exporter 常量读，避免此处再写一份而悄悄脱节
    fill = (exporter.HEADER_FILL.fgColor.rgb or "")[-6:]
    print(f"\n  表头样式 : 加粗 / 居中 / 蓝底 {fill} / 白字 / "
          f"微软雅黑 {exporter.HEADER_SIZE}pt"
          f"（跟随官方模板；数据区 {exporter.BODY_SIZE}pt）")
    print(f"  条件格式 : 完全·高度匹配=浅绿底，中低匹配=浅黄底，独有未匹配=红色字体")
    print(f"  其他     : 冻结首行 + 自动筛选（4 个 Sheet）+ 金额 #,##0.00 + "
          f"相似度 0.0 + 占比 0.0%")


def print_conflicts(outcome: matcher.MatchOutcome, n: int = 3) -> None:
    """争抢同一 B 记录的 A 记录 —— 潜在误匹配，需人工判别。"""
    busy = sorted(
        ((j, a_list) for j, a_list in outcome.b_taken.items() if len(a_list) > 1),
        key=lambda kv: -len(kv[1]),
    )
    if not busy:
        print("  [OK] 无 A 记录争抢同一 B 记录。")
        return
    section(f"潜在误匹配提示 · {len(busy)} 条 B 记录被争抢（显示前 {n} 组）")
    reasons = {r.a_index: r for r in outcome.results}
    for j, a_list in busy[:n]:
        b_name = reasons[a_list[0]].b_name
        print(f"\n  B 记录 : {b_name}")
        for ai in a_list:
            r = reasons[ai]
            print(f"      ← A[{ai:>3}] {r.a_name:<34} {r.score:6.1f}%"
                  f"   次佳 {r.runner_up_score:5.1f}%")


# --------------------------------------------------------------------------- #
#  入口
# --------------------------------------------------------------------------- #
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Topic03 多源运营数据模糊匹配工具（读取 → 清洗 → 匹配 → 导出）"
    )
    parser.add_argument("-n", "--preview", type=int, default=config.PREVIEW_ROWS,
                        help=f"预览的原始名称条数（默认 {config.PREVIEW_ROWS}）")
    parser.add_argument("--no-progress", action="store_true", help="关闭进度条")
    parser.add_argument("--workers", type=int, default=-1,
                        help="rapidfuzz 并行度，-1 = 使用全部 CPU 核心")
    parser.add_argument("--team", default=config.DEFAULT_TEAM,
                        help=f"输出文件名中的队名/姓名（默认 {config.DEFAULT_TEAM}）")
    parser.add_argument("--out-dir", type=Path, default=None,
                        help=f"成果文件输出目录（默认 {config.OUTPUT_DIR}）")
    parser.add_argument("--no-export", action="store_true", help="只跑匹配，不导出 Excel")
    parser.add_argument("-A", "--a-file", type=Path, default=None,
                        help=f"A 系统数据文件（默认 {config.A_FILE}）")
    parser.add_argument("-B", "--b-file", type=Path, default=None,
                        help=f"B 系统数据文件（默认 {config.B_FILE}）")

    g = parser.add_argument_group("分档阈值（可调）")
    g.add_argument("--high", type=float, default=90.0,
                   help="高度匹配阈值，默认 90（范围 80–100）")
    g.add_argument("--low", type=float, default=70.0,
                   help="中低匹配分界线，默认 70（范围 50–80）")
    g.add_argument("--floor", type=float, default=60.0,
                   help="匹配识别下限，低于此分判为独有，默认 60")

    a = parser.add_argument_group("相似度算法（§3.2 加分项）")
    a.add_argument("--algo", choices=list(matcher.ALGO_CHOICES),
                   default=matcher.DEFAULT_ALGO,
                   help="基础相似度算法；默认 %(default)s = 多算法加权组合。"
                        "其余为单算法对照模式，用于演示算法差异，"
                        "业务规则（字号惩罚 / 通用词折叠 / 关键字号保护）在任何模式下都生效")
    a.add_argument("--list-algo", action="store_true",
                   help="列出全部可选算法后退出")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.no_progress:
        matcher.tqdm = None

    if args.list_algo:
        print("可选的相似度算法：")
        for key in matcher.ALGO_CHOICES:
            mark = "（默认）" if key == matcher.DEFAULT_ALGO else ""
            print(f"  {key:<14} {matcher.describe_algo(key)}{mark}")
        return 0

    print_versions()
    print(f"  相似度算法  : {matcher.describe_algo(args.algo)}")

    thresholds = matcher.Thresholds(
        high=float(args.high), low=float(args.low), floor=float(args.floor)
    )

    print_sheet_inventory()

    # ---------------- 阶段二：读取 + 清洗（与 GUI 共用 src/pipeline.py） ----------------
    banner("阶段二 · 数据读取结果校验")
    t0 = time.perf_counter()
    try:
        df_a, df_b = pipeline.load_and_clean(
            args.a_file, args.b_file, log=console_log
        )
    except DataLoadError as exc:
        print(f"\n[读取失败] {exc}\n", file=sys.stderr)
        return 1
    t_clean = time.perf_counter()

    print_table_info(df_a, COL_A)
    print_table_info(df_b, COL_B)

    section("阶段三 · 原始字符串对照（供预处理与 Unicode 清洗参考）")
    print_raw_names(df_a, COL_A, args.preview)
    print_raw_names(df_b, COL_B, args.preview)

    print_clean_effect(df_a, COL_A)
    print_clean_effect(df_b, COL_B)

    # ---------------- 第二阶段：匹配 ----------------
    banner("第二阶段 · 匹配引擎")
    outcome = pipeline.run_match(
        df_a, df_b, thresholds, workers=args.workers, log=console_log,
        algo=args.algo,
    )
    elapsed = time.perf_counter() - t0

    print_match_summary(outcome, elapsed, thresholds, args.algo)
    print_conflicts(outcome)
    print_samples(matcher.sample_band(outcome.results, 70.0, 95.0, k=3))

    # ---------------- 第三阶段：导出规范成果报表 ----------------
    if args.no_export:
        print("\n[i] --no-export：已跳过 Excel 导出。")
        return 0
    try:
        report = exporter.export_report(
            df_a, df_b, outcome, team=args.team, out_dir=args.out_dir,
            thresholds=thresholds,
        )
    except Exception as exc:                      # noqa: BLE001 —— 导出失败不应吞掉前序结果
        print(f"\n[导出失败] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2
    print_report_summary(report)

    banner("全流程完成")
    a_set = set(df_a[f"{COL_A}{pp.CLEAN_SUFFIX}"].tolist())
    b_set = set(df_b[f"{COL_B}{pp.CLEAN_SUFFIX}"].tolist())
    print(f"  清洗耗时 {t_clean - t0:.3f}s，匹配耗时 {elapsed - (t_clean - t0):.3f}s")
    print(f"  清洗后 A∩B 精确一致 : {len(a_set & b_set)} 条"
          f"（官方文档口径「约 45 条」）")
    print(f"  相似度 100% 的最佳匹配 : {outcome.stats['exact_100']} 条")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
