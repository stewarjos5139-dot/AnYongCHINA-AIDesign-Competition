"""多文件 / 目录批量处理。

能力
----
* **输入**：任意数量的 Excel 文件，或一个包含多个 Excel 的文件夹
* **角色识别**：自动判断每个文件是 A 系统（ERP 客户明细）还是 B 系统（银行流水），
  依次尝试「Sheet 名 → 列名 → 文件名」三级特征，识别不出会明确报错而不是瞎猜
* **配对**：单边只有 1 个文件时与另一侧全交叉；否则按文件名相似度贪心一对一，
  落单的文件配给它最相似的对侧文件
* **输出**：每个配对生成一份与官方模板一致的 4-Sheet 成果文件；配对多于 1 个时
  额外生成一份《批量汇总》，含「数据来源文件（A/B）」列，便于跨批次汇总核对

设计上刻意与 :mod:`src.pipeline` 解耦：本模块只负责"选哪些文件、怎么配对、
结果怎么落盘"，单个配对的计算流程完全复用 ``pipeline.run_pipeline``，
保证批量模式与单文件模式的成果文件逐字节一致。
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

import pandas as pd
from rapidfuzz import fuzz

from . import config as C
from . import exporter, matcher, pipeline

ProgressFn = Callable[[int, str], None]
LogFn = Callable[[str], None]

ROLE_A = "A"
ROLE_B = "B"

EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xls")

# 三级识别特征
SHEET_HINTS = {
    ROLE_A: ("客户交易明细", "A系统", "客户明细", "ERP"),
    ROLE_B: ("银行流水明细", "B系统", "银行流水", "流水明细"),
}
COLUMN_HINTS = {
    ROLE_A: ("客户名称", "客户编码"),
    ROLE_B: ("对方户名", "对方账号", "流水号"),
}
FILE_HINTS = {
    ROLE_A: ("客户交易明细", "客户明细", "a系统", "_a", "-a", "erp", "客户"),
    ROLE_B: ("银行流水", "流水明细", "b系统", "_b", "-b", "银行", "流水"),
}

# 配对用的文件名噪声词（去掉后剩下的才是"这一批"的区分特征，通常是年份/期间）
_NOISE_WORDS = (
    "【考题材料】", "考题材料", "客户交易明细", "银行流水明细", "交易明细", "流水明细",
    "客户明细", "银行流水", "明细", "a系统", "b系统", "系统", "客户", "银行",
)


# --------------------------------------------------------------------------- #
#  数据结构
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Job:
    """一个 (A文件, B文件) 配对任务。"""

    a_file: Path
    b_file: Path
    label: str

    def __str__(self) -> str:
        return f"{self.a_file.name}  ×  {self.b_file.name}"


@dataclass
class BatchOutcome:
    """批量执行结果。"""

    pairs: list[tuple[Job, pipeline.PipelineResult]] = field(default_factory=list)
    failures: list[tuple[Job, str]] = field(default_factory=list)
    overview_path: Path | None = None
    elapsed: float = 0.0

    @property
    def n_ok(self) -> int:
        return len(self.pairs)

    @property
    def entries(self) -> list[exporter.BatchEntry]:
        return [_to_entry(job, res) for job, res in self.pairs]

    @property
    def total_a(self) -> int:
        return sum(len(r.df_a) for _, r in self.pairs)

    @property
    def total_b(self) -> int:
        return sum(len(r.df_b) for _, r in self.pairs)


def _to_entry(job: Job, res: pipeline.PipelineResult) -> exporter.BatchEntry:
    frames = res.report["frames"] if res.report else {}
    return exporter.BatchEntry(
        label=job.label,
        a_file=job.a_file,
        b_file=job.b_file,
        n_a=len(res.df_a),
        n_b=len(res.df_b),
        counts=res.counts,
        sheet1=frames.get("sheet1", pd.DataFrame()),
        sheet2=frames.get("sheet2", pd.DataFrame()),
        sheet3=frames.get("sheet3", pd.DataFrame()),
        thresholds=res.thresholds,
    )


# --------------------------------------------------------------------------- #
#  角色识别
# --------------------------------------------------------------------------- #
def detect_role(path: Path | str) -> str | None:
    """判断文件属于 A 系统还是 B 系统；无法判断返回 ``None``。

    三级特征依次尝试，任一级命中即返回：

    1. **Sheet 名**（最可靠）—— 含「客户交易明细」→ A，含「银行流水明细」→ B
    2. **列名** —— 含「客户名称 / 客户编码」→ A，含「对方户名 / 对方账号」→ B
    3. **文件名** —— 含「客户 / a系统 / erp」→ A，含「银行 / 流水 / b系统」→ B
    """
    path = Path(path)
    if not path.exists():
        return None

    # 1) Sheet 名
    try:
        sheets = pipeline.available_sheets(path)
    except Exception:                                   # noqa: BLE001
        sheets = []
    joined = " ".join(sheets)
    for role, hints in SHEET_HINTS.items():
        if any(h in joined for h in hints):
            return role

    # 2) 列名（只读表头，nrows=0 不加载数据）
    try:
        head = pd.read_excel(path, sheet_name=0, nrows=0, engine="openpyxl")
        cols = " ".join(str(c) for c in head.columns)
    except Exception:                                   # noqa: BLE001
        cols = ""
    for role, hints in COLUMN_HINTS.items():
        if any(h in cols for h in hints):
            return role

    # 3) 文件名
    name = path.stem.lower()
    for role, hints in FILE_HINTS.items():
        if any(h in name for h in hints):
            return role
    return None


def expand_inputs(paths: Iterable[Path | str]) -> list[Path]:
    """目录展开成其中的 Excel 文件；文件原样保留。去重并排序。"""
    found: list[Path] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            for f in sorted(p.iterdir()):
                if (f.suffix.lower() in EXCEL_SUFFIXES
                        and not f.name.startswith("~$")            # Excel 临时文件
                        and not f.name.startswith(".")):
                    found.append(f)
        elif p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES:
            found.append(p)
    seen, out = set(), []
    for f in found:
        key = f.resolve()
        if key not in seen:
            seen.add(key)
            out.append(f)
    return out


# --------------------------------------------------------------------------- #
#  配对
# --------------------------------------------------------------------------- #
def _stem_key(path: Path) -> str:
    """去掉通用噪声词后的文件名主干 —— 剩下的通常是年份 / 期间，用于配对。"""
    s = path.stem.lower()
    for w in _NOISE_WORDS:
        s = s.replace(w.lower(), "")
    return re.sub(r"[\s_\-—－()（）\[\]【】]+", "", s)


def plan_jobs(paths: Sequence[Path | str]) -> list[Job]:
    """把一批文件规划成 (A, B) 配对任务。

    规则：

    * 单边只有 1 个文件 → 与另一侧**全交叉**（1 个 A 对 N 个 B，或反过来）
    * 两侧都多于 1 个 → 按文件名主干相似度**贪心一对一**
    * 落单的文件 → 配给它最相似的对侧文件（保证每个文件都被处理到）
    """
    files = expand_inputs(paths)
    if not files:
        raise ValueError("没有找到任何 Excel 文件（支持 .xlsx / .xlsm / .xls）")

    a_files, b_files, unknown = [], [], []
    for f in files:
        role = detect_role(f)
        (a_files if role == ROLE_A else b_files if role == ROLE_B else unknown).append(f)

    if not a_files or not b_files:
        detail = "\n".join(
            f"    · {f.name} → {'识别不出' if f in unknown else '已识别'}"
            for f in files
        )
        raise ValueError(
            f"无法配对：识别到 A 系统文件 {len(a_files)} 个、B 系统文件 {len(b_files)} 个。\n"
            f"  A 类特征：Sheet 名含「客户交易明细」，或列名含「客户名称/客户编码」\n"
            f"  B 类特征：Sheet 名含「银行流水明细」，或列名含「对方户名/对方账号」\n"
            f"  已选文件：\n{detail}"
        )

    if unknown:
        # 单边文件足够多时，未识别的文件忽略即可；否则报错更安全
        pass

    pairs: list[tuple[Path, Path]] = []
    if len(a_files) == 1:
        pairs = [(a_files[0], b) for b in b_files]
    elif len(b_files) == 1:
        pairs = [(a, b_files[0]) for a in a_files]
    else:
        remaining = list(b_files)
        for a in a_files:
            if not remaining:
                break
            key = _stem_key(a)
            best = max(remaining, key=lambda b: fuzz.ratio(key, _stem_key(b)))
            remaining.remove(best)
            pairs.append((a, best))
        # 多出来的 B 各自配给它最像的 A
        for b in remaining:
            key = _stem_key(b)
            best_a = max(a_files, key=lambda a: fuzz.ratio(key, _stem_key(a)))
            pairs.append((best_a, b))

    used: dict[str, int] = {}
    jobs: list[Job] = []
    for a, b in pairs:
        base = re.sub(r"[\\/:*?\"<>|]", "_", a.stem)[:60]
        used[base] = used.get(base, 0) + 1
        label = base if used[base] == 1 else f"{base}_{used[base]}"
        jobs.append(Job(a_file=a, b_file=b, label=label))
    return jobs


# --------------------------------------------------------------------------- #
#  执行
# --------------------------------------------------------------------------- #
def run_batch(
    paths: Sequence[Path | str] | None = None,
    thresholds: matcher.Thresholds = matcher.DEFAULT_THRESHOLDS,
    *,
    team: str = C.DEFAULT_TEAM,
    out_dir: Path | str | None = None,
    export: bool = True,
    workers: int = -1,
    tie_break: bool = True,
    progress: ProgressFn | None = None,
    log: LogFn | None = None,
) -> BatchOutcome:
    """批量跑完所有配对，返回 :class:`BatchOutcome`。

    进度与日志按"任务"维度汇总上报：整体进度 = (已完成任务数 + 当前任务内进度) / 总任务数。
    """
    log = log or (lambda _t: None)
    jobs = plan_jobs(paths or [C.A_FILE, C.B_FILE])
    total_jobs = len(jobs)
    out_dir = Path(out_dir) if out_dir else C.OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    log(f"[批量] 共规划 {total_jobs} 个配对任务")
    for i, j in enumerate(jobs, start=1):
        log(f"[批量]   任务 {i}/{total_jobs}：{j.a_file.name} × {j.b_file.name}")

    outcome = BatchOutcome()
    t0 = time.perf_counter()
    emit = progress or (lambda _p, _m: None)

    for idx, job in enumerate(jobs):
        base = idx / total_jobs
        span = 1.0 / total_jobs

        def inner(pct: int, message: str, _b=base, _s=span, _i=idx) -> None:
            overall = int(round((_b + _s * pct / 100.0) * 100))
            emit(min(overall, 100), f"任务 {_i + 1}/{total_jobs}：{message}")

        emit(int(base * 100), f"任务 {idx + 1}/{total_jobs}：读取与清洗")
        # 单任务时沿用标准文件名，保证与单文件模式的成果文件完全同名同内容
        filename = _job_filename(job, team) if total_jobs > 1 else None
        try:
            result = pipeline.run_pipeline(
                a_file=job.a_file,
                b_file=job.b_file,
                thresholds=thresholds,
                team=team,
                out_dir=out_dir,
                filename=filename,
                export=export,
                workers=workers,
                tie_break=tie_break,
                progress=inner,
                log=lambda t, _i=idx: log(f"[任务{_i + 1}] {t}"),
            )
        except Exception as exc:                        # noqa: BLE001 —— 单个任务失败不中断整批
            msg = f"{type(exc).__name__}: {exc}"
            outcome.failures.append((job, msg))
            log(f"[任务{idx + 1}] ✗ 失败：{msg}")
            continue

        outcome.pairs.append((job, result))
        log(f"[任务{idx + 1}] ✓ 完成 → {result.path.name if result.path else '（未导出）'}")

    outcome.elapsed = time.perf_counter() - t0

    # 多于一个任务时才生成批量总览（单任务时与普通成果文件重复）
    if export and len(outcome.pairs) > 1:
        emit(98, "生成批量汇总…")
        outcome.overview_path = exporter.export_batch_overview(
            outcome.entries, team=team, out_dir=out_dir
        )
        log(f"[批量] 批量汇总 → {outcome.overview_path.name}")

    emit(100, "批量处理完成")
    log(f"[批量] 成功 {outcome.n_ok}/{total_jobs} 个任务，"
        f"总耗时 {outcome.elapsed:.2f}s")
    return outcome


def _job_filename(job: Job, team: str, when=None) -> str:
    """单任务成果文件名：在标准命名后追加任务标签，避免多任务互相覆盖。"""
    import datetime as _dt

    when = when or _dt.date.today()
    return f"模糊匹配结果_{team}_{when:%Y%m%d}_{job.label}.xlsx"
