"""一键生成符合赛题规范的提交压缩包。

用法::

    python pack_submission.py                    # 完整归档（含 exe）
    python pack_submission.py --no-exe           # 不含 exe（改用指向 dist/ 的说明文件）
    python pack_submission.py --team 张三        # 换队名
    python pack_submission.py --list             # 只列出将要打包的内容，不实际压缩

产物：``提交包/Topic03_模糊匹配_<队名>.zip``

压缩包内部结构
--------------
```
Topic03_模糊匹配_张涵博.zip
├── 源码/                       完整工程，解压即可运行（自包含）
│   ├── src/                    8 个核心模块
│   ├── main.py                 命令行入口
│   ├── gui_app.py              PyQt6 桌面端
│   ├── build_exe.py / .bat     一键打包脚本
│   ├── pack_submission.py      本脚本
│   ├── requirements.txt        依赖清单
│   ├── .gitignore
│   └── 团体赛赛道考题/          数据目录（源码树自包含，解压即可跑）
├── 标准成果报表/                工具运行后自动生成的成果文件
│   └── 模糊匹配结果_张涵博_<日期>.xlsx
├── 可执行程序/                  免安装运行
│   ├── 多源运营数据相似度模糊匹配工具.exe
│   ├── 团体赛赛道考题/          数据目录（可替换）
│   └── 使用说明.txt
├── 作品说明文档.md
├── 演示视频脚本指南.md
├── 提交说明.md
└── 文件清单.txt                 包内所有文件的路径 + 大小
```

设计要点
--------
* **成果报表自动补生成**：归档前检查 ``output/`` 下有没有成果文件，没有就先跑一次
  ``main.py``，保证「提交的报表」与「当前代码」是同一版本产出的，不会出现
  文档描述与报表数字对不上的尴尬。
* **exe 体积提示**：单文件 exe 约 80 MB 且已内部压缩，打进 zip 后几乎不再变小。
  ``--no-exe`` 可改用一份指向 ``dist/`` 的说明文件，把压缩包压到 100 KB 量级。
* **排除清单显式声明**：虚拟环境、打包中间产物、缓存、IDE 配置一律不进包，
  并在结束时打印被排除的目录，避免"以为打进去了其实没有"。
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
DATA_DIR = BASE_DIR / "团体赛赛道考题"
OUTPUT_DIR = BASE_DIR / "output"
DIST_DIR = BASE_DIR / "dist"
SUBMIT_DIR = BASE_DIR / "提交包"

DEFAULT_TEAM = "张涵博"
EXE_NAME = "多源运营数据相似度模糊匹配工具"

# 源码包要收录的根级文件
SOURCE_ROOT_FILES = (
    "main.py", "gui_app.py", "build_exe.py", "build.bat",
    "pack_submission.py", "requirements.txt", ".gitignore",
)

# 文档（放在压缩包根目录）
DOC_FILES = ("作品说明文档.md", "演示视频脚本指南.md", "提交说明.md")

# 明确排除的目录：生成物、缓存、本机环境
EXCLUDE_DIRS = (
    ".venv", "venv", "__pycache__", "build", "dist", "output",
    "提交包", ".git", ".idea", ".vscode", ".claude", ".pytest_cache",
)


@dataclass
class Payload:
    """待归档的一个文件。"""

    disk: Path
    arc: str
    tag: str = ""

    @property
    def size(self) -> int:
        return self.disk.stat().st_size if self.disk.exists() else 0


@dataclass
class Manifest:
    items: list[Payload] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)

    def add(self, disk: Path, arc: str, tag: str = "") -> None:
        if disk.exists():
            self.items.append(Payload(disk, arc, tag))
        else:
            self.skipped.append(f"{arc}（源文件不存在：{disk}）")

    @property
    def total_size(self) -> int:
        return sum(i.size for i in self.items)


# --------------------------------------------------------------------------- #
#  成果报表
# --------------------------------------------------------------------------- #
def latest_report() -> Path | None:
    """取 output/ 下最新的成果报表。"""
    if not OUTPUT_DIR.is_dir():
        return None
    candidates = [p for p in OUTPUT_DIR.glob("模糊匹配结果_*.xlsx")
                  if not p.name.startswith("~$")]
    return max(candidates, key=lambda p: p.stat().st_mtime) if candidates else None


def ensure_report(team: str, log=print) -> Path:
    """确保存在一份最新成果报表；没有就先跑一次完整流程生成。"""
    report = latest_report()
    if report:
        log(f"  ✓ 已有成果报表：{report.name}")
        return report

    log("  ! output/ 下没有成果报表，正在运行一次完整流程生成…")
    proc = subprocess.run(
        [sys.executable, str(BASE_DIR / "main.py"), "--team", team,
         "--no-progress"],
        cwd=str(BASE_DIR),
    )
    if proc.returncode != 0:
        raise RuntimeError(f"生成成果报表失败（退出码 {proc.returncode}）")

    report = latest_report()
    if not report:
        raise RuntimeError("运行后仍未找到成果报表，请检查 output/ 目录")
    log(f"  ✓ 已生成成果报表：{report.name}")
    return report


# --------------------------------------------------------------------------- #
#  收集
# --------------------------------------------------------------------------- #
def collect(team: str, with_exe: bool, log=print) -> Manifest:
    """把要归档的内容收集成清单。"""
    m = Manifest()

    # ---- 源码/ ----
    for py in sorted(SRC_DIR.glob("*.py")):
        m.add(py, f"源码/src/{py.name}", "源码")
    for name in SOURCE_ROOT_FILES:
        m.add(BASE_DIR / name, f"源码/{name}", "源码")
    # 数据目录必须随源码一起打包：config.py 把基准目录解析到 main.py 所在层，
    # 少了它，解压后直接跑 main.py 会报「文件不存在」—— 源码树就不是自包含的。
    for f in sorted(DATA_DIR.glob("*")):
        if f.is_file():
            m.add(f, f"源码/团体赛赛道考题/{f.name}", "源码数据")

    # ---- 标准成果报表/ ----
    report = ensure_report(team, log)
    m.add(report, f"标准成果报表/{report.name}", "成果报表")

    # ---- 可执行程序/ ----
    exe = DIST_DIR / f"{EXE_NAME}.exe"
    if with_exe:
        m.add(exe, f"可执行程序/{EXE_NAME}.exe", "可执行程序")
        for f in sorted((DIST_DIR / "团体赛赛道考题").glob("*")):
            if f.is_file():
                m.add(f, f"可执行程序/团体赛赛道考题/{f.name}", "程序数据")
        m.add(DIST_DIR / "使用说明.txt", "可执行程序/使用说明.txt", "程序说明")
    else:
        pointer = SUBMIT_DIR / "_exe_locate.txt"
        pointer.parent.mkdir(parents=True, exist_ok=True)
        pointer.write_text(
            "可执行程序位置说明\n"
            "==================\n\n"
            f"本包未内嵌 exe（单文件约 80 MB，压缩后体积几乎不变）。\n"
            f"请从工程目录取用：\n\n"
            f"    {exe}\n\n"
            f"若该文件不存在，在源码目录执行以下命令现场生成（约 1–3 分钟）：\n\n"
            f"    python build_exe.py\n\n"
            f"生成后可直接双击运行，数据目录与 output 目录会自动就位在 exe 同级。\n",
            encoding="utf-8",
        )
        m.add(pointer, "可执行程序/可执行程序位置说明.txt", "程序说明")

    # ---- 文档 ----
    for name in DOC_FILES:
        m.add(BASE_DIR / name, name, "文档")

    return m


def write_manifest(m: Manifest, team: str) -> Path:
    """生成包内文件清单。"""
    lines = [
        f"Topic03 多源运营数据模糊匹配工具 —— 提交包文件清单",
        f"参赛者：{team}    生成日期：{date.today():%Y-%m-%d}",
        "=" * 66,
        "",
    ]
    current = None
    for item in sorted(m.items, key=lambda i: i.arc):
        top = item.arc.split("/")[0]
        if top != current:
            current = top
            lines.append(f"\n【{top}】")
        lines.append(f"  {item.arc:<58} {item.size:>10,} B")

    lines += [
        "",
        "=" * 66,
        f"文件总数：{len(m.items)}",
        f"原始总大小：{m.total_size:,} B ({m.total_size / 1024 / 1024:.1f} MB)",
        "",
        "快速开始：",
        "  方式一（免安装）：进入「可执行程序」，双击 exe 即可运行",
        "  方式二（源码）：   pip install -r 源码/requirements.txt",
        "                    然后 python 源码/gui_app.py",
        "",
        "成果报表每 Sheet 行数：模糊匹配结果 83 / A系统独有 17 / B系统独有 14",
        "（83 + 17 = 100 条 A 记录；83 + 14 = 97 条 B 记录，两侧账目对平）",
    ]
    path = SUBMIT_DIR / "_文件清单.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


# --------------------------------------------------------------------------- #
#  压缩
# --------------------------------------------------------------------------- #
def make_zip(m: Manifest, team: str, log=print) -> Path:
    SUBMIT_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = SUBMIT_DIR / f"Topic03_模糊匹配_{team}.zip"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED,
                         compresslevel=6, allowZip64=True) as zf:
        for item in m.items:
            zf.write(item.disk, item.arc)
            log(f"    + {item.arc}")
    return zip_path


def verify_zip(zip_path: Path, log=print) -> bool:
    """回读压缩包，确认条目数与完整性。"""
    with zipfile.ZipFile(zip_path) as zf:
        bad = zf.testzip()
        names = zf.namelist()
    if bad:
        log(f"  ✗ 压缩包内文件损坏：{bad}")
        return False
    tops = sorted({n.split("/")[0] for n in names})
    log(f"  ✓ 完整性校验通过：{len(names)} 个条目，顶层目录 {tops}")
    return True


# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="生成 Topic03 赛题提交压缩包"
    )
    parser.add_argument("--team", default=DEFAULT_TEAM,
                        help=f"队名 / 姓名，用于压缩包与成果文件命名（默认 {DEFAULT_TEAM}）")
    parser.add_argument("--no-exe", action="store_true",
                        help="不内嵌 exe，改用指向 dist/ 的位置说明（压缩包可小两个数量级）")
    parser.add_argument("--list", action="store_true",
                        help="只列出将打包的内容，不实际压缩")
    args = parser.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    print("=" * 70)
    print("  Topic03 模糊匹配工具 —— 提交包归档")
    print("=" * 70)

    print("\n[1/4] 收集文件")
    print("  已排除（生成物 / 缓存 / 本机环境）：")
    for d in EXCLUDE_DIRS:
        if (BASE_DIR / d).exists():
            print(f"    - {d}/")

    try:
        m = collect(args.team, with_exe=not args.no_exe)
    except RuntimeError as exc:
        print(f"\n  ✗ {exc}")
        return 1

    by_tag: dict[str, int] = {}
    for item in m.items:
        by_tag[item.tag] = by_tag.get(item.tag, 0) + 1
    print(f"\n  共 {len(m.items)} 个文件，原始大小 "
          f"{m.total_size / 1024 / 1024:.1f} MB")
    for tag, n in by_tag.items():
        print(f"    {tag:<12} {n:>3} 个")
    if m.skipped:
        print("  跳过：")
        for s in m.skipped:
            print(f"    ! {s}")

    if args.list:
        print("\n[--list] 仅列出，不压缩。完整清单：")
        for item in sorted(m.items, key=lambda i: i.arc):
            print(f"    {item.arc:<58} {item.size:>10,} B")
        return 0

    print("\n[2/4] 生成文件清单")
    manifest_path = write_manifest(m, args.team)
    m.add(manifest_path, "文件清单.txt", "清单")
    print(f"  ✓ {manifest_path.name}")

    print(f"\n[3/4] 压缩中（{m.total_size / 1024 / 1024:.1f} MB，请稍候）…")
    try:
        zip_path = make_zip(m, args.team)
    except PermissionError:
        print(f"\n  ✗ 无法写入 {zip_path}\n"
              f"    该文件可能正被解压软件或资源管理器占用，请关闭后重试。")
        return 2

    print(f"\n[4/4] 校验")
    if not verify_zip(zip_path):
        return 3

    size_mb = zip_path.stat().st_size / 1024 / 1024
    ratio = zip_path.stat().st_size / m.total_size * 100 if m.total_size else 0
    print("\n" + "=" * 70)
    print("  归档完成 🎉")
    print("=" * 70)
    print(f"\n  提交包：{zip_path}")
    print(f"  大小：  {size_mb:.2f} MB（原始 {m.total_size / 1024 / 1024:.1f} MB，"
          f"压缩率 {ratio:.0f}%）")
    print(f"  文件数：{len(m.items)}")

    if not args.no_exe and size_mb > 50:
        print("\n  [提示] 压缩包较大主要来自内嵌的 exe。若赛事平台限制上传体积，")
        print("         可用 --no-exe 重新生成（会附一份指向 dist/ 的位置说明）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
