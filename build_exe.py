"""一键打包脚本 —— 用 PyInstaller 把 gui_app.py 打成单文件 .exe。

用法::

    python build_exe.py                 # 单文件 + 无控制台（发行版）
    python build_exe.py --console       # 保留控制台窗口（排查启动问题用）
    python build_exe.py --onedir        # 目录模式，启动更快、体积更小
    python build_exe.py --no-verify     # 只打包，不跑 exe 自检

打包要点（每一条都踩过坑）
--------------------------
1. **matplotlib 字体** —— ``--collect-all matplotlib`` 把 ``mpl-data``（含字体表、
   后端模块、样式文件）整包带进去。少了它，打包后一画图就报
   ``ImportError: cannot import name 'FigureCanvasQTAgg'`` 或中文全部变豆腐块。
2. **PyQt6 动态库** —— PyInstaller 自带 PyQt6 的 hook，会自动带上
   ``platforms/qwindows.dll`` 等插件；额外显式声明 ``backend_qtagg`` 与
   ``QtSvg``（matplotlib 的 Qt 后端在部分版本会间接引用）。
3. **rapidfuzz 的 C 扩展** —— ``--collect-submodules rapidfuzz`` 保证
   ``rapidfuzz.fuzz`` / ``rapidfuzz.process`` 这两个 Cython 模块被收录。
4. **data / output 目录挂载** —— 这是最容易翻车的一点：

   * ``--onefile`` 运行时会把包解到 ``sys._MEIPASS``，这个目录**只读且退出即删**。
     ``src/config.py`` 因此把基准目录解析成 **exe 所在目录**（``sys.frozen`` 判断），
     所以 ``output/`` 永远写在 exe 旁边，不会写进临时目录。
   * ``团体赛赛道考题/`` 采用**外部优先**：exe 同级有同名目录就用外部的
     （赛事现场换数据不用重新打包），没有才回退到打进包内的默认考题数据。
     本脚本打包完会把该目录**复制一份到 exe 旁边**，两种情形都能跑。
5. **裁剪体积** —— 排除 tkinter / PyQt5 / IPython / pytest 等完全用不到的大块头。
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "团体赛赛道考题"
ENTRY = BASE_DIR / "gui_app.py"
DIST_DIR = BASE_DIR / "dist"
BUILD_DIR = BASE_DIR / "build"
DEFAULT_NAME = "Topic03模糊匹配工具"

# 明确排除的重型依赖（本项目一条都用不到）
EXCLUDES = (
    "tkinter", "PyQt5", "PySide2", "PySide6", "IPython", "pytest",
    "matplotlib.tests", "numpy.f2py", "scipy", "notebook", "jupyter",
    "PyQt6.QtWebEngineCore", "PyQt6.QtWebEngineWidgets", "PyQt6.Qt3DCore",
    "PyQt6.QtQuick", "PyQt6.QtQml", "PyQt6.QtMultimedia", "PyQt6.QtBluetooth",
)

HIDDEN_IMPORTS = (
    "matplotlib.backends.backend_qtagg",
    "matplotlib.backends.backend_qt",
    "PyQt6.QtSvg",
    "PyQt6.QtPrintSupport",
    "openpyxl.cell._writer",
    "pandas._libs.tslibs.base",
    "tqdm",
)


# --------------------------------------------------------------------------- #
def check_environment() -> bool:
    """确认 PyInstaller 与运行依赖都在。"""
    ok = True
    try:
        import PyInstaller                                        # noqa: F401
        from PyInstaller import __version__ as ver
        print(f"  ✓ PyInstaller {ver}")
    except ImportError:
        print("  ✗ 未安装 PyInstaller。请先执行：")
        print("      python -m pip install pyinstaller")
        ok = False

    for mod in ("pandas", "openpyxl", "rapidfuzz", "matplotlib", "PyQt6"):
        try:
            m = __import__(mod)
            print(f"  ✓ {mod} {getattr(m, '__version__', '')}".rstrip())
        except ImportError:
            print(f"  ✗ 缺少运行依赖：{mod}")
            ok = False

    if not ENTRY.exists():
        print(f"  ✗ 找不到入口文件：{ENTRY}")
        ok = False
    if not DATA_DIR.is_dir():
        print(f"  ! 未找到考题数据目录：{DATA_DIR}（将不打包默认数据）")
    return ok


def build_command(name: str, onefile: bool, console: bool) -> list[str]:
    """组装 PyInstaller 命令行。"""
    sep = os.pathsep                                   # Windows 上是 ';'
    cmd: list[str] = [
        sys.executable, "-m", "PyInstaller",
        str(ENTRY),
        "--name", name,
        "--noconfirm",
        "--clean",
        "--distpath", str(DIST_DIR),
        "--workpath", str(BUILD_DIR / "work"),
        "--specpath", str(BUILD_DIR),
        "--onefile" if onefile else "--onedir",
        "--console" if console else "--windowed",
    ]
    for mod in EXCLUDES:
        cmd += ["--exclude-module", mod]
    for mod in HIDDEN_IMPORTS:
        cmd += ["--hidden-import", mod]

    cmd += ["--collect-all", "matplotlib"]
    cmd += ["--collect-submodules", "rapidfuzz"]
    cmd += ["--collect-submodules", "openpyxl"]

    if DATA_DIR.is_dir():
        cmd += ["--add-data", f"{DATA_DIR}{sep}团体赛赛道考题"]

    # 项目根目录加入搜索路径：让 src 包能被正常发现
    cmd += ["--paths", str(BASE_DIR)]
    return cmd


def run_build(cmd: list[str]) -> bool:
    print("\n" + "=" * 78)
    print("  PyInstaller 打包中（首次约 1–3 分钟，请勿中断）…")
    print("=" * 78)
    t0 = time.perf_counter()
    proc = subprocess.run(cmd, cwd=str(BASE_DIR))
    if proc.returncode != 0:
        print(f"\n✗ 打包失败（退出码 {proc.returncode}）")
        return False
    print(f"\n✓ 打包完成，用时 {time.perf_counter() - t0:.1f}s")
    return True


def exe_path(name: str, onefile: bool) -> Path:
    return (DIST_DIR / f"{name}.exe") if onefile else (DIST_DIR / name / f"{name}.exe")


def stage_assets(target: Path, onefile: bool) -> None:
    """把 data 目录与 output 目录摆到 exe 旁边（外部优先策略）。"""
    exe_dir = target.parent if onefile else target.parent
    exe_dir.mkdir(parents=True, exist_ok=True)

    external_data = exe_dir / "团体赛赛道考题"
    if DATA_DIR.is_dir() and not external_data.exists():
        shutil.copytree(DATA_DIR, external_data)
        print(f"  ✓ 已就位外部数据目录：{external_data}")

    out = exe_dir / "output"
    out.mkdir(exist_ok=True)
    print(f"  ✓ 已就位输出目录：{out}")

    readme = exe_dir / "使用说明.txt"
    if not readme.exists():
        readme.write_text(
            "Topic03 多源运营数据模糊匹配工具\n"
            "=====================================\n\n"
            "1. 双击本目录下的 Topic03模糊匹配工具.exe 启动桌面程序。\n"
            "2. 分别点「A 系统」「B 系统」两行的「选择…」按钮，各选一个 Excel 文件。\n"
            "3. 需要换数据时，把新的 Excel 放进本目录的「团体赛赛道考题」文件夹，\n"
            "   或直接在界面里选择任意位置的文件，无需重新打包。\n"
            "4. 成果报表默认输出到本目录下的 output\\ 文件夹。\n\n"
            "命令行用法（可选，支持多文件批量）：\n"
            "  Topic03模糊匹配工具.exe --selftest        离屏自检\n"
            "  python main.py                             命令行全流程\n"
            "  python main.py --batch --input <文件夹>    批量处理多个文件\n",
            encoding="utf-8",
        )
        print(f"  ✓ 已生成使用说明：{readme}")


def verify(target: Path) -> bool:
    """跑一次 exe 自检 —— 打包后 --windowed 看不到 stdout，故写入文件再读。"""
    if not target.exists():
        print(f"  ✗ 未找到可执行文件：{target}")
        return False
    log = target.parent / "_selftest.log"
    if log.exists():
        log.unlink()

    print(f"\n  正在自检：{target.name} --selftest …")
    env = dict(os.environ, QT_QPA_PLATFORM="offscreen")
    try:
        proc = subprocess.run(
            [str(target), "--selftest", "--selftest-log", str(log)],
            cwd=str(target.parent), env=env, timeout=300,
        )
    except subprocess.TimeoutExpired:
        print("  ✗ 自检超时（>300s）")
        return False

    if not log.exists():
        print(f"  ✗ 自检未产生日志（退出码 {proc.returncode}）")
        return False
    text = log.read_text(encoding="utf-8", errors="replace")
    for line in text.splitlines():
        if "[selftest]" in line:
            print("    " + line)
    passed = "全部通过" in text and proc.returncode == 0
    print(f"  {'✓ 自检通过' if passed else '✗ 自检失败'}（退出码 {proc.returncode}）")
    return passed


# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="打包 Topic03 模糊匹配工具为 exe")
    parser.add_argument("--name", default=DEFAULT_NAME, help=f"可执行文件名（默认 {DEFAULT_NAME}）")
    parser.add_argument("--onedir", action="store_true",
                        help="目录模式（启动更快、体积更小；默认为单文件）")
    parser.add_argument("--console", action="store_true",
                        help="保留控制台窗口（排查启动崩溃时用）")
    parser.add_argument("--no-verify", action="store_true", help="打包后不跑自检")
    parser.add_argument("--clean", action="store_true", help="先删除 dist/ 与 build/")
    args = parser.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    print("=" * 78)
    print(f"  打包 Topic03 多源运营数据模糊匹配工具")
    print("=" * 78)
    print("\n[1/4] 环境检查")
    if not check_environment():
        print("\n依赖不满足，打包中止。")
        return 1

    if args.clean:
        for d in (DIST_DIR, BUILD_DIR):
            if d.exists():
                shutil.rmtree(d, ignore_errors=True)
                print(f"  ✓ 已清理 {d}")

    print("\n[2/4] 执行 PyInstaller")
    onefile = not args.onedir
    if not run_build(build_command(args.name, onefile, args.console)):
        return 2

    target = exe_path(args.name, onefile)
    print("\n[3/4] 布置运行目录")
    stage_assets(target, onefile)
    size_mb = target.stat().st_size / 1024 / 1024 if target.exists() else 0
    print(f"  ✓ 可执行文件：{target}  ({size_mb:.1f} MB)")

    if args.no_verify:
        print("\n[4/4] 已跳出自检（--no-verify）")
        return 0

    print("\n[4/4] 运行 exe 自检")
    ok = verify(target)

    print("\n" + "=" * 78)
    if ok:
        print("  打包成功 🎉")
        print("=" * 78)
        print(f"\n  直接双击运行：{target}")
        print(f"  数据目录（可替换）：{target.parent / '团体赛赛道考题'}")
        print(f"  成果输出目录：      {target.parent / 'output'}")
        return 0
    print("  打包完成，但自检未通过 —— 请查看上方日志")
    print("=" * 78)
    return 3


if __name__ == "__main__":
    raise SystemExit(main())
