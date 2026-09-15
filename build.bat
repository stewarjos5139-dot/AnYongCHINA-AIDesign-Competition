@echo off
REM ============================================================
REM  Topic03 多源运营数据模糊匹配工具 —— Windows 一键打包
REM  双击本文件即可打包成 dist\Topic03模糊匹配工具.exe
REM ============================================================
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo.
echo ============================================================
echo   正在检查 Python 环境...
echo ============================================================

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 python 命令。
    echo        请先安装 Python 3.10+ 并勾选 "Add Python to PATH"。
    pause
    exit /b 1
)

python -c "import pandas, openpyxl, rapidfuzz, matplotlib, PyQt6" 2>nul
if errorlevel 1 (
    echo [提示] 运行依赖不完整，正在安装...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请手工执行： python -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [提示] 正在安装 PyInstaller...
    python -m pip install pyinstaller
)

echo.
echo ============================================================
echo   开始打包（首次约 1-3 分钟，请勿关闭窗口）
echo ============================================================
echo.

python build_exe.py %*

if errorlevel 1 (
    echo.
    echo [失败] 打包未通过，请查看上方日志。
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   打包完成！可执行文件在 dist\ 目录下。
echo ============================================================
echo.
pause
