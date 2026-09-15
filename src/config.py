"""全局配置：路径、Sheet 名、核心对比列、校验用字段清单。

后续阶段（预处理 / 相似度 / 导出）共用本文件，避免路径与列名散落在各处。

打包（PyInstaller）说明
-----------------------
``--onefile`` 模式下 ``sys._MEIPASS`` 指向一个**只读且退出即删**的临时解包目录，
因此绝不能拿它当基准目录 —— 否则 ``output/`` 会写到一个下次运行就消失的地方。
:func:`_resolve_base_dir` 统一处理这件事：

* 普通运行 → 项目根目录（``src/`` 的上一级）
* 打包运行 → ``gui_app.exe`` 所在目录

同时 ``团体赛赛道考题/`` 采用"**外部优先**"策略：exe 同级目录有就用外部的
（赛事现场换数据不用重新打包），没有就回退到包内的那份默认考题数据。
"""

from __future__ import annotations

import sys
from pathlib import Path


def _resolve_base_dir() -> Path:
    """基准目录：打包后取 exe 所在目录，否则取项目根目录。"""
    if getattr(sys, "frozen", False):            # PyInstaller 冻结标记
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _resolve_data_dir(base: Path) -> Path:
    """数据目录：exe 同级的外部目录优先，回退到打包进包内的默认数据。"""
    external = base / "团体赛赛道考题"
    if external.is_dir():
        return external
    bundled = Path(getattr(sys, "_MEIPASS", base)) / "团体赛赛道考题"
    return bundled if bundled.is_dir() else external


BASE_DIR: Path = _resolve_base_dir()
DATA_DIR: Path = _resolve_data_dir(BASE_DIR)

A_FILE: Path = DATA_DIR / "【考题材料】客户交易明细_A系统.xlsx"
B_FILE: Path = DATA_DIR / "【考题材料】银行流水明细_B系统.xlsx"
TEMPLATE_FILE: Path = DATA_DIR / "【考题材料】输出成果模板.xlsx"

# ---------------------------------------------------------------- Sheet
A_SHEET: str = "客户交易明细_A系统"
B_SHEET: str = "银行流水明细_B系统"

# ---------------------------------------------------------------- 核心对比列
A_NAME_COL: str = "客户名称"      # A 系统（ERP）标准全称
B_NAME_COL: str = "对方户名"      # B 系统（银行流水）非标户名
A_KEY_COL: str = "客户编码"
B_KEY_COL: str = "流水号"
B_ACCOUNT_COL: str = "对方账号"

# ---------------------------------------------------------------- 金额 / 日期（两表同名）
A_AMOUNT_COL: str = "交易金额（元）"
B_AMOUNT_COL: str = "交易金额（元）"
DATE_COL: str = "交易日期"

# ---------------------------------------------------------------- 成果输出
OUTPUT_DIR: Path = BASE_DIR / "output"
DEFAULT_TEAM: str = "张涵博"       # 输出文件名中的队名/姓名，可用 --team 覆盖

# ---------------------------------------------------------------- 期望字段（缺失时告警，不中断）
EXPECTED_A_COLS: tuple[str, ...] = (
    "序号", "客户编码", "客户名称", "交易金额（元）", "交易日期", "业务类型", "部门",
)
EXPECTED_B_COLS: tuple[str, ...] = (
    "序号", "对方户名", "对方账号", "交易金额（元）", "交易日期", "流水号", "摘要",
)

# ---------------------------------------------------------------- 展示参数
PREVIEW_ROWS: int = 3             # 控制台预览的原始名称条数
