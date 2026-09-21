"""Topic03 多源运营数据模糊匹配工具 —— PyQt 桌面端界面（加分项）。

启动：
    python gui_app.py

自检（无需显示器，用于 CI / 快速验证界面能否装配）：
    python gui_app.py --selftest

设计要点
--------
**多线程防阻塞**
    所有耗时计算（读取 → 清洗 → 4 个相似度矩阵 → 字号校验 → 选优 → 导出 Excel）
    全部在 :class:`MatchWorker`（``QThread`` 子类）里执行，主线程只负责刷新界面。
    Worker 通过 ``pyqtSignal`` 向外广播 进度 / 日志 / 成功 / 失败 四类信号，
    界面元件永远只在主线程被读写。

**UI 与业务解耦**
    窗口类不直接调用 matcher / exporter，只调用 ``src.pipeline.run_pipeline``；
    换言之命令行与桌面端跑的是同一份编排逻辑，产出文件逐字节一致。

**健壮性**
    阈值用 ``QSpinBox`` 约束在合法区间（高度 80–100、中低 50–80），
    并在 start 前断言 ``100 ≥ high ≥ low ≥ 60``，非法组合直接弹窗拦截。
"""

from __future__ import annotations

import argparse
import sys
import traceback
from pathlib import Path

# --------------------------------------------------------------------------- #
#  PyQt6 / PyQt5 兼容导入
# --------------------------------------------------------------------------- #
try:
    from PyQt6.QtCore import Qt, QThread, pyqtSignal
    from PyQt6.QtGui import QColor, QFont
    from PyQt6.QtWidgets import (
        QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QGroupBox,
        QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
        QProgressBar, QPushButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter,
        QStatusBar, QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit,
        QVBoxLayout, QWidget,
    )

    QT6 = True
except ImportError:                                     # pragma: no cover
    from PyQt5.QtCore import Qt, QThread, pyqtSignal  # type: ignore
    from PyQt5.QtGui import QColor, QFont             # type: ignore
    from PyQt5.QtWidgets import (                     # type: ignore
        QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QGroupBox,
        QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
        QProgressBar, QPushButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter,
        QStatusBar, QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit,
        QVBoxLayout, QWidget,
    )

    QT6 = False

# matplotlib 内嵌：用 FigureCanvasQTAgg，不用 pyplot（避免与 Qt 事件循环冲突）
import matplotlib

matplotlib.use("QtAgg" if QT6 else "Qt5Agg")

try:
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
except ImportError:                                     # pragma: no cover
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg  # type: ignore

from matplotlib.figure import Figure                              # noqa: E402

matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
matplotlib.rcParams["axes.unicode_minus"] = False                 # 负号正常显示

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src import config as C            # noqa: E402
from src import matcher, pipeline      # noqa: E402

# --------------------------------------------------------------------------- #
#  配色
# --------------------------------------------------------------------------- #
# 统一配色：图表与表格共用同一套色相，避免同一档位在两个视图里颜色对不上。
# 色相沿"置信度"递降排列：绿 → 青 → 琥珀 → 橙 → 红，另加紫色标识 B 系统独有。
BUCKET_COLORS: dict[str, str] = {
    "完全匹配": "#2E8B57",       # 绿
    "高度匹配": "#4A9DB5",       # 青
    "中低匹配": "#E5A800",       # 琥珀
    "低置信度匹配": "#E8730C",   # 橙
    "A系统独有": "#C0392B",      # 红
    "B系统独有": "#7030A0",      # 紫
}

# 表格逐档配色：``状态 -> (整行浅色底, 状态格实色底, 状态格文字色)``
# 整行浅色给"一眼扫过"的档位感，状态格实色徽标给"逐行确认"的强区分。
STATUS_STYLES: dict[str, tuple[str, str, str]] = {
    "完全匹配":     ("#D6F0DF", "#2E8B57", "#FFFFFF"),
    "高度匹配":     ("#DCEEF5", "#4A9DB5", "#FFFFFF"),
    "中低匹配":     ("#FDF2CC", "#E5A800", "#FFFFFF"),
    "低置信度匹配": ("#FCE3CE", "#E8730C", "#FFFFFF"),
}
# 低置信度的备注用深橙红，既醒目又不与行底色糊在一起
REMARK_FG = {"低置信度匹配": "#8A3B12"}
REMARK_FG_DEFAULT = "#5A6472"

# 对账全局视角的三分色（环形图用），与上面的档位色保持同一色系
GLOBAL_COLORS: dict[str, str] = {
    "成功匹配": "#2E8B57",
    "A系统独有": "#C0392B",
    "B系统独有": "#7030A0",
}

# 固定展示顺序（与 matcher.summarize 的键序一致），条形图 0 值档位也保留
BUCKET_ORDER: tuple[str, ...] = (
    "完全匹配", "高度匹配", "中低匹配", "低置信度匹配", "A系统独有",
)

WINDOW_SIZE = (1440, 900)
WINDOW_MIN_SIZE = (1360, 850)

# --------------------------------------------------------------------------- #
#  全局 QSS 主题
# --------------------------------------------------------------------------- #
# 统一在这里定义视觉规范，而不是把 setStyleSheet 散落到每个控件上 ——
# 改一处就能整体换肤，也避免各控件的行内样式互相打架。
APP_QSS = """
/* ---------- 基础 ---------- */
QWidget {
    font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
    font-size: 13px;
    color: #1F2937;
}
QMainWindow, QDialog { background: #F4F6F9; }

/* ---------- 分组卡片 ---------- */
QGroupBox {
    background: #FFFFFF;
    border: 1px solid #DDE3EC;
    border-radius: 8px;
    margin-top: 11px;
    padding: 4px 4px 4px 4px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: #1F3864;
    font-weight: bold;
}

/* ---------- 按钮 ---------- */
QPushButton {
    background: #FFFFFF;
    border: 1px solid #C9D2E0;
    border-radius: 6px;
    min-height: 30px;
    padding: 2px 14px;
    color: #1F3864;
}
QPushButton:hover   { background: #EDF3FB; border-color: #7EA6D8; }
QPushButton:pressed { background: #DCE8F7; }
QPushButton:disabled {
    background: #F1F3F7; color: #A6AEB9; border-color: #E5E9F0;
}

/* 主操作按钮：深蓝实心 */
QPushButton#PrimaryButton {
    background: #004080;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    font-size: 15px;
    font-weight: bold;
}
QPushButton#PrimaryButton:hover    { background: #0A5AA8; }
QPushButton#PrimaryButton:pressed  { background: #00305F; }
QPushButton#PrimaryButton:disabled { background: #B9C4D2; color: #EEEEEE; }

/* ---------- 输入控件 ---------- */
QSpinBox {
    background: #FFFFFF;
    border: 1px solid #C9D2E0;
    border-radius: 6px;
    min-height: 30px;
    padding: 2px 6px;
    selection-background-color: #0A5AA8;
}
QSpinBox:hover { border-color: #7EA6D8; }
QSpinBox:focus { border: 1px solid #2E75B6; }
QSpinBox::up-button, QSpinBox::down-button {
    width: 18px; border: none; background: transparent;
}

QLineEdit {
    background: #FFFFFF;
    border: 1px solid #DDE3EC;
    border-radius: 6px;
    min-height: 28px;
    padding: 2px 8px;
    color: #44546A;
}
QLineEdit:read-only { background: #FAFBFD; }

/* ---------- 表格 ---------- */
QTableView {
    background: #FFFFFF;
    border: 1px solid #DDE3EC;
    border-radius: 6px;
    gridline-color: #EDF1F6;
    selection-background-color: #D6E4F7;
    selection-color: #1F2937;
}
QTableView::item { padding: 4px 8px; }
QTableView::item:selected { background: #D6E4F7; color: #1F2937; }
QHeaderView::section {
    background: #004080;
    color: #FFFFFF;
    font-weight: bold;
    padding: 7px 8px;
    border: none;
    border-right: 1px solid #1A5AA0;
}
QHeaderView::section:last { border-right: none; }
QTableCornerButton::section { background: #004080; border: none; }

/* ---------- 选项卡 ---------- */
QTabWidget::pane {
    border: 1px solid #DDE3EC;
    border-radius: 8px;
    background: #FFFFFF;
    top: -1px;
}
QTabBar::tab {
    background: #E9EEF5;
    color: #44546A;
    padding: 8px 20px;
    margin-right: 3px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
}
QTabBar::tab:selected { background: #004080; color: #FFFFFF; font-weight: bold; }
QTabBar::tab:hover:!selected { background: #DCE6F2; }

/* ---------- 进度条 ---------- */
QProgressBar {
    border: 1px solid #DDE3EC;
    border-radius: 6px;
    background: #FFFFFF;
    text-align: center;
    color: #1F2937;
}
QProgressBar::chunk { background: #2E75B6; border-radius: 5px; }

/* ---------- 滚动条 ---------- */
QScrollBar:vertical {
    background: transparent; width: 12px; margin: 2px;
}
QScrollBar::handle:vertical {
    background: #C3CCDA; border-radius: 5px; min-height: 32px;
}
QScrollBar::handle:vertical:hover { background: #9FB0C6; }
QScrollBar:horizontal {
    background: transparent; height: 12px; margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #C3CCDA; border-radius: 5px; min-width: 32px;
}
QScrollBar::handle:horizontal:hover { background: #9FB0C6; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }

/* ---------- 滚动区（左栏） ---------- */
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }

/* ---------- 状态栏 ---------- */
QStatusBar { background: #E9EEF5; color: #44546A; }
QStatusBar::item { border: none; }

/* ---------- 具名部件 ---------- */
QLabel#Hint        { color: #8A93A2; font-size: 11px; }
QLabel#StageLabel  { color: #44546A; font-size: 12px; }

/* 统计汇总卡片：淡蓝底，让指标从白底里浮出来 */
QLabel#SummaryCard {
    background: #EEF4FB;
    border: 1px solid #D5E2F2;
    border-radius: 8px;
    padding: 11px 14px;
    color: #1F2937;
    font-size: 11.5px;
}
QTextEdit#LogBox {
    background: #1E1E1E;
    color: #D4D4D4;
    border: 1px solid #DDE3EC;
    border-radius: 6px;
}
"""


def apply_theme(app: QApplication) -> None:
    """给整个应用注入统一主题（主窗口与离屏自检共用同一份 QSS）。"""
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei", 10))
    app.setStyleSheet(APP_QSS)

# Qt 枚举在两代之间的差异
_ALIGN_CENTER = Qt.AlignmentFlag.AlignCenter if QT6 else Qt.AlignCenter
_ALIGN_RIGHT = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter if QT6 \
    else Qt.AlignRight | Qt.AlignVCenter
_ALIGN_LEFT = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter if QT6 \
    else Qt.AlignLeft | Qt.AlignVCenter
_TABLE_STRETCH = QHeaderView.ResizeMode.Stretch if QT6 else QHeaderView.Stretch
_TABLE_INTERACTIVE = (QHeaderView.ResizeMode.Interactive if QT6
                      else QHeaderView.Interactive)
_TABLE_FIXED = QHeaderView.ResizeMode.Fixed if QT6 else QHeaderView.Fixed
_SCROLL_PER_PIXEL = (QTableWidget.ScrollMode.ScrollPerPixel if QT6
                     else QTableWidget.ScrollPerPixel)
_ELIDE_RIGHT = Qt.TextElideMode.ElideRight if QT6 else Qt.ElideRight

TABLE_HEADERS = ("A系统-客户名称", "B系统-对方户名", "相似度(%)", "匹配状态", "备注")
TABLE_STATUS_COL = 3
TABLE_REMARK_COL = 4
# 列宽：前两列给名称留足可读宽度（可拖拽），后两列定长居中，备注占满剩余
TABLE_COLUMN_WIDTHS: dict[int, int] = {0: 220, 1: 195, 2: 90, 3: 100}
TABLE_INTERACTIVE_COLS: tuple[int, ...] = (0, 1)      # 名称列，允许拖拽
TABLE_FIXED_COLS_WITH_WIDTH: tuple[int, ...] = (2, 3) # 相似度 / 状态，宽度锁死
TABLE_CENTER_COLS: frozenset[int] = frozenset({2, 3})


# =========================================================================== #
#  后台工作线程
# =========================================================================== #
class MatchWorker(QThread):
    """把整条流水线放到子线程执行，通过信号回主线程。

    信号
    ----
    progress(int, str)  整体进度 0–100 与当前阶段文字
    log(str)            一行运行日志
    succeeded(object)   成功，负载是 ``pipeline.PipelineResult``
    failed(str)         失败，负载是格式化后的异常信息
    """

    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(
        self,
        a_file: str,
        b_file: str,
        thresholds: matcher.Thresholds,
        team: str,
        out_dir: str | None = None,
        algo: str = matcher.DEFAULT_ALGO,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._a_file = a_file
        self._b_file = b_file
        self._thresholds = thresholds
        self._team = team
        self._out_dir = out_dir
        self._algo = algo

    # 子线程体：绝不触碰任何界面元件
    def run(self) -> None:                              # noqa: D102
        try:
            result = pipeline.run_pipeline(
                a_file=self._a_file,
                b_file=self._b_file,
                thresholds=self._thresholds,
                team=self._team,
                out_dir=self._out_dir,
                progress=lambda pct, msg: self.progress.emit(int(pct), msg),
                log=lambda text: self.log.emit(text),
                algo=self._algo,
            )
        except Exception as exc:                        # noqa: BLE001
            detail = "".join(
                traceback.format_exception_only(type(exc), exc)
            ).strip()
            self.failed.emit(detail)
            return
        self.succeeded.emit(result)


# =========================================================================== #
#  内嵌图表
# =========================================================================== #
class ChartCanvas(FigureCanvasQTAgg):
    """匹配结果可视化：左侧对账全局占比饼图 + 右侧细分档位条形图。

    统计口径是**对账全局视角**，而不是只看 A 系统：

    * 饼图 3 片 —— 成功匹配 / A系统独有 / B系统独有
      （只画 A 侧会让「B系统独有」整类数据在图表里凭空消失）
    * 条形图 5+1 根 —— 完全 / 高度 / 中低 / 低置信度 / A系统独有 / B系统独有
      （A 的四档 + A 独有 + B 独有，与 Excel 汇总表的数字一一对应）

    布局用 matplotlib 的 **constrained layout**：窗口缩放时自动重排，
    标题、图例、坐标轴标签不会互相压叠，不需要手工调 ``subplots_adjust``。
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        self._fig = Figure(figsize=(8.6, 4.6), dpi=100, facecolor="white",
                           layout="constrained")
        super().__init__(self._fig)
        self.setParent(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumHeight(320)
        self.show_empty("运行匹配后在此展示对账全局分布")

    def show_empty(self, message: str) -> None:
        self._fig.clear()
        ax = self._fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, message, ha="center", va="center",
                fontsize=13, color="#8A8A8A")
        self.draw_idle()

    # ------------------------------------------------------------------ #
    def plot(
        self,
        counts: dict[str, int],
        thresholds: matcher.Thresholds,
        b_only: int = 0,
    ) -> None:
        """画图。

        参数
        ----
        counts      : A 侧五档计数（来自 ``matcher.summarize``）
        thresholds  : 当前阈值，写进脚注
        b_only      : B 系统独有记录数（来自 Sheet3 行数）
        """
        self._fig.clear()

        matched = sum(v for k, v in counts.items() if k != "A系统独有")
        a_only = counts.get("A系统独有", 0)
        # 饼图基数 = 两个系统去重后的全部参与对账记录
        grand = matched + a_only + b_only
        if grand <= 0:
            self.show_empty("无数据")
            return

        # ---------------- 左：对账全局环形图 ----------------
        pie_items = [("成功匹配", matched), ("A系统独有", a_only),
                     ("B系统独有", b_only)]
        pie_items = [(k, v) for k, v in pie_items if v > 0]
        pie_labels = [k for k, _ in pie_items]
        pie_values = [v for _, v in pie_items]
        pie_colors = [GLOBAL_COLORS[k] for k in pie_labels]
        hit_rate = matched / grand

        ax1 = self._fig.add_subplot(121)
        _, _, autotexts = ax1.pie(
            pie_values,
            labels=None,                     # 名称走图例，避免与百分比标签打架
            colors=pie_colors,
            startangle=90,
            counterclock=False,              # 顺时针，成功匹配从 12 点起
            autopct="%1.1f%%",
            # 环带占 r=0.6~1.0，标签放在 0.80 正好落在环带中线，
            # 用默认的 0.6 会贴到内圈边缘、和色块撞在一起
            pctdistance=0.80,
            wedgeprops={"width": 0.40, "edgecolor": "w", "linewidth": 1.6},
            textprops={"fontsize": 9.5, "color": "white", "fontweight": "bold"},
        )
        for t in autotexts:
            t.set_color("white")

        # 环心文字：大字号成功率 + 小字说明
        ax1.text(0, 0.10, f"{hit_rate:.1%}", ha="center", va="center",
                 fontsize=23, fontweight="bold", color="#1F3864")
        ax1.text(0, -0.16, "全局对账成功率", ha="center", va="center",
                 fontsize=9.5, color="#7A8798")
        ax1.set_title(
            f"对账全局占比\n（A {grand - b_only} 条 + B {grand - a_only} 条 "
            f"→ 去重后 {grand} 条）",
            fontsize=11, fontweight="bold", pad=10,
        )
        ax1.legend(
            [f"{k} {v} 条" for k, v in pie_items],
            loc="upper center", bbox_to_anchor=(0.5, -0.02),
            fontsize=9, frameon=False, ncol=len(pie_items),
            handlelength=1.2, columnspacing=1.3,
        )
        ax1.axis("equal")

        # ---------------- 右：细分档位条数（含 B 系统独有） ----------------
        # 固定展示 A 侧全部分档 + B系统独有，0 值档位也保留 —— 图表结构稳定，
        # 不会因为某档恰好为 0 就从图上"消失"，与汇总表能逐项对上。
        bar_items = [(k, counts.get(k, 0)) for k in BUCKET_ORDER]
        bar_items.append(("B系统独有", b_only))
        bar_labels = [k for k, _ in bar_items]
        bar_values = [v for _, v in bar_items]
        bar_colors = [BUCKET_COLORS.get(k, "#8FAADC") for k in bar_labels]

        ax2 = self._fig.add_subplot(122)
        bars = ax2.barh(bar_labels[::-1], bar_values[::-1],
                        color=bar_colors[::-1], height=0.62)
        ax2.set_title("细分档位条数", fontsize=11, fontweight="bold", pad=10)
        ax2.set_xlabel("记录数", fontsize=9.5)
        ax2.tick_params(labelsize=9.5)
        ax2.grid(axis="x", linestyle=":", alpha=0.45)
        ax2.set_axisbelow(True)
        span = max(bar_values) if any(bar_values) else 1
        for bar, v in zip(bars, bar_values[::-1]):
            ax2.text(bar.get_width() + span * 0.02,
                     bar.get_y() + bar.get_height() / 2,
                     str(v), va="center", fontsize=9.5,
                     color="#404040" if v else "#B0B0B0")
        ax2.set_xlim(0, span * 1.18)
        for side in ("top", "right"):
            ax2.spines[side].set_visible(False)

        # 两行总标题：单行在窄窗口下会被裁掉左半边
        self._fig.suptitle(
            f"成功匹配 {matched} 条  ·  A系统独有 {a_only} 条  ·  "
            f"B系统独有 {b_only} 条\n"
            f"阈值：完全 100% · 高度 ≥{thresholds.high:g}% · "
            f"中低 ≥{thresholds.low:g}% · 识别下限 {thresholds.floor:g}%",
            fontsize=9, color="#595959", linespacing=1.6,
        )
        self.draw_idle()


# =========================================================================== #
#  主窗口
# =========================================================================== #
class MainWindow(QMainWindow):
    """参数配置 → 后台执行 → 结果可视化 → 导出 Excel。"""

    PREVIEW_ROWS = 100          # 结果预览上限（考题数据 85 条可一次看全）

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("多源运营数据相似度模糊匹配工具  ·  Topic03")
        self.setMinimumSize(*WINDOW_MIN_SIZE)
        self.resize(*WINDOW_SIZE)

        self._worker: MatchWorker | None = None       # 必须持引用，否则线程会被 GC
        self._result: pipeline.PipelineResult | None = None
        self._default_team = C.DEFAULT_TEAM

        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        splitter = QSplitter(Qt.Orientation.Horizontal if QT6 else Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([440, 1000])
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter)

        self._load_default_paths()          # 两个面板都建好后，再填默认考题文件

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("就绪 —— 请选择 A / B 两个 Excel 文件")

    # ----------------------------------------------------------------- 左侧
    @staticmethod
    def _group(title: str) -> tuple[QGroupBox, QGridLayout]:
        """建一个带统一内边距 / 描边的分组框，返回 (GroupBox, 栅格布局)。

        内边距靠两条一起给：样式表负责边框与标题缩进，布局的 contentsMargins
        负责控件与边框之间的呼吸感。只给其中一条都会显得挤。
        """
        box = QGroupBox(title)
        grid = QGridLayout(box)
        grid.setContentsMargins(13, 18, 13, 11)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(7)
        return box, grid

    def _build_left_panel(self) -> QScrollArea:
        panel = QWidget()
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        # --- 数据源：A / B 各一个文件，两个按钮分别选择 ---
        src_box, grid = self._group("① 数据源")
        grid.setColumnStretch(1, 1)

        self.edit_a = QLineEdit(str(C.A_FILE))
        self.edit_b = QLineEdit(str(C.B_FILE))
        for row, (label, edit, tip) in enumerate((
            ("A 系统（ERP 客户明细）", self.edit_a, "核心对比列：客户名称"),
            ("B 系统（银行流水）", self.edit_b, "核心对比列：对方户名"),
        )):
            edit.setReadOnly(True)
            edit.setToolTip(f"{tip}\n{edit.text()}")
            btn = QPushButton("选择…")
            btn.setFixedWidth(76)
            btn.clicked.connect(lambda _, e=edit: self._pick_file(e))
            grid.addWidget(QLabel(label), row, 0)
            grid.addWidget(edit, row, 1)
            grid.addWidget(btn, row, 2)
        lay.addWidget(src_box)

        # 注意：初始路径在 __init__ 里等右侧面板建好后再填（日志框要用）。

        # --- 阈值 ---
        thr_box, tg = self._group("② 匹配阈值")
        self.spin_high = QSpinBox()
        self.spin_high.setRange(80, 100)
        self.spin_high.setValue(int(matcher.DEFAULT_THRESHOLDS.high))
        self.spin_high.setSuffix("  %")
        self.spin_low = QSpinBox()
        self.spin_low.setRange(50, 80)
        self.spin_low.setValue(int(matcher.DEFAULT_THRESHOLDS.low))
        self.spin_low.setSuffix("  %")
        self.spin_floor = QSpinBox()
        self.spin_floor.setRange(0, 70)
        self.spin_floor.setValue(int(matcher.DEFAULT_THRESHOLDS.floor))
        self.spin_floor.setSuffix("  %")
        for sp in (self.spin_high, self.spin_low, self.spin_floor):
            sp.setFixedWidth(100)
        tg.addWidget(QLabel("高度匹配阈值（≥ 直接确认）"), 0, 0)
        tg.addWidget(self.spin_high, 0, 1)
        tg.addWidget(QLabel("中低匹配分界线（≥ 建议复核）"), 1, 0)
        tg.addWidget(self.spin_low, 1, 1)
        tg.addWidget(QLabel("识别下限（< 判为独有）"), 2, 0)
        tg.addWidget(self.spin_floor, 2, 1)

        self.btn_reset_thresholds = QPushButton("↺  恢复默认阈值（90 / 70 / 60）")
        self.btn_reset_thresholds.setMinimumHeight(28)
        self.btn_reset_thresholds.setToolTip(
            "一键重置为赛题推荐口径：高度 90%、中低 70%、识别下限 60%"
        )
        self.btn_reset_thresholds.clicked.connect(self._reset_thresholds)
        tg.addWidget(self.btn_reset_thresholds, 3, 0, 1, 2)

        hint = QLabel("默认 90 / 70 / 60，即赛题推荐口径。")
        hint.setObjectName("Hint")
        tg.addWidget(hint, 4, 0, 1, 2)
        lay.addWidget(thr_box)

        # --- 算法（赛题 §3.2 加分项：多种相似度算法切换） ---
        algo_box, ag = self._group("③ 相似度算法")
        self.combo_algo = QComboBox()
        for key in matcher.ALGO_CHOICES:
            self.combo_algo.addItem(matcher.ALGO_LABELS[key], key)
        self.combo_algo.setCurrentIndex(
            list(matcher.ALGO_CHOICES).index(matcher.DEFAULT_ALGO)
        )
        self.combo_algo.setMinimumHeight(28)
        ag.addWidget(self.combo_algo, 0, 0, 1, 2)

        self.lbl_algo_desc = QLabel()
        self.lbl_algo_desc.setObjectName("Hint")
        self.lbl_algo_desc.setWordWrap(True)
        ag.addWidget(self.lbl_algo_desc, 1, 0, 1, 2)

        algo_hint = QLabel(
            "「加权组合」为本工具主算法。切换算法只更换最底层的相似度计算；"
            "字号差异惩罚、通用词折叠、关键字号保护等业务规则在任何模式下都生效。"
        )
        algo_hint.setObjectName("Hint")
        algo_hint.setWordWrap(True)
        ag.addWidget(algo_hint, 2, 0, 1, 2)

        self.combo_algo.currentIndexChanged.connect(self._on_algo_changed)
        self._on_algo_changed()
        lay.addWidget(algo_box)

        # --- 执行 ---
        run_box, rl = self._group("④ 执行")
        rl.setColumnStretch(0, 1)
        self.btn_run = QPushButton("▶  开始智能匹配")
        self.btn_run.setMinimumHeight(44)
        self.btn_run.setObjectName("PrimaryButton")
        self.btn_run.clicked.connect(self._start)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("已完成 %p%")
        self.progress.setMinimumHeight(22)

        self.lbl_stage = QLabel("等待开始")
        self.lbl_stage.setObjectName("StageLabel")

        self.btn_export = QPushButton("💾  导出完整报表…")
        self.btn_export.setMinimumHeight(36)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_as)

        rl.addWidget(self.btn_run, 0, 0)
        rl.addWidget(self.progress, 1, 0)
        rl.addWidget(self.lbl_stage, 2, 0)
        rl.addWidget(self.btn_export, 3, 0)
        lay.addWidget(run_box)

        # --- 摘要 ---
        self.lbl_summary = QLabel("尚未运行")
        self.lbl_summary.setWordWrap(True)
        self.lbl_summary.setObjectName("SummaryCard")
        self.lbl_summary.setAlignment(_ALIGN_LEFT)
        lay.addWidget(self.lbl_summary)

        lay.addStretch(1)

        # 窗口缩到最小尺寸时左栏内容会比视口高，套一层滚动区保证不丢控件
        scroll = QScrollArea()
        scroll.setWidget(panel)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame if QT6 else QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff if QT6 else Qt.ScrollBarAlwaysOff
        )
        scroll.setMinimumWidth(400)
        scroll.setMaximumWidth(470)
        return scroll

    # ----------------------------------------------------------------- 右侧
    def _build_right_panel(self) -> QWidget:
        tabs = QTabWidget()

        # 图表
        self.chart = ChartCanvas()
        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(14, 12, 14, 12)
        wl.addWidget(self.chart)
        tabs.addTab(wrap, "📊 匹配分布图表")

        # 预览
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(TABLE_HEADERS)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers if QT6
                                   else QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows if QT6
            else QTableWidget.SelectRows
        )
        # 大表滚动：按像素滚动更顺滑；关掉自动换行，避免行高忽高忽低
        self.table.setVerticalScrollMode(_SCROLL_PER_PIXEL)
        self.table.setHorizontalScrollMode(_SCROLL_PER_PIXEL)
        self.table.setWordWrap(False)
        self.table.setTextElideMode(_ELIDE_RIGHT)

        hh = self.table.horizontalHeader()
        # 显式列宽：名称两列给固定初始宽度（可拖拽微调），
        # 相似度 / 匹配状态是定长内容给 Fixed，备注吃掉全部剩余空间。
        for col in TABLE_INTERACTIVE_COLS:          # 名称列：给足宽度，可拖拽微调
            hh.setSectionResizeMode(col, _TABLE_INTERACTIVE)
        for col in TABLE_FIXED_COLS_WITH_WIDTH:     # 相似度 / 匹配状态：定长内容，锁死
            hh.setSectionResizeMode(col, _TABLE_FIXED)
        for col, width in TABLE_COLUMN_WIDTHS.items():
            self.table.setColumnWidth(col, width)
        hh.setSectionResizeMode(TABLE_REMARK_COL, _TABLE_STRETCH)
        hh.setMinimumSectionSize(70)
        hh.setStretchLastSection(True)
        self.table.setMinimumHeight(280)
        tabs.addTab(self.table, f"📋 结果预览（前 {self.PREVIEW_ROWS} 条）")

        # 日志
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 9))
        self.log_box.setObjectName("LogBox")
        tabs.addTab(self.log_box, "📝 运行日志")

        self.tabs = tabs
        return tabs

    # ================================================================= 交互
    def _set_path(self, target: QLineEdit, path: str) -> None:
        """写入路径并把光标归零 —— 否则长路径会从尾部显示，开头被截掉。"""
        target.setText(path)
        target.setCursorPosition(0)
        target.setToolTip(path)

    def _pick_file(self, target: QLineEdit) -> None:
        """为 A 或 B 选择**单个** Excel 文件。"""
        start = Path(target.text()).parent if target.text() else C.DATA_DIR
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 数据文件", str(start),
            "Excel 工作簿 (*.xlsx *.xlsm *.xls);;所有文件 (*)",
        )
        if not path:
            return
        self._set_path(target, path)
        self.log_box.append(f"[选择] {path}")

    def _reset_thresholds(self) -> None:
        """把三个阈值一键重置为赛题推荐的 90 / 70 / 60。"""
        t = matcher.DEFAULT_THRESHOLDS
        self.spin_high.setValue(int(t.high))
        self.spin_low.setValue(int(t.low))
        self.spin_floor.setValue(int(t.floor))
        self.log_box.append(f"[阈值] 已恢复默认：{t.label}")
        self.statusBar().showMessage(f"阈值已重置为 {t.high:g} / {t.low:g} / {t.floor:g}")

    def _load_default_paths(self) -> None:
        """启动时把两个路径框填成考题默认文件（一次性，不提供按钮）。"""
        self._set_path(self.edit_a, str(C.A_FILE))
        self._set_path(self.edit_b, str(C.B_FILE))
        self.edit_a.setToolTip(f"核心对比列：客户名称\n{C.A_FILE}")
        self.edit_b.setToolTip(f"核心对比列：对方户名\n{C.B_FILE}")

    def inputs(self) -> list[str]:
        """当前选定的两个文件路径 ``[A, B]``。"""
        return [self.edit_a.text().strip(), self.edit_b.text().strip()]

    def _thresholds(self) -> matcher.Thresholds:
        """从界面读取阈值；``Thresholds`` 自身会校验顺序合法性。"""
        return matcher.Thresholds(
            high=float(self.spin_high.value()),
            low=float(self.spin_low.value()),
            floor=float(self.spin_floor.value()),
        )

    def _algo(self) -> str:
        """当前选定的基础相似度算法（§3.2 加分项）。"""
        return self.combo_algo.currentData() or matcher.DEFAULT_ALGO

    def _on_algo_changed(self) -> None:
        """切换算法时刷新说明文字，并把选择写进日志。"""
        algo = self._algo()
        self.lbl_algo_desc.setText(matcher.describe_algo(algo))
        self.combo_algo.setToolTip(
            f"{matcher.ALGO_LABELS[algo]}\n{matcher.describe_algo(algo)}"
        )
        if hasattr(self, "log_box"):
            self.log_box.append(f"[算法] 已选择：{matcher.describe_algo(algo)}")
        self.statusBar().showMessage(f"相似度算法：{matcher.ALGO_LABELS[algo]}")

    # ================================================================= 执行
    def _start(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return

        a_file, b_file = self.inputs()
        for label, path in (("A 系统", a_file), ("B 系统", b_file)):
            if not path:
                self._warn(f"请先选择{label}的数据文件。")
                return
            if not Path(path).exists():
                self._warn(f"{label}文件不存在：\n{path}")
                return
        if a_file == b_file:
            self._warn("A 系统与 B 系统选了同一个文件，请分别选择。")
            return

        try:
            thresholds = self._thresholds()
        except ValueError as exc:
            self._warn(f"阈值组合非法：\n{exc}")
            return

        self.log_box.clear()
        self.table.setRowCount(0)
        self.chart.show_empty("正在匹配…")
        self.progress.setValue(0)
        self.lbl_stage.setText("正在启动…")
        self.lbl_summary.setText("运行中…")
        self.btn_run.setEnabled(False)
        self.btn_export.setEnabled(False)

        self._worker = MatchWorker(a_file, b_file, thresholds,
                                   self._default_team, None, self._algo(), self)
        self._worker.progress.connect(self._on_progress)
        self._worker.log.connect(self._on_log)
        self._worker.succeeded.connect(self._on_success)
        self._worker.failed.connect(self._on_failure)
        self._worker.finished.connect(lambda: self.btn_run.setEnabled(True))
        self._worker.start()

    # ------------------------------------------------------------ 槽函数
    def _on_progress(self, pct: int, message: str) -> None:
        self.progress.setValue(pct)
        self.lbl_stage.setText(message)
        self.statusBar().showMessage(f"{message}（{pct}%）")

    def _on_log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum()
        )

    def _on_success(self, result: pipeline.PipelineResult) -> None:
        """槽：主线程收到匹配结果 → 刷新界面 → 弹提示框。"""
        self._apply_result(result)

        counts = result.counts
        body = "\n".join(f"    {k:<14} {v:>4} 行" for k, v in result.sheets.items())
        QMessageBox.information(
            self, "匹配完成",
            f"匹配已完成，用时 {result.elapsed:.2f} 秒。\n\n"
            f"阈值：{result.thresholds.label}\n"
            f"（完全 / 高度 / 中低 / 识别下限）\n\n"
            "分档结果：\n"
            + "\n".join(f"    {k:<10} {v:>4} 条" for k, v in counts.items())
            + f"\n\n成果报表各 Sheet 行数：\n{body}\n\n"
            f"已自动保存至：\n{result.path}\n\n"
            "可点击「导出完整报表…」另存到指定位置。",
        )

    def _apply_result(self, result: pipeline.PipelineResult) -> None:
        """把结果填进图表 / 表格 / 摘要（无模态弹窗，便于离屏自检复用）。"""
        self._result = result
        self.progress.setValue(100)
        self.lbl_stage.setText("✅ 全部完成")
        self.statusBar().showMessage(f"完成，用时 {result.elapsed:.2f} 秒")

        self.chart.plot(result.counts, result.thresholds, self._b_only(result))
        self._fill_table(result)
        self._fill_summary(result)
        self.btn_export.setEnabled(True)
        self.tabs.setCurrentIndex(0)

    @staticmethod
    def _b_only(result: pipeline.PipelineResult) -> int:
        """B 系统独有记录数（取 Sheet3 行数）。"""
        if not result.report:
            return 0
        return len(result.report["frames"].get("sheet3", []))

    def _on_failure(self, detail: str) -> None:
        self.lbl_stage.setText("❌ 运行失败")
        self.progress.setValue(0)
        self.statusBar().showMessage("运行失败")
        self.log_box.append(f"[错误] {detail}")
        QMessageBox.critical(self, "运行失败", detail)

    # ------------------------------------------------------------ 结果填充
    def _fill_table(self, result: pipeline.PipelineResult) -> None:
        frame = result.frame.head(self.PREVIEW_ROWS)
        self.table.setColumnCount(len(TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(list(TABLE_HEADERS))
        self.table.setRowCount(len(frame))
        for r, (_, row) in enumerate(frame.iterrows()):
            status = str(row["匹配状态"])
            row_fill, badge_bg, badge_fg = STATUS_STYLES.get(
                status, (None, None, None)
            )
            remark = str(row["备注"])
            cells = (
                (str(row["A系统-客户名称"]), _ALIGN_LEFT, None),
                (str(row["B系统-对方户名"]), _ALIGN_LEFT, None),
                (f"{float(row['相似度(%)']):.1f}", _ALIGN_CENTER, None),
                (status, _ALIGN_CENTER, None),
                (remark, _ALIGN_LEFT,
                 REMARK_FG.get(status, REMARK_FG_DEFAULT)),
            )
            for col, (text, align, fg) in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(align)
                item.setToolTip(text)
                if col == TABLE_STATUS_COL and badge_bg:
                    # 状态列：实色徽标 + 加粗白字，四档一眼可辨
                    item.setBackground(QColor(badge_bg))
                    item.setForeground(QColor(badge_fg))
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                elif row_fill:
                    item.setBackground(QColor(row_fill))
                    if fg:
                        item.setForeground(QColor(fg))
                elif fg:
                    item.setForeground(QColor(fg))
                self.table.setItem(r, col, item)

        # 行高固定，避免上百行时因字号差异导致行高参差
        self.table.verticalHeader().setDefaultSectionSize(26)

    def _fill_summary(self, result: pipeline.PipelineResult) -> None:
        t = result.thresholds
        c = result.counts
        b_only = self._b_only(result)
        matched = len(result.frame)
        grand = matched + c["A系统独有"] + b_only
        self.lbl_summary.setText(
            f"<b>对账全局</b>（A {len(result.df_a)} 条 · B {len(result.df_b)} 条 "
            f"→ 参与对账 {grand} 条）<br>"
            f"成功匹配：<b>{matched}</b> 条"
            f"（{matched / grand:.1%}）<br>"
            f"A系统独有：<b>{c['A系统独有']}</b> 条"
            f"（{c['A系统独有'] / grand:.1%}）<br>"
            f"B系统独有：<b>{b_only}</b> 条（{b_only / grand:.1%}）<br>"
            "<hr style='border:none;border-top:1px solid #E3E8EF;margin:6px 0;'>"
            f"<b>成功匹配细分</b><br>"
            f"完全匹配（100%）：<b>{c['完全匹配']}</b> 条<br>"
            f"高度匹配（≥{t.high:g}%）：<b>{c['高度匹配']}</b> 条<br>"
            f"中低匹配（{t.low:g}–{t.high:g}%）：<b>{c['中低匹配']}</b> 条<br>"
            f"低置信度（{t.floor:g}–{t.low:g}%）：<b>{c['低置信度匹配']}</b> 条<br>"
            f"<span style='color:#7A7A7A'>耗时 {result.elapsed:.2f}s ｜ "
            f"清洗 {result.clean_seconds:.2f}s ｜ 匹配 {result.match_seconds:.2f}s</span>"
        )

    # ------------------------------------------------------------ 导出
    def _export_as(self) -> None:
        if self._result is None or self._result.report is None:
            return
        from src import exporter

        suggestion = str(self._result.path
                         or exporter.default_filename(self._default_team))
        path, _ = QFileDialog.getSaveFileName(
            self, "导出完整报表", suggestion, "Excel 工作簿 (*.xlsx)"
        )
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".xlsx":
            target = target.with_suffix(".xlsx")
        try:
            report = exporter.export_report(
                self._result.df_a, self._result.df_b, self._result.outcome,
                out_dir=target.parent, filename=target.name,
                thresholds=self._result.thresholds,
            )
        except PermissionError:
            self._warn(f"无法写入文件（可能正在 Excel 中打开）：\n{target}\n\n"
                       "请关闭该文件后重试。")
            return
        except Exception as exc:                        # noqa: BLE001
            self._warn(f"导出失败：{type(exc).__name__}: {exc}")
            return

        self.log_box.append(f"[导出] 已另存为 {report['path']}")
        QMessageBox.information(self, "导出成功", f"报表已保存至：\n{report['path']}")

    def _warn(self, message: str) -> None:
        QMessageBox.warning(self, "提示", message)

    # ================================================================= 收尾
    def closeEvent(self, event) -> None:                # noqa: N802
        if self._worker is not None and self._worker.isRunning():
            self._worker.requestInterruption()
            self._worker.wait(3000)
        event.accept()


# =========================================================================== #
#  入口
# =========================================================================== #
def selftest() -> int:
    """离屏自检：装配窗口 → 同步跑一次完整流水线 → 填充界面，验证无异常。

    产物写到 ``output/_selftest/``，避免覆盖正式成果文件
    （同时也不会被"文件正开在 Excel 里"的文件锁影响）。
    """
    import os

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    app = QApplication(sys.argv)                        # noqa: F841
    apply_theme(app)

    win = MainWindow()
    win.show()
    print("[selftest] 窗口装配成功")
    a_file, b_file = win.inputs()
    print(f"[selftest] A 系统：{Path(a_file).name}")
    print(f"[selftest] B 系统：{Path(b_file).name}")
    print(f"[selftest] 阈值控件：high={win.spin_high.value()} "
          f"low={win.spin_low.value()} floor={win.spin_floor.value()} "
          f"→ {win._thresholds().label}")
    print(f"[selftest] 算法控件：{matcher.ALGO_LABELS[win._algo()]} "
          f"（{win._algo()}）｜可选 {len(matcher.ALGO_CHOICES)} 种")

    # ---- 交互回归：把阈值改乱，再点「恢复默认阈值」应回到 90/70/60 ----
    win.spin_high.setValue(85)
    win.spin_low.setValue(55)
    win.spin_floor.setValue(30)
    win.btn_reset_thresholds.click()
    restored = (win.spin_high.value(), win.spin_low.value(), win.spin_floor.value())
    print(f"[selftest] 恢复默认阈值按钮 → {restored}"
          f" {'✓' if restored == (90, 70, 60) else '✗ 期望 (90, 70, 60)'}")
    assert restored == (90, 70, 60), "恢复默认阈值未生效"

    # ---- 交互回归：算法下拉框每一项都能选中，且默认项就是主算法 ----
    assert win._algo() == matcher.DEFAULT_ALGO, "算法下拉框默认值不是主算法"
    picked = []
    for idx, key in enumerate(matcher.ALGO_CHOICES):
        win.combo_algo.setCurrentIndex(idx)
        assert win._algo() == key, f"选中 {key} 后 _algo() 返回 {win._algo()}"
        picked.append(win._algo())
    win.combo_algo.setCurrentIndex(
        list(matcher.ALGO_CHOICES).index(matcher.DEFAULT_ALGO)
    )
    print(f"[selftest] 算法下拉切换 → {picked} ✓")

    out_dir = C.OUTPUT_DIR / "_selftest"
    result = pipeline.run_pipeline(
        a_file=a_file, b_file=b_file, thresholds=win._thresholds(),
        out_dir=out_dir, log=lambda t: print(f"[selftest] {t}"),
        algo=win._algo(),
    )
    win._apply_result(result)                          # 不弹模态框，便于离屏运行
    print(f"[selftest] 图表子图数：{len(win.chart.figure.axes)}")
    print(f"[selftest] 表格 {win.table.rowCount()} 行 × "
          f"{win.table.columnCount()} 列")
    print(f"[selftest] 导出按钮可用：{win.btn_export.isEnabled()}")
    print(f"[selftest] 成果文件：{result.path}")

    # ---- 真实 QThread 回归：直接驱动 MatchWorker，验证 4 类信号都能发出 ----
    from PyQt6.QtCore import QEventLoop

    got: dict[str, object] = {"progress": [], "logs": 0, "ok": None, "err": None}
    worker = MatchWorker(a_file, b_file, win._thresholds(), C.DEFAULT_TEAM,
                         str(out_dir), win._algo())
    worker.progress.connect(lambda p, m: got["progress"].append(p))
    worker.log.connect(lambda t: got.__setitem__("logs", int(got["logs"]) + 1))
    worker.succeeded.connect(lambda r: got.__setitem__("ok", r))
    worker.failed.connect(lambda e: got.__setitem__("err", e))

    loop = QEventLoop()
    worker.finished.connect(loop.quit)
    worker.start()
    loop.exec()

    if got["err"]:
        print(f"[selftest] ✗ Worker 线程报错：{got['err']}")
        return 1
    pcts = got["progress"]
    print(f"[selftest] QThread 信号：progress {len(pcts)} 次"
          f"（{min(pcts)}% → {max(pcts)}%），log {got['logs']} 条，"
          f"成功率 {got['ok'] is not None}")
    assert got["ok"] is not None, "Worker 未发出 succeeded 信号"

    # ---- 截图，便于人工确认界面布局 ----
    win.resize(*WINDOW_SIZE)
    shot = out_dir / "gui_preview.png"
    shot.parent.mkdir(parents=True, exist_ok=True)
    win.grab().save(str(shot))
    print(f"[selftest] 界面截图：{shot}")

    # ---- 单独导出图表（供作品说明文档引用）----
    # 说明：整窗截图走 Qt 离屏渲染，中文字体在无字体环境下会退化成方块；
    # 而 matplotlib 图表用自带的字体配置渲染，中文始终正常。故文档配图
    # 单独从 Figure 落盘，不走 grab()。
    chart_png = out_dir / "gui_chart.png"
    win.chart.figure.set_dpi(144)
    win.chart.figure.savefig(str(chart_png), bbox_inches="tight",
                             facecolor="white")
    print(f"[selftest] 图表导出：{chart_png}")

    print("[selftest] 全部通过 ✅")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Topic03 模糊匹配工具 · PyQt 桌面端")
    parser.add_argument("--selftest", action="store_true",
                        help="离屏装配窗口并跑一遍完整流程后退出（用于验证/CI）")
    parser.add_argument("--selftest-log", type=Path, default=None, metavar="FILE",
                        help="把自检输出写入文件 —— 打包为 --windowed 后 stdout 不可见，"
                             "构建脚本靠这个文件校验 exe 是否真的跑通")
    args = parser.parse_args(argv)

    if args.selftest:
        if args.selftest_log:
            args.selftest_log.parent.mkdir(parents=True, exist_ok=True)
            handle = open(args.selftest_log, "w", encoding="utf-8")  # noqa: SIM115
            sys.stdout = sys.stderr = handle
        try:
            return selftest()
        finally:
            if args.selftest_log:
                sys.stdout.flush()

    app = QApplication(sys.argv)
    apply_theme(app)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
