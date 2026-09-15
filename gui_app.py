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
        QApplication, QFileDialog, QGridLayout, QGroupBox, QHBoxLayout,
        QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar,
        QPushButton, QSizePolicy, QSpinBox, QSplitter, QStatusBar, QTabWidget,
        QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget,
    )

    QT6 = True
except ImportError:                                     # pragma: no cover
    from PyQt5.QtCore import Qt, QThread, pyqtSignal  # type: ignore
    from PyQt5.QtGui import QColor, QFont             # type: ignore
    from PyQt5.QtWidgets import (                     # type: ignore
        QApplication, QFileDialog, QGridLayout, QGroupBox, QHBoxLayout,
        QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar,
        QPushButton, QSizePolicy, QSpinBox, QSplitter, QStatusBar, QTabWidget,
        QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget,
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
#  配色（与 Excel 报表保持一致）
# --------------------------------------------------------------------------- #
C_GREEN = "#C6EFCE"
C_YELLOW = "#FFEB9C"
C_RED = "#C00000"
C_BLUE = "#004080"
C_BG = "#F4F6F9"

BUCKET_COLORS = ("#2E75B6", "#70AD47", "#FFC000", "#ED7D31", "#C00000")

# Qt 枚举在两代之间的差异
_ALIGN_CENTER = Qt.AlignmentFlag.AlignCenter if QT6 else Qt.AlignCenter
_ALIGN_RIGHT = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter if QT6 \
    else Qt.AlignRight | Qt.AlignVCenter
_ALIGN_LEFT = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter if QT6 \
    else Qt.AlignLeft | Qt.AlignVCenter
_TABLE_STRETCH = QHeaderView.ResizeMode.Stretch if QT6 else QHeaderView.Stretch
_TABLE_FIXED = QHeaderView.ResizeMode.Interactive if QT6 else QHeaderView.Interactive


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
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._a_file = a_file
        self._b_file = b_file
        self._thresholds = thresholds
        self._team = team
        self._out_dir = out_dir

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
    """匹配分档饼图 + 条形图（共用一张 Figure 的左右两个子图）。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        self._fig = Figure(figsize=(8, 4.2), dpi=100, facecolor="white")
        super().__init__(self._fig)
        self.setParent(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.show_empty("运行匹配后在此展示各档位分布")

    def show_empty(self, message: str) -> None:
        self._fig.clear()
        ax = self._fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, message, ha="center", va="center",
                fontsize=13, color="#8A8A8A")
        self.draw_idle()

    def plot(self, counts: dict[str, int], thresholds: matcher.Thresholds) -> None:
        """画饼图（占比）+ 条形图（绝对条数），忽略 0 值档位。"""
        self._fig.clear()
        items = [(k, v) for k, v in counts.items() if v > 0]
        if not items:
            self.show_empty("无数据")
            return
        labels = [k for k, _ in items]
        values = [v for _, v in items]
        colors = [BUCKET_COLORS[list(counts).index(k) % len(BUCKET_COLORS)]
                  for k, _ in items]
        total = sum(values)

        ax1 = self._fig.add_subplot(121)
        ax1.pie(
            values, labels=None, colors=colors, autopct="%1.1f%%", startangle=90,
            wedgeprops={"edgecolor": "white", "linewidth": 1.2},
            textprops={"fontsize": 9},
        )
        ax1.set_title(f"匹配档位占比（共 {total} 条 A 记录）",
                      fontsize=11, fontweight="bold", pad=12)
        ax1.legend(labels, loc="lower center", bbox_to_anchor=(0.5, -0.18),
                   fontsize=8, frameon=False, ncol=2)
        ax1.axis("equal")

        ax2 = self._fig.add_subplot(122)
        bars = ax2.barh(labels[::-1], values[::-1], color=colors[::-1], height=0.6)
        ax2.set_title("各档位条数", fontsize=11, fontweight="bold", pad=12)
        ax2.set_xlabel("记录数", fontsize=9)
        ax2.tick_params(labelsize=9)
        ax2.grid(axis="x", linestyle=":", alpha=0.45)
        ax2.set_axisbelow(True)
        span = max(values) if values else 1
        for bar, v in zip(bars, values[::-1]):
            ax2.text(bar.get_width() + span * 0.02, bar.get_y() + bar.get_height() / 2,
                     str(v), va="center", fontsize=9)
        ax2.set_xlim(0, span * 1.18)
        for side in ("top", "right"):
            ax2.spines[side].set_visible(False)

        self._fig.suptitle(
            f"阈值：完全 100%  ·  高度 ≥{thresholds.high:g}%  ·  "
            f"中低 ≥{thresholds.low:g}%  ·  识别下限 {thresholds.floor:g}%",
            fontsize=9, color="#595959", y=0.02,
        )
        self._fig.tight_layout(rect=(0, 0.05, 1, 1))
        self.draw_idle()


# =========================================================================== #
#  主窗口
# =========================================================================== #
class MainWindow(QMainWindow):
    """参数配置 → 后台执行 → 结果可视化 → 导出 Excel。"""

    PREVIEW_ROWS = 15

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("多源运营数据相似度模糊匹配工具  ·  Topic03")
        self.resize(1360, 860)

        self._worker: MatchWorker | None = None       # 必须持引用，否则线程会被 GC
        self._result: pipeline.PipelineResult | None = None
        self._default_team = C.DEFAULT_TEAM

        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        splitter = QSplitter(Qt.Orientation.Horizontal if QT6 else Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([420, 940])
        root.addWidget(splitter)

        self._reset_files()                 # 两个面板都建好后，再填默认考题文件

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("就绪 —— 请选择 A / B 两个 Excel 文件")

    # ----------------------------------------------------------------- 左侧
    def _build_left_panel(self) -> QWidget:
        panel = QWidget()
        panel.setMinimumWidth(400)
        panel.setMaximumWidth(520)
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        # --- 数据源：A / B 各一个文件，两个按钮分别选择 ---
        src_box = QGroupBox("① 数据源")
        grid = QGridLayout(src_box)
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

        reset = QPushButton("恢复默认考题文件")
        reset.clicked.connect(self._reset_files)
        grid.addWidget(reset, 2, 1, 1, 2)
        lay.addWidget(src_box)

        # 注意：初始路径在 __init__ 里等右侧面板建好后再填（日志框要用）。

        # --- 阈值 ---
        thr_box = QGroupBox("② 匹配阈值（可调 · 加分项）")
        tg = QGridLayout(thr_box)
        self.spin_high = QSpinBox()
        self.spin_high.setRange(80, 100)
        self.spin_high.setValue(90)
        self.spin_high.setSuffix("  %")
        self.spin_low = QSpinBox()
        self.spin_low.setRange(50, 80)
        self.spin_low.setValue(70)
        self.spin_low.setSuffix("  %")
        self.spin_floor = QSpinBox()
        self.spin_floor.setRange(0, 70)
        self.spin_floor.setValue(60)
        self.spin_floor.setSuffix("  %")
        for sp in (self.spin_high, self.spin_low, self.spin_floor):
            sp.setFixedWidth(96)
        tg.addWidget(QLabel("高度匹配阈值（≥ 直接确认）"), 0, 0)
        tg.addWidget(self.spin_high, 0, 1)
        tg.addWidget(QLabel("中低匹配分界线（≥ 建议复核）"), 1, 0)
        tg.addWidget(self.spin_low, 1, 1)
        tg.addWidget(QLabel("识别下限（< 判为独有）"), 2, 0)
        tg.addWidget(self.spin_floor, 2, 1)
        hint = QLabel("默认 90 / 70 / 60，即赛题推荐口径。")
        hint.setStyleSheet("color:#7A7A7A; font-size:11px;")
        tg.addWidget(hint, 3, 0, 1, 2)
        lay.addWidget(thr_box)

        # --- 执行 ---
        run_box = QGroupBox("③ 执行")
        rl = QVBoxLayout(run_box)
        self.btn_run = QPushButton("▶  开始智能匹配")
        self.btn_run.setMinimumHeight(46)
        self.btn_run.setStyleSheet(
            f"QPushButton{{background:{C_BLUE};color:white;font-size:15px;"
            f"font-weight:bold;border-radius:6px;}}"
            f"QPushButton:hover{{background:#0A5AA8;}}"
            f"QPushButton:disabled{{background:#B9C4D2;color:#EEEEEE;}}"
        )
        self.btn_run.clicked.connect(self._start)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("已完成 %p%")
        self.progress.setMinimumHeight(24)

        self.lbl_stage = QLabel("等待开始")
        self.lbl_stage.setStyleSheet("color:#44546A;font-size:12px;")

        self.btn_export = QPushButton("💾  导出完整报表…")
        self.btn_export.setMinimumHeight(38)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_as)

        rl.addWidget(self.btn_run)
        rl.addWidget(self.progress)
        rl.addWidget(self.lbl_stage)
        rl.addWidget(self.btn_export)
        lay.addWidget(run_box)

        # --- 摘要 ---
        self.lbl_summary = QLabel("尚未运行")
        self.lbl_summary.setWordWrap(True)
        self.lbl_summary.setStyleSheet(
            "background:white;border:1px solid #D6DCE4;border-radius:6px;"
            "padding:10px;color:#333333;font-size:12px;line-height:160%;"
        )
        self.lbl_summary.setAlignment(_ALIGN_LEFT)
        lay.addWidget(self.lbl_summary)

        lay.addStretch(1)
        return panel

    # ----------------------------------------------------------------- 右侧
    def _build_right_panel(self) -> QWidget:
        tabs = QTabWidget()

        # 图表
        self.chart = ChartCanvas()
        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(6, 6, 6, 6)
        wl.addWidget(self.chart)
        tabs.addTab(wrap, "📊 匹配分布图表")

        # 预览
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["A系统-客户名称", "B系统-对方户名", "相似度(%)", "匹配状态", "备注"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers if QT6
                                   else QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows if QT6
            else QTableWidget.SelectRows
        )
        hh = self.table.horizontalHeader()
        for i in range(4):
            hh.setSectionResizeMode(i, _TABLE_FIXED)
        hh.setSectionResizeMode(4, _TABLE_STRETCH)
        self.table.setColumnWidth(0, 250)
        self.table.setColumnWidth(1, 250)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 100)
        tabs.addTab(self.table, f"📋 结果预览（前 {self.PREVIEW_ROWS} 条）")

        # 日志
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 9))
        self.log_box.setStyleSheet("background:#1E1E1E;color:#D4D4D4;border:none;")
        tabs.addTab(self.log_box, "📝 运行日志")

        self.tabs = tabs
        return tabs

    # ================================================================= 交互
    def _pick_file(self, target: QLineEdit) -> None:
        """为 A 或 B 选择**单个** Excel 文件。"""
        start = Path(target.text()).parent if target.text() else C.DATA_DIR
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 数据文件", str(start),
            "Excel 工作簿 (*.xlsx *.xlsm *.xls);;所有文件 (*)",
        )
        if not path:
            return
        target.setText(path)
        target.setToolTip(path)
        self.log_box.append(f"[选择] {path}")

    def _reset_files(self) -> None:
        self.edit_a.setText(str(C.A_FILE))
        self.edit_b.setText(str(C.B_FILE))
        if hasattr(self, "log_box"):
            self.log_box.append(f"[选择] 已恢复默认考题文件")
        self.statusBar().showMessage("已恢复为考题默认文件")

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
                                   self._default_team, None, self)
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

        self.chart.plot(result.counts, result.thresholds)
        self._fill_table(result)
        self._fill_summary(result)
        self.btn_export.setEnabled(True)
        self.tabs.setCurrentIndex(0)

    def _on_failure(self, detail: str) -> None:
        self.lbl_stage.setText("❌ 运行失败")
        self.progress.setValue(0)
        self.statusBar().showMessage("运行失败")
        self.log_box.append(f"[错误] {detail}")
        QMessageBox.critical(self, "运行失败", detail)

    # ------------------------------------------------------------ 结果填充
    def _fill_table(self, result: pipeline.PipelineResult) -> None:
        frame = result.frame.head(self.PREVIEW_ROWS)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["A系统-客户名称", "B系统-对方户名", "相似度(%)", "匹配状态", "备注"]
        )
        self.table.setRowCount(len(frame))
        for r, (_, row) in enumerate(frame.iterrows()):
            status = str(row["匹配状态"])
            bg = C_GREEN if status in ("完全匹配", "高度匹配") else (
                C_YELLOW if status == "中低匹配" else None
            )
            remark = str(row["备注"])
            cells = (
                (str(row["A系统-客户名称"]), _ALIGN_LEFT, None),
                (str(row["B系统-对方户名"]), _ALIGN_LEFT, None),
                (f"{float(row['相似度(%)']):.1f}", _ALIGN_CENTER, None),
                (status, _ALIGN_CENTER, None),
                (remark, _ALIGN_LEFT,
                 C_RED if status == "低置信度匹配" else "#7A7A7A"),
            )
            for col, (text, align, fg) in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(align)
                if bg:
                    item.setBackground(QColor(bg))
                if fg:
                    item.setForeground(QColor(fg))
                if col == 4:
                    item.setToolTip(remark)
                self.table.setItem(r, col, item)

        hh = self.table.horizontalHeader()
        for i, w in enumerate((250, 250, 90, 100)):
            hh.setSectionResizeMode(i, _TABLE_FIXED)
            self.table.setColumnWidth(i, w)
        hh.setSectionResizeMode(4, _TABLE_STRETCH)

    def _fill_summary(self, result: pipeline.PipelineResult) -> None:
        t = result.thresholds
        c = result.counts
        b_only = len(result.report["frames"]["sheet3"]) if result.report else 0
        self.lbl_summary.setText(
            f"<b>分档结果</b>（A {len(result.df_a)} 行 · B {len(result.df_b)} 行）<br>"
            f"完全匹配（100%）：<b>{c['完全匹配']}</b> 条<br>"
            f"高度匹配（≥{t.high:g}%）：<b>{c['高度匹配']}</b> 条<br>"
            f"中低匹配（{t.low:g}–{t.high:g}%）：<b>{c['中低匹配']}</b> 条<br>"
            f"低置信度（{t.floor:g}–{t.low:g}%）：<b>{c['低置信度匹配']}</b> 条<br>"
            f"A系统独有（<{t.floor:g}%）：<b>{c['A系统独有']}</b> 条<br>"
            f"B系统独有：<b>{b_only}</b> 条<br>"
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

    win = MainWindow()
    win.show()
    print("[selftest] 窗口装配成功")
    a_file, b_file = win.inputs()
    print(f"[selftest] A 系统：{Path(a_file).name}")
    print(f"[selftest] B 系统：{Path(b_file).name}")
    print(f"[selftest] 阈值控件：high={win.spin_high.value()} "
          f"low={win.spin_low.value()} floor={win.spin_floor.value()} "
          f"→ {win._thresholds().label}")

    out_dir = C.OUTPUT_DIR / "_selftest"
    result = pipeline.run_pipeline(
        a_file=a_file, b_file=b_file, thresholds=win._thresholds(),
        out_dir=out_dir, log=lambda t: print(f"[selftest] {t}"),
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
                         str(out_dir))
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
    win.resize(1360, 860)
    shot = out_dir / "gui_preview.png"
    shot.parent.mkdir(parents=True, exist_ok=True)
    win.grab().save(str(shot))
    print(f"[selftest] 界面截图：{shot}")

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
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei", 10))
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
