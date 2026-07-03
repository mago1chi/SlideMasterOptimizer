from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from slidemasteroptimizer.core.pptx_optimizer import (
    AnalysisResult,
    OptimizeResult,
    PptxOptimizationError,
    analyze_pptx,
    default_output_path,
    optimize_pptx,
)


class WorkerThread(QThread):
    analysis_finished = Signal(object)
    optimization_finished = Signal(object)
    failed = Signal(str)

    def __init__(self, action: str, input_path: Path, output_path: Path | None = None) -> None:
        super().__init__()
        self._action = action
        self._input_path = input_path
        self._output_path = output_path

    def run(self) -> None:
        try:
            if self._action == "analyze":
                self.analysis_finished.emit(analyze_pptx(self._input_path))
                return
            if self._output_path is None:
                raise PptxOptimizationError("Output path is not set.")
            self.optimization_finished.emit(optimize_pptx(self._input_path, self._output_path))
        except Exception as exc:  # noqa: BLE001 - surface failures to the GUI log
            self.failed.emit(str(exc))


class DropZone(QFrame):
    file_dropped = Signal(Path)
    rejected = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setObjectName("dropZone")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        label = QLabel("Drop a .pptx file here")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setObjectName("dropLabel")
        sublabel = QLabel("or choose a file with the button below")
        sublabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sublabel.setObjectName("dropSubLabel")

        layout = QVBoxLayout(self)
        layout.addStretch(1)
        layout.addWidget(label)
        layout.addWidget(sublabel)
        layout.addStretch(1)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        path_or_error = self._single_pptx_from_event(event)
        if isinstance(path_or_error, Path):
            event.acceptProposedAction()
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
            return
        event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: ANN001 - Qt event type differs by binding
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        path_or_error = self._single_pptx_from_event(event)
        if isinstance(path_or_error, Path):
            event.acceptProposedAction()
            self.file_dropped.emit(path_or_error)
            return
        self.rejected.emit(path_or_error)
        event.ignore()

    def _single_pptx_from_event(self, event) -> Path | str:  # noqa: ANN001 - Qt event type differs by binding
        urls = event.mimeData().urls()
        if len(urls) != 1:
            return "Drop exactly one .pptx file."
        if not urls[0].isLocalFile():
            return "Only local .pptx files are supported."
        path = Path(urls[0].toLocalFile())
        if not path.is_file():
            return "Folders are not supported."
        if path.suffix.lower() != ".pptx":
            return "Only .pptx files are supported. .ppt and .pptm are not supported."
        return path


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("SlideMasterOptimizer")
        icon_path = Path(__file__).resolve().parents[2] / "assets" / "favicon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.resize(940, 720)
        self._input_path: Path | None = None
        self._analysis: AnalysisResult | None = None
        self._worker: WorkerThread | None = None

        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self.set_input_file)
        self.drop_zone.rejected.connect(self.log)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.choose_button = QPushButton("Choose .pptx")
        self.choose_button.clicked.connect(self.choose_file)
        self.analyze_button = QPushButton("Analyze")
        self.analyze_button.clicked.connect(self.analyze_current_file)
        self.optimize_button = QPushButton("Optimize")
        self.optimize_button.clicked.connect(self.optimize_current_file)
        self.optimize_button.setEnabled(False)

        self.file_name_value = QLabel("-")
        self.slide_count_value = QLabel("-")
        self.master_count_value = QLabel("-")
        self.unused_count_value = QLabel("-")
        self.output_path_value = QLabel("-")
        self.output_path_value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Type", "Part", "Parent / Relationship"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(130)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.addWidget(self.drop_zone)

        picker_layout = QHBoxLayout()
        picker_layout.addWidget(self.file_path_edit, 1)
        picker_layout.addWidget(self.choose_button)
        picker_layout.addWidget(self.analyze_button)
        picker_layout.addWidget(self.optimize_button)
        layout.addLayout(picker_layout)

        stats = QGridLayout()
        stats.addWidget(QLabel("File"), 0, 0)
        stats.addWidget(self.file_name_value, 0, 1)
        stats.addWidget(QLabel("Slides"), 1, 0)
        stats.addWidget(self.slide_count_value, 1, 1)
        stats.addWidget(QLabel("Slide masters"), 2, 0)
        stats.addWidget(self.master_count_value, 2, 1)
        stats.addWidget(QLabel("Removal candidates"), 3, 0)
        stats.addWidget(self.unused_count_value, 3, 1)
        stats.addWidget(QLabel("Output"), 4, 0)
        stats.addWidget(self.output_path_value, 4, 1)
        layout.addLayout(stats)

        layout.addWidget(QLabel("Removal candidates"))
        layout.addWidget(self.table, 1)
        layout.addWidget(QLabel("Log"))
        layout.addWidget(self.log_view)

        self.setStyleSheet(
            """
            QMainWindow { background: #f7f7f5; }
            QLabel { color: #252525; font-size: 13px; }
            QPushButton {
                background: #2f5d50;
                color: white;
                border: 0;
                border-radius: 6px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton:disabled { background: #9ca3a0; }
            QLineEdit, QTextEdit, QTableWidget {
                background: white;
                color: #252525;
                border: 1px solid #d7d7d2;
                border-radius: 6px;
                padding: 6px;
            }
            QHeaderView::section {
                background: #ecece7;
                color: #252525;
                border: 1px solid #d7d7d2;
                padding: 5px;
                font-weight: 600;
            }
            #dropZone {
                min-height: 150px;
                border: 2px dashed #7d8b82;
                border-radius: 8px;
                background: #fbfbf8;
            }
            #dropZone[dragActive="true"] {
                border-color: #2f5d50;
                background: #edf4ef;
            }
            #dropLabel { font-size: 22px; font-weight: 700; }
            #dropSubLabel { color: #646b66; }
            """
        )

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose PowerPoint file",
            "",
            "PowerPoint presentations (*.pptx)",
        )
        if path:
            self.set_input_file(Path(path))

    def set_input_file(self, path: Path) -> None:
        self._input_path = path
        self._analysis = None
        self.file_path_edit.setText(str(path))
        self.file_name_value.setText(path.name)
        self.slide_count_value.setText("-")
        self.master_count_value.setText("-")
        self.unused_count_value.setText("-")
        self.output_path_value.setText(str(default_output_path(path)))
        self.table.setRowCount(0)
        self.optimize_button.setEnabled(False)
        self.log(f"Selected: {path}")
        self.analyze_current_file()

    def analyze_current_file(self) -> None:
        if self._input_path is None:
            self.log("Choose or drop a .pptx file first.")
            return
        self._set_busy(True)
        self.log("Analyzing presentation...")
        self._worker = WorkerThread("analyze", self._input_path)
        self._worker.analysis_finished.connect(self._analysis_finished)
        self._worker.failed.connect(self._worker_failed)
        self._worker.finished.connect(lambda: self._set_busy(False))
        self._worker.start()

    def optimize_current_file(self) -> None:
        if self._input_path is None or self._analysis is None:
            self.log("Analyze a .pptx file before optimizing.")
            return
        output_path = default_output_path(self._input_path)
        if output_path.exists():
            result = QMessageBox.question(
                self,
                "Overwrite output?",
                f"{output_path.name} already exists. Overwrite it?",
            )
            if result != QMessageBox.StandardButton.Yes:
                return

        self._set_busy(True)
        self.log("Optimizing presentation...")
        self._worker = WorkerThread("optimize", self._input_path, output_path)
        self._worker.optimization_finished.connect(self._optimization_finished)
        self._worker.failed.connect(self._worker_failed)
        self._worker.finished.connect(lambda: self._set_busy(False))
        self._worker.start()

    def _analysis_finished(self, result: AnalysisResult) -> None:
        self._analysis = result
        self.slide_count_value.setText(str(result.slide_count))
        self.master_count_value.setText(
            f"{result.total_masters} ({result.used_master_count} used)"
        )
        self.unused_count_value.setText(
            f"{result.removal_candidate_count} "
            f"({result.unused_master_count} masters, {result.unused_layout_count} layouts)"
        )
        self.table.setRowCount(result.removal_candidate_count)
        row = 0
        for row, candidate in enumerate(result.unused_masters):
            self.table.setItem(row, 0, QTableWidgetItem("Slide master"))
            self.table.setItem(row, 1, QTableWidgetItem(candidate.part_name))
            self.table.setItem(row, 2, QTableWidgetItem(candidate.relationship_id or "-"))
        row = len(result.unused_masters)
        for offset, candidate in enumerate(result.unused_layouts):
            table_row = row + offset
            self.table.setItem(table_row, 0, QTableWidgetItem("Slide layout"))
            self.table.setItem(table_row, 1, QTableWidgetItem(candidate.part_name))
            self.table.setItem(
                table_row,
                2,
                QTableWidgetItem(f"{candidate.master_part_name} / {candidate.relationship_id or '-'}"),
            )
        self.table.resizeColumnsToContents()
        self.optimize_button.setEnabled(result.removal_candidate_count > 0 and not result.warnings)
        if result.warnings:
            for warning in result.warnings:
                self.log(f"Warning: {warning}")
            self.log("Optimization is disabled because unresolved references were found.")
        elif result.removal_candidate_count:
            self.log(f"Found {result.unused_master_count} unused slide master(s).")
            self.log(f"Found {result.unused_layout_count} unused slide layout(s).")
        else:
            self.log("No removal candidates were found.")

    def _optimization_finished(self, result: OptimizeResult) -> None:
        self.output_path_value.setText(str(result.output_path))
        self.log(
            "Done. Removed "
            f"{result.removed_master_count} slide master(s) and "
            f"{result.removed_layout_count} layout(s)."
        )

    def _worker_failed(self, message: str) -> None:
        self.optimize_button.setEnabled(False)
        self.log(f"Error: {message}")

    def _set_busy(self, busy: bool) -> None:
        self.choose_button.setEnabled(not busy)
        self.analyze_button.setEnabled(not busy and self._input_path is not None)
        can_optimize = (
            not busy
            and self._analysis is not None
            and self._analysis.removal_candidate_count > 0
            and not self._analysis.warnings
        )
        self.optimize_button.setEnabled(can_optimize)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor) if busy else QApplication.restoreOverrideCursor()

    def log(self, message: str) -> None:
        self.log_view.append(message)
