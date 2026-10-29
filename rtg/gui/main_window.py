"""Main window: project handling plus the Module/Scenario/Test Case/Step tree."""

from __future__ import annotations

import html
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import QObject, QSettings, QSize, Qt, Signal
from PySide6.QtGui import QAction, QColor, QFont, QGuiApplication, QIcon, QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTextBrowser,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .. import __app_name__, __version__
from .. import capture as capture_api
from ..capture import MODE_MONITOR, MODE_REGION, CaptureError
from ..config import load_config
from ..errors import MissingRequiredColumnsError, ResultGeneratorError
from ..excel_parser import inspect_workbook
from ..excel_writer import ExcelWriteError, default_result_name
from ..hotkey import HotkeyError, HotkeyListener, format_hotkey
from ..pdf_export import PdfExportError, available_converters, docx_to_pdf
from ..project import Project, ProjectError
from .capture_dialogs import (
    CaptureSettingsDialog,
    CaptureTargetDialog,
    RegionSelector,
    pil_to_qpixmap,
)
from . import theme
from .icons import clear_cache, icon
from .results_panel import ResultsPanel
from .dialogs import (
    ExportDialog,
    MetadataDialog,
    NewProjectDialog,
    OrphanDialog,
    ScenarioChangesDialog,
    SheetPickerDialog,
)

NODE_ROLE = Qt.ItemDataRole.UserRole
UID_ROLE = Qt.ItemDataRole.UserRole + 1
FILE_ROLE = Qt.ItemDataRole.UserRole + 2

THUMB = QSize(192, 108)

#: Let the compositor actually remove our window before grabbing the screen.
HIDE_DELAY_SECONDS = 0.25


class HotkeyBridge(QObject):
    """Hands the hotkey over from pynput's thread to the Qt thread.

    pynput fires its callback on its own listener thread; touching widgets
    from there would be a crash waiting to happen. Emitting a signal makes Qt
    queue the call onto the GUI thread instead.
    """

    triggered = Signal()

    def fire(self) -> None:
        self.triggered.emit()


def _open_in_explorer(path: Path) -> None:
    """Show a file or folder in the OS file manager."""
    try:
        if sys.platform.startswith("win"):
            if path.is_dir():
                os.startfile(str(path))  # noqa: S606 - user initiated
            else:
                subprocess.run(["explorer", "/select,", str(path)], check=False)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except OSError:
        pass


class MainWindow(QMainWindow):
    """The application window."""

    def __init__(self, project: Optional[Project] = None) -> None:
        super().__init__()
        self.setWindowTitle(__app_name__)
        self.resize(1280, 800)
        self.settings = QSettings("ResultTestGenerator", "app")
        self.project: Optional[Project] = None
        self._listener: Optional[HotkeyListener] = None
        self._bridge = HotkeyBridge()
        self._bridge.triggered.connect(self.capture_now)
        self._capturing = False
        self._last_uid: Optional[str] = None

        self._build_ui()
        self._build_actions()
        self._update_enabled()
        if project is not None:
            self._set_project(project)
        self._restore_window_state()
        self.apply_icons()

    # -- construction ------------------------------------------------------ #

    def _build_ui(self) -> None:
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left.setObjectName("sidebarPane")
        left.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        left_layout.setContentsMargins(10, 10, 6, 6)
        left_layout.setSpacing(8)

        self.search = QLineEdit()
        self.search.setObjectName("search")
        self.search.setPlaceholderText("Cari test case atau langkah")
        self.search.setClearButtonEnabled(True)
        self.search_action = self.search.addAction(
            QIcon(), QLineEdit.ActionPosition.LeadingPosition
        )
        self.search.textChanged.connect(self._apply_filter)
        left_layout.addWidget(self.search)

        self.tree = QTreeWidget()
        self.tree.setObjectName("sidebar")
        self.tree.setColumnCount(3)
        self.tree.setHeaderLabels(["Module / Scenario / Test Case / Step", "Screenshot", "Keterangan"])
        self.tree.setHeaderHidden(True)
        self.tree.setFrameShape(QTreeWidget.Shape.NoFrame)
        self.tree.setIndentation(14)
        self.tree.setAnimated(True)
        self.tree.setIconSize(QSize(16, 16))
        self.tree.setAlternatingRowColors(False)
        self.tree.setUniformRowHeights(True)
        self.tree.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tree.currentItemChanged.connect(lambda *_: self._on_tree_selection())
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_context_menu)
        header = self.tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, header.ResizeMode.Stretch)
        header.setSectionResizeMode(1, header.ResizeMode.Fixed)
        header.resizeSection(1, 34)
        self.tree.setColumnHidden(2, True)      # kept for tooltips and tests
        left_layout.addWidget(self.tree, 1)
        splitter.addWidget(left)

        self.tabs = QTabWidget()
        self.tabs.tabBar().setExpanding(False)
        self.tabs.tabBar().setDrawBase(False)

        self.detail = QTextBrowser()
        self.detail.setOpenLinks(False)
        self.detail.document().setDocumentMargin(14)
        self.tabs.addTab(self.detail, "Detail")

        shots = QWidget()
        shots_layout = QVBoxLayout(shots)
        shots_layout.setContentsMargins(12, 0, 12, 10)
        self.shot_label = QLabel("Pilih sebuah test case di tree.")
        self.shot_label.setProperty("role", "secondary")
        self.shot_label.setWordWrap(True)
        shots_layout.addWidget(self.shot_label)

        self.shot_list = QListWidget()
        self.shot_list.setViewMode(QListWidget.ViewMode.IconMode)
        self.shot_list.setIconSize(THUMB)
        self.shot_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.shot_list.setMovement(QListWidget.Movement.Snap)
        self.shot_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.shot_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.shot_list.setSpacing(12)
        self.shot_list.setWordWrap(True)
        self.shot_list.model().rowsMoved.connect(lambda *_: self._persist_shot_order())
        self.shot_list.itemDoubleClicked.connect(self._preview_shot)
        shots_layout.addWidget(self.shot_list, 1)

        buttons = QHBoxLayout()
        buttons.setSpacing(2)
        self.btn_add_shot = _tool_button("Tambah dari file...", self._add_screenshot_from_file)
        buttons.addWidget(self.btn_add_shot)
        self.btn_caption = _tool_button("Caption...", self._edit_caption)
        buttons.addWidget(self.btn_caption)
        self.btn_move_shot = _tool_button("Pindah test case...", self._move_screenshot)
        buttons.addWidget(self.btn_move_shot)
        buttons.addStretch(1)
        self.btn_delete_shot = _tool_button("Hapus", self._delete_screenshot)
        buttons.addWidget(self.btn_delete_shot)
        shots_layout.addLayout(buttons)
        self.tabs.addTab(shots, "Screenshot")

        self.results = ResultsPanel()
        self.tabs.addTab(self.results, "Hasil")

        self.warnings = QPlainTextEdit()
        self.warnings.setReadOnly(True)
        self.tabs.addTab(self.warnings, "Peringatan")

        splitter.addWidget(self.tabs)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)
        splitter.setHandleWidth(1)
        splitter.setChildrenCollapsible(False)
        left.setMinimumWidth(280)
        splitter.setSizes([420, 860])
        self.splitter = splitter

        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_welcome())
        self.stack.addWidget(splitter)
        self.setCentralWidget(self.stack)

        self.status = self.statusBar()
        self.status.setSizeGripEnabled(False)
        self.progress_label = QLabel("")
        self.progress_label.setProperty("role", "secondary")
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progress")
        self.progress_bar.setFixedWidth(120)
        self.progress_bar.setTextVisible(False)
        self.status.addPermanentWidget(self.progress_label)
        self.status.addPermanentWidget(self.progress_bar)
        self.progress_label.hide()
        self.progress_bar.hide()
        self.status.showMessage("Belum ada project dibuka.")

    def _build_welcome(self) -> QWidget:
        """Empty state shown until a project is opened."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(8)

        self.welcome_icon = QLabel()
        self.welcome_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.welcome_icon)

        title = QLabel(__app_name__)
        title.setProperty("role", "title")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Buat dokumen Result Test UAT dari skenario Excel.")
        subtitle.setProperty("role", "secondary")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)
        layout.addSpacing(12)

        row = QHBoxLayout()
        row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        new_button = QPushButton("Project baru")
        new_button.setProperty("primary", True)
        new_button.clicked.connect(lambda: self.new_project())
        open_button = QPushButton("Buka project")
        open_button.clicked.connect(lambda: self.open_project())
        row.addWidget(new_button)
        row.addWidget(open_button)
        layout.addLayout(row)

        last = str(self.settings.value("last_project", "") or "")
        self.welcome_last = QPushButton(f"Buka terakhir: {Path(last).name}" if last else "")
        self.welcome_last.setFlat(True)
        self.welcome_last.setProperty("role", "secondary")
        self.welcome_last.clicked.connect(lambda: self._open_folder(last))
        self.welcome_last.setVisible(bool(last) and Path(last).is_dir())
        layout.addWidget(self.welcome_last, 0, Qt.AlignmentFlag.AlignCenter)
        return page

    def _build_actions(self) -> None:
        self.act_new = QAction("Project &baru...", self)
        self.act_new.setShortcut(QKeySequence.StandardKey.New)
        self.act_new.triggered.connect(self.new_project)

        self.act_open = QAction("&Buka project...", self)
        self.act_open.setShortcut(QKeySequence.StandardKey.Open)
        self.act_open.triggered.connect(self.open_project)

        self.act_save = QAction("&Simpan project", self)
        self.act_save.setShortcut(QKeySequence.StandardKey.Save)
        self.act_save.triggered.connect(self.save_project)

        self.act_reimport = QAction("&Re-import Excel", self)
        self.act_reimport.setShortcut("F5")
        self.act_reimport.triggered.connect(self.reimport)

        self.act_sheets = QAction("Pilih &module (sheet)...", self)
        self.act_sheets.triggered.connect(self.choose_sheets)

        self.act_metadata = QAction("&Metadata && judul module...", self)
        self.act_metadata.triggered.connect(self.edit_metadata)

        self.act_orphans = QAction("Screenshot &orphan...", self)
        self.act_orphans.triggered.connect(self.show_orphans)

        self.act_review = QAction("&Review perubahan skenario...", self)
        self.act_review.triggered.connect(self.review_changes)

        self.act_capture = QAction("&Ambil screenshot sekarang", self)
        self.act_capture.triggered.connect(self.capture_now)

        self.act_capture_settings = QAction("&Pengaturan screenshot...", self)
        self.act_capture_settings.triggered.connect(self.edit_capture_settings)

        self.act_generate = QAction("&Generate dokumen Word...", self)
        self.act_generate.setShortcut("Ctrl+G")
        self.act_generate.triggered.connect(self.generate)

        self.act_write_excel = QAction("Tulis hasil ke &salinan Excel...", self)
        self.act_write_excel.setShortcut("Ctrl+E")
        self.act_write_excel.triggered.connect(self.write_result_excel)

        self.act_export_pdf = QAction("Export &PDF...", self)
        self.act_export_pdf.triggered.connect(self.export_pdf)

        self.act_open_folder = QAction("Buka &folder project", self)
        self.act_open_folder.triggered.connect(
            lambda: self.project and _open_in_explorer(self.project.folder)
        )

        self.act_quit = QAction("&Keluar", self)
        self.act_quit.setShortcut(QKeySequence.StandardKey.Quit)
        self.act_quit.triggered.connect(self.close)

        self.act_about = QAction("&Tentang", self)
        self.act_about.triggered.connect(self.about)

        menu = self.menuBar()
        file_menu = menu.addMenu("&File")
        file_menu.addAction(self.act_new)
        file_menu.addAction(self.act_open)
        file_menu.addAction(self.act_save)
        file_menu.addSeparator()
        file_menu.addAction(self.act_reimport)
        file_menu.addAction(self.act_sheets)
        file_menu.addSeparator()
        file_menu.addAction(self.act_open_folder)
        file_menu.addSeparator()
        file_menu.addAction(self.act_quit)

        edit_menu = menu.addMenu("&Project")
        edit_menu.addAction(self.act_metadata)
        edit_menu.addAction(self.act_orphans)
        edit_menu.addAction(self.act_review)

        shot_menu = menu.addMenu("&Screenshot")
        shot_menu.addAction(self.act_capture)
        shot_menu.addAction(self.act_capture_settings)

        export_menu = menu.addMenu("&Export")
        export_menu.addAction(self.act_generate)
        export_menu.addAction(self.act_export_pdf)
        export_menu.addSeparator()
        export_menu.addAction(self.act_write_excel)

        help_menu = menu.addMenu("&Bantuan")
        help_menu.addAction(self.act_about)

        for action, tip in (
            (self.act_new, "Project baru"), (self.act_open, "Buka project"),
            (self.act_save, "Simpan project"), (self.act_reimport, "Re-import Excel"),
            (self.act_metadata, "Metadata & judul module"),
        ):
            shortcut = action.shortcut().toString(QKeySequence.SequenceFormat.NativeText)
            action.setToolTip(f"{tip} ({shortcut})" if shortcut else tip)

        toolbar = self.addToolBar("Utama")
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setIconSize(QSize(18, 18))
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        for action in (self.act_new, self.act_open, self.act_save):
            toolbar.addAction(action)
        toolbar.addSeparator()
        for action in (self.act_reimport, self.act_metadata):
            toolbar.addAction(action)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        self.tb_capture = QToolButton()
        self.tb_capture.setText("Capture")
        self.tb_capture.setToolTip("Ambil screenshot sekarang")
        self.tb_capture.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.tb_capture.setProperty("soft", True)
        self.tb_capture.clicked.connect(lambda: self.capture_now())
        toolbar.addWidget(self.tb_capture)

        self.tb_generate = QToolButton()
        self.tb_generate.setText("Generate")
        self.tb_generate.setToolTip("Generate dokumen Word (Ctrl+G)")
        self.tb_generate.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.tb_generate.setProperty("primary", True)
        self.tb_generate.clicked.connect(lambda: self.generate())
        toolbar.addWidget(self.tb_generate)

    def apply_icons(self) -> None:
        """(Re-)apply icons for the current light/dark theme."""
        clear_cache()
        t = theme.tokens()
        for action, name in (
            (self.act_new, "new"), (self.act_open, "open"), (self.act_save, "save"),
            (self.act_reimport, "reimport"), (self.act_sheets, "module"),
            (self.act_metadata, "metadata"), (self.act_orphans, "orphan"),
            (self.act_review, "review"), (self.act_capture, "capture"),
            (self.act_capture_settings, "settings"), (self.act_generate, "generate"),
            (self.act_write_excel, "excel"), (self.act_export_pdf, "pdf"),
            (self.act_open_folder, "folder"), (self.act_about, "info"),
        ):
            action.setIcon(icon(name))
            action.setIconVisibleInMenu(False)      # icons live in the toolbar only
        self.tb_capture.setIcon(icon("capture", t["accent"]))
        self.tb_generate.setIcon(icon("generate", t["on_accent"]))
        for button, name, color in (
            (self.btn_add_shot, "add", None), (self.btn_caption, "caption", None),
            (self.btn_move_shot, "move", None), (self.btn_delete_shot, "delete", t["danger"]),
        ):
            button.setIcon(icon(name, color))
        self.search_action.setIcon(icon("search"))
        self.results.apply_icons()
        pixmap = icon("generate", t["accent"]).pixmap(QSize(64, 64))
        self.welcome_icon.setPixmap(pixmap)
        for widget in (self.tb_capture, self.tb_generate):
            widget.style().unpolish(widget)
            widget.style().polish(widget)
        self._refresh_tree()
        self._refresh_status()

    # -- project ----------------------------------------------------------- #

    def _preferred_capture(self) -> Dict[str, object]:
        """The hotkey and capture mode the user chose last.

        Stored per user, not per project: changing the hotkey once should hold
        for every project created afterwards.
        """
        return {
            "hotkey": str(self.settings.value("hotkey", "") or "") or None,
            "capture_mode": str(self.settings.value("capture_mode", "") or "") or None,
            "capture_monitor": int(self.settings.value("capture_monitor", 0) or 0),
        }

    def _remember_capture(self, project: Project) -> None:
        self.settings.setValue("hotkey", project.hotkey)
        self.settings.setValue("capture_mode", project.capture_mode)
        self.settings.setValue("capture_monitor", project.capture_monitor)

    def _last_dir(self) -> Path:
        value = self.settings.value("last_dir", "")
        return Path(value) if value else Path.home()

    def _remember_dir(self, path: Path) -> None:
        self.settings.setValue("last_dir", str(Path(path).parent))

    def _set_project(self, project: Project) -> None:
        self.project = project
        self._last_uid = None
        self.setWindowTitle(f"{__app_name__} - {project.name}")
        self.settings.setValue("last_project", str(project.folder))
        self._refresh_all()
        self._update_enabled()
        self._start_hotkey()

    def _update_enabled(self) -> None:
        has = self.project is not None
        self.stack.setCurrentIndex(1 if has else 0)
        for button in (self.tb_capture, self.tb_generate):
            button.setEnabled(has)
        for action in (
            self.act_save, self.act_reimport, self.act_sheets, self.act_metadata,
            self.act_orphans, self.act_review, self.act_generate, self.act_open_folder,
            self.act_capture, self.act_capture_settings,
            self.act_write_excel, self.act_export_pdf,
        ):
            action.setEnabled(has)
        self._update_shot_buttons()

    def new_project(self) -> None:
        dialog = NewProjectDialog(parent=self, start_dir=self._last_dir())
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        excel = dialog.excel_path
        folder = dialog.folder
        self._remember_dir(excel)

        try:
            with _busy():
                inspection = inspect_workbook(excel)
        except ResultGeneratorError as exc:
            self._error("Gagal membaca file Excel", exc)
            return

        if not inspection.scenario_sheets and not inspection.mapping_candidates:
            QMessageBox.warning(
                self, "Tidak ada sheet skenario",
                "File ini tidak punya sheet dengan kolom Test Case ID, Test Case "
                "dan Test Step. Periksa lagi file skenarionya.",
            )
            return

        config = load_config(search_from=folder)
        picker = SheetPickerDialog(
            inspection,
            config,
            reinspect=lambda overrides: inspect_workbook(
                excel, config=config, overrides=overrides
            ),
            parent=self,
        )
        if picker.exec() != QDialog.DialogCode.Accepted:
            return

        try:
            with _busy():
                project = Project.create(
                    folder,
                    excel,
                    sheets=picker.selected_sheets(),
                    overrides={n: o.to_dict() for n, o in picker.overrides.items()},
                    name=dialog.project_name,
                    **self._preferred_capture(),
                )
        except (ResultGeneratorError, OSError) as exc:
            self._error("Gagal membuat project", exc)
            return
        self._set_project(project)
        self.status.showMessage(f"Project dibuat di {project.folder}", 8000)

    def open_project(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Buka folder project", str(self._last_dir())
        )
        if not folder:
            return
        self._open_folder(folder)

    def _open_folder(self, folder: str) -> None:
        if not folder:
            return
        try:
            with _busy():
                project = Project.open(Path(folder))
        except ResultGeneratorError as exc:
            self._error("Gagal membuka project", exc)
            return
        self.settings.setValue("last_dir", folder)
        self._set_project(project)
        self.status.showMessage(f"Project dibuka: {project.folder}", 8000)
        self.announce_scenario_changes()

    def save_project(self) -> None:
        if self.project is None:
            return
        try:
            self.project.save()
        except OSError as exc:
            self._error("Gagal menyimpan project", exc)
            return
        self.status.showMessage("Project disimpan.", 4000)

    def choose_sheets(self) -> None:
        if self.project is None:
            return
        project = self.project
        try:
            with _busy():
                inspection = project.inspect()
        except ResultGeneratorError as exc:
            self._error("Gagal membaca file Excel", exc)
            return
        picker = SheetPickerDialog(
            inspection,
            project.config,
            reinspect=lambda overrides: inspect_workbook(
                project.excel_path, config=project.config, overrides=overrides
            ),
            selected=project.selected_sheets,
            overrides=project.column_overrides,
            parent=self,
        )
        if picker.exec() != QDialog.DialogCode.Accepted:
            return
        self._do_reimport(
            sheets=picker.selected_sheets(),
            overrides={n: o.to_dict() for n, o in picker.overrides.items()},
        )

    def reimport(self) -> None:
        self._do_reimport()

    def _do_reimport(self, **kwargs) -> None:
        if self.project is None:
            return
        try:
            with _busy():
                report = self.project.reimport(**kwargs)
        except MissingRequiredColumnsError as exc:
            self._error(
                "Kolom wajib tidak ditemukan",
                exc,
                hint="Buka File > Pilih module (sheet) lalu pakai 'Mapping kolom manual'.",
            )
            return
        except ResultGeneratorError as exc:
            self._error("Gagal membaca ulang file Excel", exc)
            return

        self._refresh_all()
        QMessageBox.information(
            self, "Re-import selesai", "\n".join(report.summary_lines())
        )
        if report.changes is not None and report.changes.needs_attention:
            self._review(report.changes)

    def announce_scenario_changes(self) -> None:
        """After opening: tell the user once if the workbook changed meanwhile."""
        project = self.project
        if project is None or not project.excel_changed:
            return
        changes = project.last_changes
        if changes is None or not changes.needs_attention:
            QMessageBox.information(
                self, "File skenario berubah",
                "File skenario berubah sejak project ini terakhir dibuka, tetapi "
                "tidak ada bukti yang terpengaruh.",
            )
            return
        QMessageBox.information(
            self, "File skenario berubah",
            "File skenario berubah sejak project ini terakhir dibuka.\n\n"
            + "\n".join(changes.summary_lines()),
        )
        self._review(changes)

    def review_changes(self) -> None:
        if self.project is None:
            return
        try:
            changes = self.project.check_changes()
        except ResultGeneratorError as exc:
            self._error("Gagal membandingkan skenario", exc)
            return
        if not changes.needs_attention:
            QMessageBox.information(
                self, "Review perubahan skenario",
                "Tidak ada perubahan skenario yang perlu direview.",
            )
            self._refresh_all()
            return
        self._review(changes)

    def _review(self, changes) -> None:
        project = self.project
        if project is None:
            return
        dialog = ScenarioChangesDialog(project, changes, parent=self)
        accepted = dialog.exec() == QDialog.DialogCode.Accepted
        if dialog.undo_requested:
            try:
                with _busy():
                    project.restore_backup()
            except ResultGeneratorError as exc:
                self._error("Gagal mengembalikan backup", exc)
                return
            self.status.showMessage("Re-import dibatalkan; project dikembalikan.", 6000)
        elif accepted:
            try:
                dialog.apply()
            except (ResultGeneratorError, OSError) as exc:
                self._error("Gagal menerapkan perubahan", exc)
                return
            self.status.showMessage("Perubahan skenario diterapkan.", 6000)
        self._refresh_all()

    def mark_reviewed(self, uid: str) -> None:
        if self.project is None:
            return
        self.project.acknowledge(uid)
        self._refresh_all()

    def _tree_context_menu(self, pos) -> None:
        item = self.tree.itemAt(pos)
        if item is None or item.data(0, NODE_ROLE) != "case":
            return
        uid = item.data(0, UID_ROLE)
        if self.project is None or uid not in self.project.needs_review:
            return
        menu = QMenu(self)
        action = menu.addAction("Tandai sudah direview")
        if menu.exec(self.tree.viewport().mapToGlobal(pos)) is action:
            self.mark_reviewed(uid)

    def edit_metadata(self) -> None:
        if self.project is None:
            return
        dialog = MetadataDialog(self.project, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        dialog.apply()
        self.project.save()
        self._refresh_tree()
        self.status.showMessage("Metadata & judul module disimpan.", 4000)

    def show_orphans(self) -> None:
        if self.project is None:
            return
        if not self.project.orphan_uids:
            QMessageBox.information(
                self, "Tidak ada orphan",
                "Semua screenshot masih menempel pada test case yang ada.",
            )
            return
        OrphanDialog(self.project, parent=self).exec()
        self.project.save()
        self._refresh_all()

    def generate(self) -> None:
        if self.project is None:
            return
        if self.project.needs_review:
            answer = QMessageBox.warning(
                self, "Ada test case yang perlu dicek",
                f"{len(self.project.needs_review)} test case berubah sejak bukti "
                "diambil dan belum direview, jadi screenshot-nya mungkin sudah tidak "
                "sesuai.\n\nTetap generate dokumen?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        dialog = ExportDialog(self.project, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        dialog.apply()
        try:
            with _busy():
                results = self.project.generate(modules=dialog.selected_modules())
                self.project.save()
        except (ResultGeneratorError, OSError) as exc:
            self._error("Gagal membuat dokumen", exc)
            return

        lines = [f"{r.path.name}  ({r.test_case_count} test case, "
                 f"{r.screenshot_count} screenshot)" for r in results]
        missing = sum(len(r.without_screenshot) for r in results)
        if missing:
            lines.append("")
            lines.append(f"{missing} test case masih bertanda '[Belum ada screenshot]'.")
        box = QMessageBox(self)
        box.setWindowTitle("Dokumen selesai dibuat")
        box.setText("\n".join(lines))
        box.setInformativeText(f"Folder: {self.project.output_dir}")
        open_button = box.addButton("Buka folder", QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Ok)
        box.exec()
        if box.clickedButton() is open_button:
            _open_in_explorer(self.project.output_dir)

    # -- global hotkey ----------------------------------------------------- #

    def _start_hotkey(self) -> None:
        """(Re-)register the project's hotkey, reporting failures once."""
        self._stop_hotkey()
        if self.project is None:
            return
        listener = HotkeyListener(self.project.hotkey, self._bridge.fire)
        try:
            listener.start()
        except HotkeyError as exc:
            self._error(
                "Hotkey tidak bisa didaftarkan", exc,
                hint="Screenshot tetap bisa diambil lewat menu Screenshot > "
                     "Ambil screenshot sekarang.",
            )
            return
        self._listener = listener
        self._refresh_status()

    def _stop_hotkey(self) -> None:
        if self._listener is not None:
            self._listener.stop()
            self._listener = None

    def edit_capture_settings(self) -> None:
        if self.project is None:
            return
        try:
            monitors = capture_api.list_monitors()
        except CaptureError:
            monitors = []
        dialog = CaptureSettingsDialog(self.project, monitors, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        dialog.apply()
        self.project.save()
        self._remember_capture(self.project)
        self._start_hotkey()
        self.status.showMessage(
            f"Hotkey: {self.project.hotkey}  |  mode: {self.project.capture_mode}", 5000
        )

    def _restore_window_state(self) -> None:
        try:
            geometry = self.settings.value("geometry")
            if geometry:
                self.restoreGeometry(geometry)
            state = self.settings.value("splitter")
            if state:
                self.splitter.restoreState(state)
        except (TypeError, ValueError):
            pass

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt naming
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("splitter", self.splitter.saveState())
        self._stop_hotkey()
        super().closeEvent(event)

    # -- capture ----------------------------------------------------------- #

    def _grab(self) -> Tuple[Optional[object], Optional[QPixmap]]:
        """Grab according to the project's capture mode.

        For the region mode the whole desktop is captured first and the
        selection is cropped from that frozen image, so the overlay itself can
        never appear in the screenshot.
        """
        project = self.project
        assert project is not None
        mode = project.capture_mode
        if mode == MODE_REGION:
            full = capture_api.grab(0)
            geometry = QGuiApplication.primaryScreen().virtualGeometry()
            selector = RegionSelector(pil_to_qpixmap(full), geometry)
            if selector.exec() != QDialog.DialogCode.Accepted:
                return None, None
            box = selector.selected_box()
            if box is None:
                return None, None
            scale = full.width / max(1, geometry.width())
            image = capture_api.crop(full, box, scale=scale)
        elif mode == MODE_MONITOR:
            image = capture_api.grab(project.capture_monitor or 1)
        else:
            image = capture_api.grab(0)
        return image, pil_to_qpixmap(image)

    def capture_now(self) -> None:
        """Take a screenshot and ask which test case it belongs to."""
        project = self.project
        if project is None or self._capturing:
            return
        if not project.test_cases:
            QMessageBox.information(
                self, "Belum ada test case",
                "Project ini belum punya test case, jadi screenshot tidak punya "
                "tujuan. Pilih module dulu lewat File > Pilih module (sheet).",
            )
            return

        self._capturing = True
        was_visible = self.isVisible() and not self.isMinimized()
        try:
            if was_visible:
                self.hide()
                QApplication.processEvents()
                time.sleep(HIDE_DELAY_SECONDS)
            try:
                image, pixmap = self._grab()
            except CaptureError as exc:
                self._restore(was_visible)
                self._error("Gagal mengambil screenshot", exc)
                return
            if image is None:
                self._restore(was_visible)       # region selection cancelled
                return

            dialog = CaptureTargetDialog(project, pixmap, last_uid=self._last_uid)
            accepted = dialog.exec() == QDialog.DialogCode.Accepted
            self._restore(was_visible)
            if not accepted or not dialog.chosen_uid:
                self.status.showMessage("Screenshot dibuang.", 4000)
                return

            uid = dialog.chosen_uid
            handle, temp_name = tempfile.mkstemp(suffix=".png", prefix="rtg-shot-")
            os.close(handle)
            temp_path = Path(temp_name)
            try:
                capture_api.save_png(image, temp_path)
                entry = project.add_screenshot(
                    uid, temp_path, caption=dialog.chosen_caption, move=True
                )
            except ResultGeneratorError as exc:
                temp_path.unlink(missing_ok=True)
                self._error("Gagal menyimpan screenshot", exc)
                return
            finally:
                temp_path.unlink(missing_ok=True)

            self._last_uid = uid
            case = project.test_case(uid)
            self._refresh_tree()
            self._refresh_status()
            self._select_case(uid)
            self.status.showMessage(
                f"Screenshot disimpan ke {case.display_label if case else uid} "
                f"({Path(entry.file).name})", 6000
            )
        finally:
            self._capturing = False

    def _restore(self, was_visible: bool) -> None:
        if was_visible:
            self.show()
            self.raise_()
            self.activateWindow()

    def _select_case(self, uid: str) -> None:
        """Point the tree at a test case, so its screenshots show up."""
        for i in range(self.tree.topLevelItemCount()):
            module_item = self.tree.topLevelItem(i)
            for j in range(module_item.childCount()):
                group_item = module_item.child(j)
                for k in range(group_item.childCount()):
                    case_item = group_item.child(k)
                    if case_item.data(0, UID_ROLE) == uid:
                        self.tree.setCurrentItem(case_item)
                        self.tree.scrollToItem(case_item)
                        self.tabs.setCurrentIndex(1)
                        return

    def write_result_excel(self) -> None:
        """Write the recorded results into a copy of the scenario workbook."""
        project = self.project
        if project is None:
            return
        if not project.results_count():
            QMessageBox.information(
                self, "Belum ada hasil",
                "Belum ada Actual Results / Pass-Fail yang diisi. Buka tab "
                "'Hasil' pada sebuah test case untuk mengisinya.",
            )
            return

        project.results_dir.mkdir(parents=True, exist_ok=True)
        suggestion = project.results_dir / default_result_name(project.excel_path)
        path, _filter = QFileDialog.getSaveFileName(
            self, "Simpan salinan Excel berisi hasil", str(suggestion),
            "Excel (*.xlsx *.xlsm)",
        )
        if not path:
            return
        try:
            with _busy():
                report = project.write_result_workbook(Path(path), overwrite=True)
        except (ResultGeneratorError, ExcelWriteError, OSError) as exc:
            self._error("Gagal menulis hasil ke Excel", exc)
            return

        box = QMessageBox(self)
        box.setWindowTitle("Hasil ditulis ke salinan Excel")
        box.setText("\n".join(report.summary_lines()))
        box.setInformativeText(
            "File skenario asli tidak diubah sama sekali."
        )
        open_button = box.addButton("Buka folder", QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Ok)
        box.exec()
        if box.clickedButton() is open_button:
            _open_in_explorer(report.path.parent)

    def export_pdf(self) -> None:
        """Convert the .docx files in the project's output folder to PDF."""
        project = self.project
        if project is None:
            return
        documents = sorted(project.output_dir.glob("*.docx"))
        if not documents:
            QMessageBox.information(
                self, "Belum ada dokumen",
                "Folder output belum berisi file .docx. Jalankan "
                "Export > Generate dokumen Word dulu.",
            )
            return

        usable = [info for info in available_converters() if info.available]
        if not usable:
            detail = "\n".join(f"- {i.label}: {i.detail}" for i in available_converters())
            QMessageBox.warning(
                self, "Tidak ada converter PDF",
                "Export PDF butuh Microsoft Word (plus pywin32) atau LibreOffice.\n\n"
                + detail,
            )
            return

        names = "\n".join(f"- {doc.name}" for doc in documents)
        answer = QMessageBox.question(
            self, "Export PDF",
            f"Ubah {len(documents)} dokumen menjadi PDF memakai "
            f"{usable[0].label}?\n\n{names}",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        produced: List[Path] = []
        try:
            with _busy():
                for document in documents:
                    produced.append(
                        docx_to_pdf(document, converter=usable[0].name)
                    )
        except (PdfExportError, ResultGeneratorError) as exc:
            self._error("Gagal membuat PDF", exc)
            return

        box = QMessageBox(self)
        box.setWindowTitle("PDF selesai")
        box.setText("\n".join(f"{p.name}  ({p.stat().st_size // 1024} KB)" for p in produced))
        box.setInformativeText(f"Folder: {project.output_dir}")
        open_button = box.addButton("Buka folder", QMessageBox.ButtonRole.ActionRole)
        box.addButton(QMessageBox.StandardButton.Ok)
        box.exec()
        if box.clickedButton() is open_button:
            _open_in_explorer(project.output_dir)

    def about(self) -> None:
        QMessageBox.about(
            self, f"Tentang {__app_name__}",
            f"<b>{__app_name__}</b> versi {__version__}<br><br>"
            "Membuat dokumen Result Test UAT (Word) dari file skenario Excel, "
            "lengkap dengan screenshot per test case.<br><br>"
            "File Excel asli tidak pernah diubah.",
        )

    # -- tree -------------------------------------------------------------- #

    def _refresh_all(self) -> None:
        self.results.set_test_case(self.project, None)
        self._refresh_tree()
        self._refresh_warnings()
        self._refresh_status()
        self._on_tree_selection()

    def _refresh_tree(self) -> None:
        self.tree.clear()
        project = self.project
        if project is None or project.scenario_file is None:
            return
        naming = project.naming()
        t = theme.tokens()
        for module in project.scenario_file.modules:
            module_shots = sum(
                project.screenshot_count(case.uid) for case in module.test_cases
            )
            module_item = QTreeWidgetItem([
                module.name,
                str(module_shots),
                f"{len(module.test_cases)} test case, {module.step_count} step"
                f"  |  judul: {naming[module.name].title}",
            ])
            module_item.setData(0, NODE_ROLE, "module")
            module_item.setData(0, UID_ROLE, module.name)
            module_item.setIcon(0, icon("module"))
            module_font = QFont(module_item.font(0))
            module_font.setBold(True)
            module_item.setFont(0, module_font)
            module_item.setToolTip(0, module_item.text(2))
            self._style_count(module_item, module_shots)
            self.tree.addTopLevelItem(module_item)

            for group in module.scenarios:
                group_item = QTreeWidgetItem([
                    group.title, "", f"{len(group.test_cases)} test case",
                ])
                group_item.setData(0, NODE_ROLE, "scenario")
                group_item.setIcon(0, icon("scenario"))
                group_item.setForeground(0, QColor(t["text2"]))
                module_item.addChild(group_item)

                for case in group.test_cases:
                    count = project.screenshot_count(case.uid)
                    review = project.needs_review.get(case.uid)
                    case_item = QTreeWidgetItem([
                        case.display_label,
                        str(count) if count else "-",
                        ("\u26a0 perlu dicek ulang  |  " if review else "")
                        + f"{len(case.steps)} step",
                    ])
                    if review:
                        tip = "Skenario berubah sejak bukti diambil:\n" + "\n".join(review)
                        for column in range(3):
                            case_item.setToolTip(column, tip)
                    case_item.setData(0, NODE_ROLE, "case")
                    case_item.setData(0, UID_ROLE, case.uid)
                    if review:
                        case_item.setIcon(0, icon("review", t["warning"]))
                    elif count:
                        case_item.setIcon(0, icon("done", t["success"]))
                    else:
                        case_item.setIcon(0, icon("empty", t["text3"]))
                    if not review:
                        for column in range(3):
                            case_item.setToolTip(column, case.display_label)
                    self._style_count(case_item, count)
                    group_item.addChild(case_item)

                    for index, step in enumerate(case.steps, start=1):
                        text = step.test_step.replace("\n", " ")
                        label = step.test_step_id or str(index)
                        step_item = QTreeWidgetItem([
                            f"{label}. {text}", "", step.expected_results.replace("\n", " "),
                        ])
                        step_item.setData(0, NODE_ROLE, "step")
                        step_item.setForeground(0, QColor(t["text2"]))
                        step_item.setToolTip(0, f"{label}. {text}")
                        case_item.addChild(step_item)
        self.tree.expandToDepth(1)
        self._apply_filter(self.search.text())

    @staticmethod
    def _style_count(item: QTreeWidgetItem, count: int) -> None:
        """Right-aligned, muted screenshot count in column 1."""
        item.setTextAlignment(
            1, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        item.setForeground(
            1, QColor(theme.tokens()["text2" if count else "text3"])
        )

    def _apply_filter(self, text: str) -> None:
        needle = (text or "").strip().lower()
        for i in range(self.tree.topLevelItemCount()):
            module_item = self.tree.topLevelItem(i)
            module_visible = False
            for j in range(module_item.childCount()):
                group_item = module_item.child(j)
                group_visible = False
                for k in range(group_item.childCount()):
                    case_item = group_item.child(k)
                    match = not needle or _item_matches(case_item, needle)
                    case_item.setHidden(not match)
                    group_visible = group_visible or match
                group_item.setHidden(not group_visible)
                module_visible = module_visible or group_visible
            module_item.setHidden(not module_visible)

    def _current_uid(self) -> Optional[str]:
        item = self.tree.currentItem()
        while item is not None:
            if item.data(0, NODE_ROLE) == "case":
                return item.data(0, UID_ROLE)
            item = item.parent()
        return None

    def _on_tree_selection(self) -> None:
        self._refresh_detail()
        self._refresh_shots()
        self.results.set_test_case(self.project, self._current_uid())

    def _refresh_detail(self) -> None:
        item = self.tree.currentItem()
        project = self.project
        if project is None or item is None:
            self.detail.setHtml("")
            return
        t = theme.tokens()
        esc = html.escape

        def cell(value) -> str:
            return esc(str(value)).replace("\n", "<br>")

        def meta(rows) -> str:
            body = "".join(
                f'<tr><td style="color:{t["text2"]};padding:2px 16px 2px 0">{esc(k)}</td>'
                f"<td>{cell(v)}</td></tr>"
                for k, v in rows
            )
            return f'<table cellspacing="0">{body}</table>'

        def title(text: str) -> str:
            return (f'<div style="font-size:15pt;font-weight:600;'
                    f'margin-bottom:8px">{esc(text)}</div>')

        def heading(text: str) -> str:
            return (f'<div style="color:{t["text2"]};font-size:9pt;font-weight:600;'
                    f'margin:16px 0 4px 0">{esc(text.upper())}</div>')

        def badge(value: str) -> str:
            lowered = value.strip().lower()
            if lowered in ("p", "pass"):
                color = t["success"]
            elif lowered in ("f", "fail"):
                color = t["danger"]
            else:
                return cell(value)
            return f'<b style="color:{color}">{cell(value)}</b>'

        parts: List[str] = []
        kind = item.data(0, NODE_ROLE)
        if kind == "module":
            module = project.scenario_file.module(item.data(0, UID_ROLE))
            if module is not None:
                naming = project.naming()[module.name]
                parts.append(title(module.name))
                parts.append(meta([
                    ("Heading 1", f"{naming.title}   (sumber: {naming.source})"),
                    ("Nama file", f"Result Test UAT {naming.file_label}.docx"),
                    ("Baris header", module.header_row),
                    ("Test case", len(module.test_cases)),
                    ("Test step", module.step_count),
                ]))
                if module.metadata:
                    parts.append(heading("Metadata"))
                    parts.append(meta(list(module.metadata.items())))
                if module.unmapped_columns:
                    parts.append(heading("Kolom tidak dikenali"))
                    parts.append(esc(", ".join(text for _c, text in module.unmapped_columns)))
        else:
            uid = self._current_uid()
            case = project.test_case(uid) if uid else None
            if case is not None:
                parts.append(title(case.display_label))
                parts.append(meta([
                    ("UID", case.uid),
                    ("Heading 2", repr(case.heading)),
                    ("Scenario", f"{case.scenario_id} {case.test_scenario}".strip()),
                    ("Screenshot", project.screenshot_count(case.uid)),
                    ("Folder", f"screenshots/.../{case.folder_name}"),
                ]))
                for index, step in enumerate(case.steps, start=1):
                    label = step.test_step_id or str(index)
                    parts.append(
                        f'<div style="margin-top:14px"><b>{esc(label)}</b>'
                        f"&nbsp;&nbsp;{cell(step.test_step)}</div>"
                    )
                    rows = []
                    for field, value in (
                        ("Pre-Condition", step.pre_condition),
                        ("Test Data", step.test_data),
                        ("Expected", step.expected_results),
                        ("Actual", step.actual_results),
                        ("Pass/Fail", step.pass_fail),
                        ("Date", step.date_executed),
                        ("By", step.test_executed_by),
                        ("Notes", step.notes),
                    ):
                        if value:
                            rows.append((field, value))
                    if rows:
                        body = "".join(
                            f'<tr><td style="color:{t["text2"]};padding:1px 16px 1px 0">'
                            f"{esc(k)}</td><td>"
                            f"{badge(v) if k == 'Pass/Fail' else cell(v)}</td></tr>"
                            for k, v in rows
                        )
                        parts.append(f'<table cellspacing="0">{body}</table>')
        self.detail.setHtml("".join(parts))

    def _refresh_warnings(self) -> None:
        project = self.project
        if project is None:
            self.warnings.setPlainText("")
            return
        lines = list(project.warnings)
        missing = project.missing_files()
        for uid, entry in missing:
            lines.append(f"File screenshot hilang: {entry.file} (test case {uid})")
        if project.orphan_uids:
            lines.append(
                f"{project.orphan_count()} screenshot orphan dari "
                f"{len(project.orphan_uids)} test case yang sudah tidak ada di Excel."
            )
        orphan_results = project.orphan_results()
        if orphan_results:
            lines.append(
                f"Hasil eksekusi tersimpan untuk {len(orphan_results)} test case "
                "yang sudah tidak ada di Excel: " + ", ".join(orphan_results[:5])
                + ("..." if len(orphan_results) > 5 else "")
            )
        if project.needs_review:
            lines.append(
                f"{len(project.needs_review)} test case berubah sejak bukti diambil "
                "(ditandai di tree); cek lewat Project > Review perubahan skenario."
            )
        if project.orphan_step_results:
            parked = sum(len(steps) for steps in project.orphan_step_results.values())
            lines.append(
                f"{parked} hasil eksekusi milik step yang sudah tidak ada di Excel "
                "disimpan terpisah dan tidak ikut ditulis ke Excel."
            )
        self.warnings.setPlainText(
            "\n".join(f"- {line}" for line in lines) if lines else "Tidak ada peringatan."
        )
        index = self.tabs.indexOf(self.warnings)
        self.tabs.setTabText(index, f"Peringatan ({len(lines)})" if lines else "Peringatan")

    def _refresh_status(self) -> None:
        project = self.project
        if project is None:
            self.status.showMessage("Belum ada project dibuka.")
            self.progress_label.hide()
            self.progress_bar.hide()
            return
        cases = len(project.test_cases)
        with_shots = sum(1 for uid in project.known_uids if project.screenshot_count(uid))
        self.progress_label.setText(f"{with_shots}/{cases}")
        self.progress_bar.setRange(0, max(cases, 1))
        self.progress_bar.setValue(with_shots)
        self.progress_label.show()
        self.progress_bar.show()
        hotkey = (
            f"  |  hotkey {project.hotkey}" if self._listener is not None
            else "  |  hotkey tidak aktif"
        )
        self.status.showMessage(
            f"{project.name}  |  {len(project.selected_sheets)} module, {cases} test case, "
            f"{project.total_screenshots()} screenshot ({with_shots}/{cases} test case terisi)"
            + (f"  |  {project.orphan_count()} orphan" if project.orphan_count() else "")
            + hotkey
        )

    # -- screenshots ------------------------------------------------------- #

    def _refresh_shots(self) -> None:
        self.shot_list.blockSignals(True)
        self.shot_list.clear()
        self.shot_list.blockSignals(False)
        project = self.project
        uid = self._current_uid()
        if project is None or uid is None:
            self.shot_label.setText("Pilih sebuah test case di tree.")
            self._update_shot_buttons()
            return
        case = project.test_case(uid)
        self.shot_label.setText(
            f"<b>{case.display_label if case else uid}</b><br>"
            "Seret thumbnail untuk mengubah urutan; urutan ini dipakai di dokumen Word."
        )
        for entry in project.entries(uid):
            path = project.absolute(entry)
            item = QListWidgetItem(entry.caption or path.name)
            item.setData(FILE_ROLE, entry.file)
            if path.is_file():
                pixmap = QPixmap(str(path))
                if not pixmap.isNull():
                    item.setIcon(QIcon(pixmap.scaled(
                        THUMB,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )))
                item.setToolTip(str(path))
            else:
                item.setText(f"[file hilang] {path.name}")
                item.setToolTip(f"File tidak ditemukan: {path}")
            self.shot_list.addItem(item)
        self._update_shot_buttons()

    def _update_shot_buttons(self) -> None:
        uid = self._current_uid() if self.project else None
        has_case = uid is not None
        has_shot = has_case and self.shot_list.count() > 0
        self.btn_add_shot.setEnabled(bool(has_case))
        for button in (self.btn_caption, self.btn_move_shot, self.btn_delete_shot):
            button.setEnabled(bool(has_shot))

    def _persist_shot_order(self) -> None:
        """Write the order the user dragged the thumbnails into back to the project."""
        project = self.project
        uid = self._current_uid()
        if project is None or uid is None:
            return
        order = [
            self.shot_list.item(row).data(FILE_ROLE)
            for row in range(self.shot_list.count())
        ]
        entries = project.screenshots.get(uid, [])
        by_file = {entry.file: entry for entry in entries}
        reordered = [by_file[name] for name in order if name in by_file]
        if len(reordered) == len(entries):
            project.screenshots[uid] = reordered
            project.save()
            self.status.showMessage("Urutan screenshot disimpan.", 3000)

    def _selected_shot_index(self) -> int:
        return self.shot_list.currentRow()

    def _add_screenshot_from_file(self) -> None:
        project = self.project
        uid = self._current_uid()
        if project is None or uid is None:
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Pilih gambar", str(self._last_dir()),
            "Gambar (*.png *.jpg *.jpeg *.bmp)",
        )
        if not paths:
            return
        self._remember_dir(Path(paths[0]))
        try:
            for path in paths:
                project.add_screenshot(uid, Path(path), save=False)
            project.save()
        except ResultGeneratorError as exc:
            self._error("Gagal menambahkan screenshot", exc)
            return
        self._refresh_shots()
        self._refresh_tree()
        self._refresh_status()

    def _edit_caption(self) -> None:
        project = self.project
        uid = self._current_uid()
        index = self._selected_shot_index()
        if project is None or uid is None or index < 0:
            return
        entry = project.entries(uid)[index]
        text, ok = QInputDialog.getText(
            self, "Caption screenshot", "Caption (boleh dikosongkan):", text=entry.caption
        )
        if not ok:
            return
        project.set_caption(uid, index, text.strip())
        self._refresh_shots()

    def _move_screenshot(self) -> None:
        project = self.project
        uid = self._current_uid()
        index = self._selected_shot_index()
        if project is None or uid is None or index < 0:
            return
        choices = [
            (f"{case.module} - {case.display_label}", case.uid)
            for case in project.test_cases
            if case.uid != uid
        ]
        if not choices:
            return
        label, ok = QInputDialog.getItem(
            self, "Pindah ke test case lain", "Test case tujuan:",
            [text for text, _uid in choices], 0, False,
        )
        if not ok:
            return
        target = dict((text, value) for text, value in choices)[label]
        try:
            project.move_screenshot(uid, index, target)
        except ResultGeneratorError as exc:
            self._error("Gagal memindahkan screenshot", exc)
            return
        self._refresh_shots()
        self._refresh_tree()
        self._refresh_status()

    def _delete_screenshot(self) -> None:
        project = self.project
        uid = self._current_uid()
        index = self._selected_shot_index()
        if project is None or uid is None or index < 0:
            return
        answer = QMessageBox.question(
            self, "Hapus screenshot",
            "Hapus screenshot ini beserta file-nya?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        project.delete_screenshot(uid, index)
        self._refresh_shots()
        self._refresh_tree()
        self._refresh_status()

    def _preview_shot(self, item: QListWidgetItem) -> None:
        if self.project is None:
            return
        path = self.project.folder / item.data(FILE_ROLE)
        if path.is_file():
            _open_in_explorer(path)

    # -- helpers ----------------------------------------------------------- #

    def _error(self, title: str, exc: Exception, *, hint: str = "") -> None:
        message = getattr(exc, "message", str(exc))
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle(title)
        box.setText(message)
        if hint:
            box.setInformativeText(hint)
        detail = getattr(exc, "detail", None)
        if detail:
            box.setDetailedText(str(detail))
        box.exec()


def _tool_button(tooltip: str, slot) -> QToolButton:
    """Icon-only flat button; the icon is set later by ``apply_icons``."""
    button = QToolButton()
    button.setToolTip(tooltip)
    button.setAutoRaise(True)
    button.setIconSize(QSize(18, 18))
    button.clicked.connect(lambda: slot())
    return button


def _item_matches(case_item: QTreeWidgetItem, needle: str) -> bool:
    if needle in case_item.text(0).lower():
        return True
    for index in range(case_item.childCount()):
        if needle in case_item.child(index).text(0).lower():
            return True
    return False


class _busy:
    """Context manager showing the wait cursor during a slow operation."""

    def __enter__(self):
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        return self

    def __exit__(self, *exc_info):
        QApplication.restoreOverrideCursor()
        return False
