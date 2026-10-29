"""Capture-related UI: region selector, destination popup, capture settings."""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from ..capture import CAPTURE_MODES, MODE_FULLSCREEN, MODE_MONITOR, MODE_REGION, MODE_LABELS
from ..capture import Monitor
from ..hotkey import HotkeyError, format_hotkey, parse_hotkey
from ..models import TestCase
from ..project import Project
from . import theme

UID_ROLE = Qt.ItemDataRole.UserRole


def pil_to_qpixmap(image) -> QPixmap:
    """Convert a Pillow image to a QPixmap without going through a file."""
    rgb = image.convert("RGB")
    data = rgb.tobytes("raw", "RGB")
    qimage = QImage(data, rgb.width, rgb.height, 3 * rgb.width,
                    QImage.Format.Format_RGB888).copy()
    return QPixmap.fromImage(qimage)


# --------------------------------------------------------------------------- #
# Region selector
# --------------------------------------------------------------------------- #

class RegionSelector(QDialog):
    """Snipping-Tool style overlay for picking a rectangle.

    The screen is captured *before* this overlay appears and shown as its
    background, so the overlay itself can never end up in the screenshot and
    the selection is made against a frozen image.
    """

    def __init__(self, pixmap: QPixmap, geometry: QRect,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setCursor(Qt.CursorShape.CrossCursor)
        self.setMouseTracking(True)
        self._pixmap = pixmap
        self._origin: Optional[QPoint] = None
        self._current: Optional[QPoint] = None
        self._rect = QRect()
        self.setGeometry(geometry)

    # -- painting ---------------------------------------------------------- #

    def paintEvent(self, event) -> None:  # noqa: N802 - Qt naming
        painter = QPainter(self)
        target = self.rect()
        painter.drawPixmap(target, self._pixmap)
        painter.fillRect(target, QColor(0, 0, 0, 110))

        selection = self._selection()
        if not selection.isNull() and selection.width() > 0 and selection.height() > 0:
            source = QRect(
                int(selection.x() * self._pixmap.width() / max(1, target.width())),
                int(selection.y() * self._pixmap.height() / max(1, target.height())),
                int(selection.width() * self._pixmap.width() / max(1, target.width())),
                int(selection.height() * self._pixmap.height() / max(1, target.height())),
            )
            painter.drawPixmap(selection, self._pixmap, source)
            pen = QPen(QColor(0, 140, 255), 2)
            painter.setPen(pen)
            painter.drawRect(selection.adjusted(0, 0, -1, -1))

            label = f"{selection.width()} x {selection.height()}"
            painter.setPen(QPen(QColor(255, 255, 255)))
            text_point = selection.topLeft() + QPoint(6, -8)
            if text_point.y() < 14:
                text_point = selection.topLeft() + QPoint(6, 18)
            painter.fillRect(
                QRect(text_point.x() - 4, text_point.y() - 14, 11 * len(label), 20),
                QColor(0, 0, 0, 160),
            )
            painter.drawText(text_point, label)
        else:
            painter.setPen(QPen(QColor(255, 255, 255)))
            painter.drawText(
                target.adjusted(0, 24, 0, 0),
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                "Seret untuk memilih area  -  Esc untuk batal",
            )
        painter.end()

    def _selection(self) -> QRect:
        """The dragged rectangle, measured the same way in every direction.

        ``QRect(point, point)`` treats both corners as inclusive and loses a
        pixel on each side once normalised, so dragging right-down and
        left-up would return different sizes for the same visible area.
        """
        if self._origin is None or self._current is None:
            return QRect()
        left, right = sorted((self._origin.x(), self._current.x()))
        top, bottom = sorted((self._origin.y(), self._current.y()))
        return QRect(left, top, right - left, bottom - top)

    # -- interaction ------------------------------------------------------- #

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self._origin = event.position().toPoint()
            self._current = self._origin
            self.update()

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._origin is not None:
            self._current = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() != Qt.MouseButton.LeftButton or self._origin is None:
            return
        self._current = event.position().toPoint()
        self._rect = self._selection()
        if self._rect.width() < 4 or self._rect.height() < 4:
            self._origin = self._current = None
            self._rect = QRect()
            self.update()
            return
        self.accept()

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key.Key_Escape:
            self._rect = QRect()
            self.reject()
            return
        super().keyPressEvent(event)

    def selected_box(self) -> Optional[Tuple[int, int, int, int]]:
        """``(left, top, width, height)`` in widget coordinates, or None."""
        if self._rect.isNull() or self._rect.width() <= 0:
            return None
        return (self._rect.x(), self._rect.y(), self._rect.width(), self._rect.height())


# --------------------------------------------------------------------------- #
# Destination popup
# --------------------------------------------------------------------------- #

class CaptureTargetDialog(QDialog):
    """Ask which test case the fresh screenshot belongs to.

    Defaults to the test case used last so a run of captures only needs Enter,
    with *Next test case* to step forward through the scenario.
    """

    def __init__(
        self,
        project: Project,
        pixmap: Optional[QPixmap] = None,
        *,
        last_uid: Optional[str] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Simpan screenshot ke test case")
        self.resize(760, 620)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self._project = project
        self._cases: List[TestCase] = list(project.test_cases)
        self._chosen_uid: Optional[str] = None

        layout = QVBoxLayout(self)

        if pixmap is not None and not pixmap.isNull():
            preview = QLabel()
            preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
            preview.setPixmap(pixmap.scaled(
                720, 260,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            ))
            preview.setStyleSheet(f"border: 1px solid {theme.tokens()['border']};")
            layout.addWidget(preview)
            self.size_label = QLabel(f"Ukuran: {pixmap.width()} x {pixmap.height()} px")
            layout.addWidget(self.size_label)

        row = QHBoxLayout()
        row.addWidget(QLabel("Module:"))
        self.module_combo = QComboBox()
        self.module_combo.addItem("(semua module)", None)
        for module in (project.scenario_file.modules if project.scenario_file else []):
            self.module_combo.addItem(module.name, module.name)
        self.module_combo.currentIndexChanged.connect(lambda _i: self._refill())
        row.addWidget(self.module_combo, 1)
        layout.addLayout(row)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Cari:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText(
            "Ketik Test Case ID atau nama test case - mencari di semua module"
        )
        self.search.textChanged.connect(lambda _t: self._refill())
        search_row.addWidget(self.search, 1)
        layout.addLayout(search_row)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list.itemDoubleClicked.connect(lambda _item: self._on_save())
        layout.addWidget(self.list, 1)

        caption_row = QHBoxLayout()
        caption_row.addWidget(QLabel("Caption (opsional):"))
        self.caption = QLineEdit()
        caption_row.addWidget(self.caption, 1)
        layout.addLayout(caption_row)

        buttons_row = QHBoxLayout()
        self.next_button = QPushButton("Next test case")
        self.next_button.setToolTip("Pilih test case berikutnya di urutan skenario")
        self.next_button.clicked.connect(self.select_next)
        buttons_row.addWidget(self.next_button)
        buttons_row.addStretch(1)
        self.discard_button = QPushButton("Buang")
        self.discard_button.clicked.connect(self.reject)
        buttons_row.addWidget(self.discard_button)
        self.save_button = QPushButton("Simpan")
        self.save_button.setDefault(True)
        self.save_button.clicked.connect(self._on_save)
        buttons_row.addWidget(self.save_button)
        layout.addLayout(buttons_row)

        self._refill(select_uid=last_uid or (self._cases[0].uid if self._cases else None))

    # -- list -------------------------------------------------------------- #

    def _visible_cases(self) -> List[TestCase]:
        needle = self.search.text().strip().lower()
        module = self.module_combo.currentData()
        cases = self._cases
        # A search always looks across every module, as asked for.
        if needle:
            return [
                case for case in cases
                if needle in case.test_case_id.lower()
                or needle in case.test_case.lower()
                or needle in case.module.lower()
            ]
        if module:
            return [case for case in cases if case.module == module]
        return list(cases)

    def _refill(self, *, select_uid: Optional[str] = None) -> None:
        keep = select_uid or self.selected_uid()
        self.list.clear()
        show_module = (
            bool(self.search.text().strip())
            or self.module_combo.currentData() is None
        )
        for case in self._visible_cases():
            count = self._project.screenshot_count(case.uid)
            text = f"{case.test_case_id} - {case.test_case}"
            if show_module:
                text = f"[{case.module}] {text}"
            if count:
                text += f"   ({count} screenshot)"
            item = QListWidgetItem(text)
            item.setData(UID_ROLE, case.uid)
            self.list.addItem(item)
        self._select(keep)
        self._update_buttons()

    def _select(self, uid: Optional[str]) -> None:
        if uid:
            for row in range(self.list.count()):
                if self.list.item(row).data(UID_ROLE) == uid:
                    self.list.setCurrentRow(row)
                    self.list.scrollToItem(self.list.item(row))
                    return
        if self.list.count():
            self.list.setCurrentRow(0)

    def _update_buttons(self) -> None:
        enabled = self.list.currentItem() is not None
        self.save_button.setEnabled(enabled)
        self.next_button.setEnabled(bool(self._cases))

    def selected_uid(self) -> Optional[str]:
        item = self.list.currentItem()
        return item.data(UID_ROLE) if item is not None else None

    def select_next(self) -> None:
        """Move to the next test case in scenario order, wrapping around."""
        current = self.selected_uid()
        order = [case.uid for case in self._cases]
        if current in order:
            nxt = order[(order.index(current) + 1) % len(order)]
        elif order:
            nxt = order[0]
        else:
            return
        # Clear the filters so the next test case is always reachable.
        case = self._project.test_case(nxt)
        if case is not None and not self._matches_filters(case):
            self.search.blockSignals(True)
            self.search.clear()
            self.search.blockSignals(False)
            index = self.module_combo.findData(case.module)
            self.module_combo.blockSignals(True)
            self.module_combo.setCurrentIndex(index if index >= 0 else 0)
            self.module_combo.blockSignals(False)
        self._refill(select_uid=nxt)

    def _matches_filters(self, case: TestCase) -> bool:
        return any(
            item.data(UID_ROLE) == case.uid
            for item in (self.list.item(row) for row in range(self.list.count()))
        )

    # -- result ------------------------------------------------------------ #

    def _on_save(self) -> None:
        uid = self.selected_uid()
        if not uid:
            return
        self._chosen_uid = uid
        self.accept()

    @property
    def chosen_uid(self) -> Optional[str]:
        return self._chosen_uid

    @property
    def chosen_caption(self) -> str:
        return self.caption.text().strip()


# --------------------------------------------------------------------------- #
# Capture settings
# --------------------------------------------------------------------------- #

class CaptureSettingsDialog(QDialog):
    """Change the global hotkey and the capture mode."""

    def __init__(
        self,
        project: Project,
        monitors: Sequence[Monitor],
        *,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Pengaturan screenshot")
        self.resize(560, 400)
        self._project = project
        self._monitors = list(monitors)

        layout = QVBoxLayout(self)

        hotkey_box = QGroupBox("Global hotkey")
        hotkey_layout = QVBoxLayout(hotkey_box)
        hotkey_layout.addWidget(QLabel(
            "Hotkey ini aktif walaupun aplikasi sedang diminimize.<br>"
            "Contoh penulisan: <code>Ctrl+Shift+S</code>, <code>Alt+F9</code>, "
            "<code>Ctrl+Alt+PrintScreen</code>."
        ))
        self.hotkey_edit = QLineEdit(project.hotkey)
        self.hotkey_edit.textChanged.connect(self._validate_hotkey)
        hotkey_layout.addWidget(self.hotkey_edit)
        self.hotkey_status = QLabel()
        self.hotkey_status.setWordWrap(True)
        hotkey_layout.addWidget(self.hotkey_status)
        layout.addWidget(hotkey_box)

        mode_box = QGroupBox("Mode capture")
        mode_layout = QVBoxLayout(mode_box)
        self._mode_buttons = {}
        for mode in CAPTURE_MODES:
            button = QRadioButton(MODE_LABELS[mode])
            button.setChecked(project.capture_mode == mode)
            button.toggled.connect(lambda _c: self._update_monitor_row())
            self._mode_buttons[mode] = button
            mode_layout.addWidget(button)

        monitor_row = QHBoxLayout()
        monitor_row.addWidget(QLabel("Monitor:"))
        self.monitor_combo = QComboBox()
        for monitor in self._monitors:
            if monitor.index == 0:
                continue
            self.monitor_combo.addItem(monitor.label, monitor.index)
        if self.monitor_combo.count() == 0:
            self.monitor_combo.addItem("Monitor 1", 1)
        index = self.monitor_combo.findData(project.capture_monitor)
        self.monitor_combo.setCurrentIndex(index if index >= 0 else 0)
        monitor_row.addWidget(self.monitor_combo, 1)
        self._monitor_widget = QWidget()
        self._monitor_widget.setLayout(monitor_row)
        mode_layout.addWidget(self._monitor_widget)
        layout.addWidget(mode_box)
        layout.addStretch(1)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self._validate_hotkey()
        self._update_monitor_row()

    def _validate_hotkey(self) -> None:
        text = self.hotkey_edit.text()
        try:
            normalised = format_hotkey(text)
        except HotkeyError as exc:
            self.hotkey_status.setText(f"<span style='color:#b00020'>{exc.message}</span>")
            self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)
            return
        self.hotkey_status.setText(f"Akan didaftarkan sebagai: <b>{normalised}</b>")
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(True)

    def _update_monitor_row(self) -> None:
        self._monitor_widget.setVisible(self._mode_buttons[MODE_MONITOR].isChecked())

    def selected_mode(self) -> str:
        for mode, button in self._mode_buttons.items():
            if button.isChecked():
                return mode
        return MODE_FULLSCREEN

    def apply(self) -> None:
        self._project.hotkey = format_hotkey(self.hotkey_edit.text())
        self._project.capture_mode = self.selected_mode()
        self._project.capture_monitor = int(self.monitor_combo.currentData() or 1)
