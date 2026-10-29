"""Offscreen tests for the capture UI and its wiring into the main window."""

from __future__ import annotations

import pytest

pytest.importorskip("PySide6")
pytest.importorskip("PIL")

from PIL import Image                                             # noqa: E402
from PySide6.QtCore import QEvent, QPointF, QRect, Qt             # noqa: E402
from PySide6.QtGui import QMouseEvent, QPixmap, QPointingDevice   # noqa: E402
from PySide6.QtWidgets import QApplication, QDialog               # noqa: E402

from rtg.capture import MODE_FULLSCREEN, MODE_MONITOR, MODE_REGION, Monitor  # noqa: E402
from rtg.config import load_config                                # noqa: E402
from rtg.gui import main_window as mw                             # noqa: E402
from rtg.gui.capture_dialogs import (                             # noqa: E402
    UID_ROLE,
    CaptureSettingsDialog,
    CaptureTargetDialog,
    RegionSelector,
    pil_to_qpixmap,
)
from rtg.gui.main_window import HotkeyBridge, MainWindow          # noqa: E402
from rtg.project import Project                                   # noqa: E402
from tests.fixtures import build_workbook, simple_workbook, write_scenario_sheet  # noqa: E402

CONFIG = load_config()

ROWS = [
    {"scenario_id": "SC-01", "test_scenario": "Alpha", "test_case_id": "1.1",
     "test_case": "Case Satu", "test_step": "buka aplikasi"},
    {"test_case_id": "1.2", "test_case": "Case Dua", "test_step": "cek laporan"},
]


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def project(tmp_path):
    excel = simple_workbook(tmp_path / "Scenario.xlsx", rows=ROWS)
    return Project.create(tmp_path / "proj", excel, config=CONFIG)


@pytest.fixture
def two_module_project(tmp_path):
    def builder(wb):
        write_scenario_sheet(wb.create_sheet("Module A"), rows=ROWS)
        write_scenario_sheet(wb.create_sheet("Module B"), rows=ROWS)

    excel = build_workbook(tmp_path / "two.xlsx", builder)
    return Project.create(tmp_path / "p", excel, config=CONFIG)


@pytest.fixture
def picture():
    return Image.new("RGB", (320, 180), (20, 60, 120))


# --------------------------------------------------------------------------- #
# Image conversion
# --------------------------------------------------------------------------- #

def test_pil_to_qpixmap_keeps_size_and_colour(qapp, picture):
    pixmap = pil_to_qpixmap(picture)
    assert (pixmap.width(), pixmap.height()) == (320, 180)
    colour = pixmap.toImage().pixelColor(10, 10)
    assert (colour.red(), colour.green(), colour.blue()) == (20, 60, 120)


# --------------------------------------------------------------------------- #
# Destination popup
# --------------------------------------------------------------------------- #

def test_target_dialog_lists_id_and_name(qapp, project, picture):
    dialog = CaptureTargetDialog(project, pil_to_qpixmap(picture))
    try:
        texts = [dialog.list.item(i).text() for i in range(dialog.list.count())]
        assert texts == ["[Scenario] 1.1 - Case Satu", "[Scenario] 1.2 - Case Dua"]
        assert dialog.selected_uid() == "Scenario::1.1"      # first by default
        assert "320 x 180" in dialog.size_label.text()
    finally:
        dialog.close()


def test_target_dialog_defaults_to_the_last_used_test_case(qapp, project, picture):
    dialog = CaptureTargetDialog(
        project, pil_to_qpixmap(picture), last_uid="Scenario::1.2"
    )
    try:
        assert dialog.selected_uid() == "Scenario::1.2"
    finally:
        dialog.close()


def test_target_dialog_shows_existing_screenshot_counts(qapp, project, picture, tmp_path):
    from tests.fixtures import make_png

    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    dialog = CaptureTargetDialog(project, pil_to_qpixmap(picture))
    try:
        assert "(1 screenshot)" in dialog.list.item(0).text()
        assert "screenshot)" not in dialog.list.item(1).text()
    finally:
        dialog.close()


def test_target_dialog_module_filter(qapp, two_module_project, picture):
    dialog = CaptureTargetDialog(two_module_project, pil_to_qpixmap(picture))
    try:
        assert dialog.list.count() == 4                      # all modules
        index = dialog.module_combo.findData("Module B")
        dialog.module_combo.setCurrentIndex(index)
        uids = [dialog.list.item(i).data(UID_ROLE) for i in range(dialog.list.count())]
        assert uids == ["Module B::1.1", "Module B::1.2"]
        # Filtered to one module, the prefix is dropped.
        assert dialog.list.item(0).text() == "1.1 - Case Satu"
    finally:
        dialog.close()


def test_target_dialog_search_spans_every_module(qapp, two_module_project, picture):
    dialog = CaptureTargetDialog(two_module_project, pil_to_qpixmap(picture))
    try:
        # Narrow to one module first: the search must still look everywhere.
        dialog.module_combo.setCurrentIndex(dialog.module_combo.findData("Module A"))
        assert dialog.list.count() == 2

        dialog.search.setText("case dua")
        uids = [dialog.list.item(i).data(UID_ROLE) for i in range(dialog.list.count())]
        assert uids == ["Module A::1.2", "Module B::1.2"]
        assert dialog.list.item(0).text().startswith("[Module A]")

        dialog.search.setText("module b")                    # module name matches too
        uids = [dialog.list.item(i).data(UID_ROLE) for i in range(dialog.list.count())]
        assert uids == ["Module B::1.1", "Module B::1.2"]
    finally:
        dialog.close()


def test_next_test_case_steps_forward_and_wraps(qapp, project, picture):
    dialog = CaptureTargetDialog(project, pil_to_qpixmap(picture))
    try:
        assert dialog.selected_uid() == "Scenario::1.1"
        dialog.select_next()
        assert dialog.selected_uid() == "Scenario::1.2"
        dialog.select_next()
        assert dialog.selected_uid() == "Scenario::1.1"      # wraps around
    finally:
        dialog.close()


def test_next_test_case_crosses_module_boundaries(qapp, two_module_project, picture):
    dialog = CaptureTargetDialog(
        two_module_project, pil_to_qpixmap(picture), last_uid="Module A::1.2"
    )
    try:
        dialog.select_next()
        assert dialog.selected_uid() == "Module B::1.1"
    finally:
        dialog.close()


def test_next_test_case_clears_a_filter_that_would_hide_it(qapp, two_module_project, picture):
    dialog = CaptureTargetDialog(
        two_module_project, pil_to_qpixmap(picture), last_uid="Module A::1.2"
    )
    try:
        dialog.module_combo.setCurrentIndex(dialog.module_combo.findData("Module A"))
        dialog.select_next()
        # The filter moved along instead of trapping the selection.
        assert dialog.selected_uid() == "Module B::1.1"
    finally:
        dialog.close()


def test_target_dialog_save_and_discard(qapp, project, picture):
    dialog = CaptureTargetDialog(project, pil_to_qpixmap(picture))
    try:
        dialog.caption.setText("  Bukti login  ")
        dialog._on_save()
        assert dialog.result() == QDialog.DialogCode.Accepted
        assert dialog.chosen_uid == "Scenario::1.1"
        assert dialog.chosen_caption == "Bukti login"
    finally:
        dialog.close()

    other = CaptureTargetDialog(project, pil_to_qpixmap(picture))
    try:
        other.reject()
        assert other.chosen_uid is None
    finally:
        other.close()


def test_target_dialog_works_without_a_preview(qapp, project):
    dialog = CaptureTargetDialog(project, None)
    try:
        assert dialog.list.count() == 2
    finally:
        dialog.close()


# --------------------------------------------------------------------------- #
# Region selector
# --------------------------------------------------------------------------- #

def _mouse(widget, kind, x, y):
    event = QMouseEvent(
        kind, QPointF(x, y), Qt.MouseButton.LeftButton,
        Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier,
        QPointingDevice.primaryPointingDevice(),
    )
    if kind == QEvent.Type.MouseButtonPress:
        widget.mousePressEvent(event)
    elif kind == QEvent.Type.MouseMove:
        widget.mouseMoveEvent(event)
    else:
        widget.mouseReleaseEvent(event)


def test_region_selector_drag_produces_a_box(qapp, picture):
    selector = RegionSelector(pil_to_qpixmap(picture), QRect(0, 0, 320, 180))
    try:
        assert selector.selected_box() is None
        _mouse(selector, QEvent.Type.MouseButtonPress, 20, 30)
        _mouse(selector, QEvent.Type.MouseMove, 140, 110)
        _mouse(selector, QEvent.Type.MouseButtonRelease, 140, 110)
        assert selector.selected_box() == (20, 30, 120, 80)
        assert selector.result() == QDialog.DialogCode.Accepted
    finally:
        selector.close()


def test_region_selector_drag_backwards_is_normalised(qapp, picture):
    selector = RegionSelector(pil_to_qpixmap(picture), QRect(0, 0, 320, 180))
    try:
        _mouse(selector, QEvent.Type.MouseButtonPress, 140, 110)
        _mouse(selector, QEvent.Type.MouseButtonRelease, 20, 30)
        assert selector.selected_box() == (20, 30, 120, 80)
    finally:
        selector.close()


def test_region_selector_ignores_an_accidental_click(qapp, picture):
    selector = RegionSelector(pil_to_qpixmap(picture), QRect(0, 0, 320, 180))
    try:
        _mouse(selector, QEvent.Type.MouseButtonPress, 20, 30)
        _mouse(selector, QEvent.Type.MouseButtonRelease, 21, 31)
        assert selector.selected_box() is None               # still selecting
        assert selector.result() != QDialog.DialogCode.Accepted
    finally:
        selector.close()


# --------------------------------------------------------------------------- #
# Capture settings
# --------------------------------------------------------------------------- #

MONITORS = [
    Monitor(0, 0, 0, 3840, 1080),
    Monitor(1, 0, 0, 1920, 1080),
    Monitor(2, 1920, 0, 1920, 1080),
]


def test_capture_settings_defaults(qapp, project):
    dialog = CaptureSettingsDialog(project, MONITORS)
    try:
        assert dialog.hotkey_edit.text() == "Ctrl+Shift+S"
        assert dialog.selected_mode() == MODE_FULLSCREEN
        assert dialog.monitor_combo.count() == 2             # monitor 0 is excluded
        assert not dialog._monitor_widget.isVisible()
    finally:
        dialog.close()


def test_capture_settings_rejects_an_invalid_hotkey(qapp, project):
    from PySide6.QtWidgets import QDialogButtonBox

    dialog = CaptureSettingsDialog(project, MONITORS)
    ok = dialog.buttons.button(QDialogButtonBox.StandardButton.Ok)
    try:
        dialog.hotkey_edit.setText("Ctrl")
        assert not ok.isEnabled()
        assert "satu tombol selain modifier" in dialog.hotkey_status.text()

        dialog.hotkey_edit.setText("S")
        assert not ok.isEnabled()

        dialog.hotkey_edit.setText("alt+f9")
        assert ok.isEnabled()
        assert "Alt+F9" in dialog.hotkey_status.text()
    finally:
        dialog.close()


def test_capture_settings_apply(qapp, project):
    dialog = CaptureSettingsDialog(project, MONITORS)
    try:
        dialog.hotkey_edit.setText("shift+ctrl+f8")
        dialog._mode_buttons[MODE_MONITOR].setChecked(True)
        dialog.monitor_combo.setCurrentIndex(dialog.monitor_combo.findData(2))
        dialog.apply()

        assert project.hotkey == "Ctrl+Shift+F8"             # normalised
        assert project.capture_mode == MODE_MONITOR
        assert project.capture_monitor == 2
    finally:
        dialog.close()

    project.save()
    reopened = Project.open(project.folder, config=CONFIG)
    assert reopened.hotkey == "Ctrl+Shift+F8"
    assert reopened.capture_mode == MODE_MONITOR
    assert reopened.capture_monitor == 2


def test_unknown_capture_mode_in_project_json_falls_back(qapp, project):
    import json

    data = json.loads(project.project_file.read_text(encoding="utf-8"))
    data["capture_mode"] = "telepati"
    project.project_file.write_text(json.dumps(data), encoding="utf-8")

    assert Project.open(project.folder, config=CONFIG).capture_mode == MODE_FULLSCREEN


def test_capture_settings_without_any_monitor_info(qapp, project):
    dialog = CaptureSettingsDialog(project, [])
    try:
        assert dialog.monitor_combo.count() == 1             # a usable fallback
    finally:
        dialog.close()


# --------------------------------------------------------------------------- #
# Wiring into the main window
# --------------------------------------------------------------------------- #

def test_hotkey_bridge_emits_on_the_qt_side(qapp):
    bridge = HotkeyBridge()
    seen = []
    bridge.triggered.connect(lambda: seen.append(1))
    bridge.fire()
    qapp.processEvents()
    assert seen == [1]


@pytest.fixture
def window(qapp, project, monkeypatch):
    # Never register a real OS hotkey during the tests.
    monkeypatch.setattr(MainWindow, "_start_hotkey", lambda self: None)
    win = MainWindow(project)
    yield win
    win.close()


def test_capture_now_saves_to_the_chosen_test_case(window, project, picture, monkeypatch):
    monkeypatch.setattr(
        MainWindow, "_grab", lambda self: (picture, pil_to_qpixmap(picture))
    )

    class FakeDialog:
        def __init__(self, *args, **kwargs):
            self.chosen_uid = "Scenario::1.2"
            self.chosen_caption = "Hasil cek laporan"

        def exec(self):
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(mw, "CaptureTargetDialog", FakeDialog)

    window.capture_now()

    assert project.screenshot_count("Scenario::1.2") == 1
    entry = project.entries("Scenario::1.2")[0]
    assert entry.caption == "Hasil cek laporan"
    assert project.absolute(entry).is_file()
    assert entry.file.startswith("screenshots/scenario/1.2_case-dua/")
    assert window._last_uid == "Scenario::1.2"
    # The tree jumped to that test case and opened the Screenshot tab.
    assert window._current_uid() == "Scenario::1.2"
    assert window.tabs.currentIndex() == 1
    assert window.shot_list.count() == 1


def test_capture_now_discard_keeps_nothing(window, project, picture, monkeypatch):
    monkeypatch.setattr(
        MainWindow, "_grab", lambda self: (picture, pil_to_qpixmap(picture))
    )

    class FakeDialog:
        def __init__(self, *args, **kwargs):
            self.chosen_uid = None
            self.chosen_caption = ""

        def exec(self):
            return QDialog.DialogCode.Rejected

    monkeypatch.setattr(mw, "CaptureTargetDialog", FakeDialog)

    window.capture_now()
    assert project.total_screenshots() == 0
    assert "dibuang" in window.statusBar().currentMessage()


def test_cancelled_region_selection_saves_nothing(window, project, monkeypatch):
    monkeypatch.setattr(MainWindow, "_grab", lambda self: (None, None))
    window.capture_now()
    assert project.total_screenshots() == 0


def test_capture_error_is_reported_not_raised(window, project, monkeypatch):
    from rtg.capture import CaptureError

    def boom(self):
        raise CaptureError("Gagal mengambil screenshot: driver bermasalah")

    monkeypatch.setattr(MainWindow, "_grab", boom)
    shown = []
    monkeypatch.setattr(
        MainWindow, "_error",
        lambda self, title, exc, **kw: shown.append((title, exc.message)),
    )

    window.capture_now()
    assert shown and "driver bermasalah" in shown[0][1]
    assert project.total_screenshots() == 0


def test_capture_is_not_reentrant(window, project, picture, monkeypatch):
    monkeypatch.setattr(
        MainWindow, "_grab", lambda self: (picture, pil_to_qpixmap(picture))
    )
    calls = []

    class FakeDialog:
        def __init__(self, *args, **kwargs):
            calls.append(1)
            self.chosen_uid = "Scenario::1.1"
            self.chosen_caption = ""

        def exec(self):
            # A second hotkey press while the popup is open must be ignored.
            window.capture_now()
            return QDialog.DialogCode.Accepted

    monkeypatch.setattr(mw, "CaptureTargetDialog", FakeDialog)

    window.capture_now()
    assert calls == [1]
    assert project.screenshot_count("Scenario::1.1") == 1


def test_capture_without_test_cases_explains_itself(qapp, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QMessageBox

    excel = simple_workbook(tmp_path / "s.xlsx", rows=ROWS)
    project = Project.create(tmp_path / "p2", excel, config=CONFIG)
    project.scenario_file.modules.clear()                    # simulate "no module chosen"

    monkeypatch.setattr(MainWindow, "_start_hotkey", lambda self: None)
    shown = []
    monkeypatch.setattr(
        QMessageBox, "information", staticmethod(lambda *a, **k: shown.append(a[2]))
    )
    win = MainWindow(project)
    try:
        win.capture_now()
        assert shown and "belum punya test case" in shown[0]
    finally:
        win.close()


# --------------------------------------------------------------------------- #
# The chosen hotkey must survive into the next project
# --------------------------------------------------------------------------- #

def test_capture_settings_are_remembered_for_new_projects(window, project):
    """Regression: a new project used to reset the hotkey to Ctrl+Shift+S."""
    dialog = CaptureSettingsDialog(project, MONITORS, parent=window)
    try:
        dialog.hotkey_edit.setText("Ctrl+Shift+F")
        dialog._mode_buttons[MODE_REGION].setChecked(True)
        dialog.apply()
    finally:
        dialog.close()
    window.project.save()
    window._remember_capture(window.project)

    preferred = window._preferred_capture()
    assert preferred["hotkey"] == "Ctrl+Shift+F"
    assert preferred["capture_mode"] == MODE_REGION

    # A project created with those preferences keeps them.
    second = Project.create(
        project.folder.parent / "proj2", project.excel_path,
        config=CONFIG, **preferred,
    )
    assert second.hotkey == "Ctrl+Shift+F"
    assert second.capture_mode == MODE_REGION


def test_opening_an_old_project_does_not_change_the_preference(window, project, tmp_path):
    window._remember_capture_called = False
    window.settings.setValue("hotkey", "Ctrl+Shift+F")

    older = Project.create(
        tmp_path / "older", project.excel_path, config=CONFIG, hotkey="Alt+F9"
    )
    window._set_project(older)

    # The project keeps its own hotkey...
    assert window.project.hotkey == "Alt+F9"
    # ...but the preference for new projects is untouched.
    assert window._preferred_capture()["hotkey"] == "Ctrl+Shift+F"
