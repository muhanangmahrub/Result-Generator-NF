"""Smoke tests for the PySide6 front-end, run with the offscreen platform.

They do not try to cover the look of the UI, only that the window and dialogs
build, that the tree reflects the parsed scenarios and the screenshot counts,
and that the dialogs return what the rest of the application expects.
"""

from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt                                    # noqa: E402
from PySide6.QtWidgets import QApplication                       # noqa: E402

from rtg.config import load_config                               # noqa: E402
from rtg.excel_parser import inspect_workbook                    # noqa: E402
from rtg.gui.dialogs import (                                    # noqa: E402
    ColumnMappingDialog,
    ExportDialog,
    MetadataDialog,
    OrphanDialog,
    SheetPickerDialog,
    field_label,
)
from rtg.gui.main_window import MainWindow                       # noqa: E402
from rtg.project import Project                                  # noqa: E402
from tests.fixtures import (                                     # noqa: E402
    DEFAULT_METADATA,
    FULL_FIELD_ORDER,
    build_workbook,
    make_png,
    simple_workbook,
    write_noise_sheet,
    write_scenario_sheet,
)

CONFIG = load_config()

ROWS = [
    {"scenario_id": "SC-01", "test_scenario": "Alpha", "test_case_id": "1.1",
     "test_case": "Case Satu", "test_step": "buka aplikasi"},
    {"test_step": "login dengan akun uat"},
    {"test_case_id": "1.2", "test_case": "Case Dua", "test_step": "cek laporan"},
]


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def excel(tmp_path):
    return simple_workbook(tmp_path / "Scenario.xlsx", rows=ROWS)


@pytest.fixture
def project(tmp_path, excel):
    return Project.create(tmp_path / "proj", excel, config=CONFIG)


@pytest.fixture
def window(qapp, project):
    win = MainWindow(project)
    yield win
    win.close()


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #

def test_window_without_a_project_disables_project_actions(qapp):
    win = MainWindow()
    try:
        assert win.act_new.isEnabled()
        assert win.act_open.isEnabled()
        for action in (win.act_save, win.act_reimport, win.act_generate,
                       win.act_metadata, win.act_orphans):
            assert not action.isEnabled()
        assert win.tree.topLevelItemCount() == 0
    finally:
        win.close()


def test_tree_shows_module_scenario_case_step(window):
    tree = window.tree
    assert tree.topLevelItemCount() == 1

    module_item = tree.topLevelItem(0)
    assert module_item.text(0) == "Scenario"
    assert "2 test case, 3 step" in module_item.text(2)
    assert "SCENARIO" in module_item.text(2)          # resolved Heading 1

    group_item = module_item.child(0)
    assert group_item.text(0) == "SC-01  Alpha"
    assert group_item.childCount() == 2

    case_item = group_item.child(0)
    assert case_item.text(0) == "1.1  Case Satu"
    assert case_item.text(1) == "-"                   # no screenshot yet
    assert case_item.childCount() == 2
    assert case_item.child(0).text(0) == "1. buka aplikasi"


def test_tree_shows_screenshot_counts(window, project, tmp_path):
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "b.png"))
    window._refresh_all()

    module_item = window.tree.topLevelItem(0)
    case_item = module_item.child(0).child(0)
    assert case_item.text(1) == "2"
    assert module_item.text(1) == "2"


def test_status_bar_counts(window, project, tmp_path):
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    window._refresh_all()
    message = window.statusBar().currentMessage()
    assert "2 test case" in message
    assert "1 screenshot" in message
    assert "1/2 test case terisi" in message


def test_filter_hides_non_matching_test_cases(window):
    window.search.setText("case dua")
    group_item = window.tree.topLevelItem(0).child(0)
    assert group_item.child(0).isHidden() is True
    assert group_item.child(1).isHidden() is False

    # A step text matches its test case too.
    window.search.setText("login")
    assert group_item.child(0).isHidden() is False
    assert group_item.child(1).isHidden() is True

    window.search.setText("")
    assert group_item.child(0).isHidden() is False
    assert group_item.child(1).isHidden() is False


def test_selecting_a_test_case_fills_the_detail_tab(window):
    group_item = window.tree.topLevelItem(0).child(0)
    window.tree.setCurrentItem(group_item.child(0))

    text = window.detail.toPlainText()
    assert "Scenario::1.1" in text
    assert "buka aplikasi" in text
    assert "1.1_case-satu" in text


def test_selecting_a_module_shows_its_naming(window):
    window.tree.setCurrentItem(window.tree.topLevelItem(0))
    text = window.detail.toPlainText()
    assert "Heading 1" in text
    assert "Result Test UAT Scenario.docx" in text


def test_selecting_a_step_still_resolves_its_test_case(window, project, tmp_path):
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    window._refresh_all()
    case_item = window.tree.topLevelItem(0).child(0).child(0)
    window.tree.setCurrentItem(case_item.child(1))      # a step

    assert window._current_uid() == "Scenario::1.1"
    assert window.shot_list.count() == 1


def test_screenshot_panel_lists_thumbnails(window, project, tmp_path):
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"), caption="Bukti A")
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "b.png"))
    window._refresh_all()
    window.tree.setCurrentItem(window.tree.topLevelItem(0).child(0).child(0))

    assert window.shot_list.count() == 2
    assert window.shot_list.item(0).text() == "Bukti A"
    assert not window.shot_list.item(0).icon().isNull()
    assert window.btn_delete_shot.isEnabled()


def test_missing_screenshot_file_is_flagged_in_the_panel(window, project, tmp_path):
    entry = project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    project.absolute(entry).unlink()
    window._refresh_all()
    window.tree.setCurrentItem(window.tree.topLevelItem(0).child(0).child(0))

    assert "[file hilang]" in window.shot_list.item(0).text()
    assert "hilang" in window.warnings.toPlainText()


def test_dragging_thumbnails_persists_the_new_order(window, project, tmp_path):
    first = project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    second = project.add_screenshot("Scenario::1.1", make_png(tmp_path / "b.png"))
    window._refresh_all()
    window.tree.setCurrentItem(window.tree.topLevelItem(0).child(0).child(0))

    # Simulate the drop result: move the second thumbnail to the front.
    item = window.shot_list.takeItem(1)
    window.shot_list.insertItem(0, item)
    window._persist_shot_order()

    assert [e.file for e in project.entries("Scenario::1.1")] == [second.file, first.file]
    assert [e.file for e in Project.open(project.folder, config=CONFIG)
            .entries("Scenario::1.1")] == [second.file, first.file]


def test_warnings_tab_reports_parse_warnings(qapp, tmp_path):
    rows = ROWS + [{"test_scenario": "Baris nyasar tanpa step"}]
    excel = simple_workbook(tmp_path / "warn.xlsx", rows=rows)
    project = Project.create(tmp_path / "p", excel, config=CONFIG)
    win = MainWindow(project)
    try:
        assert "diabaikan" in win.warnings.toPlainText()
        assert win.tabs.tabText(win.tabs.indexOf(win.warnings)).startswith("Peringatan (")
    finally:
        win.close()


def test_orphan_count_reaches_the_status_bar_and_warnings(window, project, tmp_path, excel):
    project.add_screenshot("Scenario::1.2", make_png(tmp_path / "a.png"))
    simple_workbook(excel, rows=ROWS[:2])
    project.reimport()
    window._refresh_all()

    assert "1 orphan" in window.statusBar().currentMessage()
    assert "orphan" in window.warnings.toPlainText()


# --------------------------------------------------------------------------- #
# Sheet picker
# --------------------------------------------------------------------------- #

@pytest.fixture
def multi_excel(tmp_path):
    def builder(wb):
        write_noise_sheet(wb.create_sheet("Main"), "Main")
        write_scenario_sheet(wb.create_sheet("Module A"), rows=ROWS)
        write_scenario_sheet(wb.create_sheet("Module B"), rows=ROWS)
        write_scenario_sheet(
            wb.create_sheet("Broken"),
            fields=tuple(f for f in FULL_FIELD_ORDER if f != "test_case"),
            rows=ROWS,
        )

    return build_workbook(tmp_path / "multi.xlsx", builder)


def test_sheet_picker_lists_every_sheet_with_status(qapp, multi_excel):
    inspection = inspect_workbook(multi_excel, config=CONFIG)
    dialog = SheetPickerDialog(
        inspection, CONFIG,
        reinspect=lambda ov: inspect_workbook(multi_excel, config=CONFIG, overrides=ov),
    )
    try:
        assert dialog.table.rowCount() == 4
        statuses = {
            dialog.table.item(row, 0).text(): dialog.table.item(row, 1).text()
            for row in range(dialog.table.rowCount())
        }
        assert statuses["Module A"] == "sheet skenario"
        assert statuses["Main"] == "bukan skenario"
        assert statuses["Broken"] == "? perlu mapping"

        # Scenario sheets start checked, the others are not even checkable.
        assert dialog.selected_sheets() == ["Module A", "Module B"]
        main_row = list(statuses).index("Main")
        item = dialog.table.item(main_row, 0)
        assert not (item.flags() & Qt.ItemFlag.ItemIsUserCheckable)
    finally:
        dialog.close()


def test_sheet_picker_select_and_clear(qapp, multi_excel):
    inspection = inspect_workbook(multi_excel, config=CONFIG)
    dialog = SheetPickerDialog(
        inspection, CONFIG,
        reinspect=lambda ov: inspect_workbook(multi_excel, config=CONFIG, overrides=ov),
        selected=["Module B"],
    )
    try:
        assert dialog.selected_sheets() == ["Module B"]
        dialog._select_all()
        assert dialog.selected_sheets() == ["Module A", "Module B"]
        dialog._clear_all()
        assert dialog.selected_sheets() == []
    finally:
        dialog.close()


def test_sheet_picker_detail_text(qapp, multi_excel):
    inspection = inspect_workbook(multi_excel, config=CONFIG)
    dialog = SheetPickerDialog(
        inspection, CONFIG,
        reinspect=lambda ov: inspect_workbook(multi_excel, config=CONFIG, overrides=ov),
    )
    try:
        dialog.table.setCurrentCell(3, 0)               # the Broken sheet
        text = dialog.detail.text()
        assert "Kolom wajib tidak ditemukan" in text
        assert "Test Case ID=kol" in text
    finally:
        dialog.close()


# --------------------------------------------------------------------------- #
# Column mapping
# --------------------------------------------------------------------------- #

def test_column_mapping_dialog_builds_and_validates(qapp, multi_excel):
    inspection = inspect_workbook(multi_excel, config=CONFIG).sheet("Broken")
    dialog = ColumnMappingDialog(
        inspection, CONFIG, header_choices=inspection.headers
    )
    try:
        # Auto-detected columns are preselected, the missing one is empty.
        assert dialog._combos["test_case_id"].currentData() == \
            inspection.column_map["test_case_id"]
        assert dialog._combos["test_case"].currentData() is None

        dialog._on_accept()                             # must refuse
        assert dialog.result() == 0
        assert "Test Case" in dialog.warning.text()

        # Point Test Case at the Test Scenario column: now it is complete.
        combo = dialog._combos["test_case"]
        combo.setCurrentIndex(combo.findData(inspection.column_map["test_scenario"]))
        dialog._on_accept()
        assert "dipakai untuk lebih dari satu field" in dialog.warning.text()

        free_column = max(col for col, _text in inspection.headers)
        combo.setCurrentIndex(combo.findData(free_column))
        dialog._on_accept()

        override = dialog.result_override()
        assert override is not None
        assert override.header_row == inspection.header_row
        assert override.columns["test_case"] == free_column
    finally:
        dialog.close()


def test_column_mapping_reset_returns_none(qapp, multi_excel):
    inspection = inspect_workbook(multi_excel, config=CONFIG).sheet("Module A")
    dialog = ColumnMappingDialog(
        inspection, CONFIG, header_choices=inspection.headers
    )
    try:
        dialog._on_reset()
        assert dialog.result_override() is None
    finally:
        dialog.close()


def test_field_label_is_human_readable():
    assert field_label("test_case_id") == "Test Case ID"
    assert field_label("pass_fail") == "Pass/Fail"
    assert field_label("sesuatu_baru") == "Sesuatu Baru"


# --------------------------------------------------------------------------- #
# Metadata, export and orphan dialogs
# --------------------------------------------------------------------------- #

def test_metadata_dialog_round_trip(qapp, tmp_path):
    excel = simple_workbook(
        tmp_path / "meta.xlsx", header_row=9, metadata=DEFAULT_METADATA,
        metadata_start_row=2, rows=ROWS,
    )
    project = Project.create(tmp_path / "p", excel, config=CONFIG)
    dialog = MetadataDialog(project)
    try:
        # Parsed values show up as placeholders, not as values.
        assert dialog._meta_edits["project_name"].placeholderText() == "New Validation AMF"

        dialog._meta_edits["project_name"].setText("Diedit")
        dialog._title_edits["Scenario"].setText("JUDUL BARU")
        dialog.apply()

        assert project.metadata["project_name"] == "Diedit"
        assert project.module_titles["Scenario"] == "JUDUL BARU"

        # Clearing a field removes the override again.
        dialog._meta_edits["project_name"].setText("")
        dialog._title_edits["Scenario"].setText("")
        dialog.apply()
        assert "project_name" not in project.metadata
        assert "Scenario" not in project.module_titles
    finally:
        dialog.close()


def test_export_dialog_summary_and_apply(qapp, project, tmp_path):
    project.add_screenshot("Scenario::1.1", make_png(tmp_path / "a.png"))
    dialog = ExportDialog(project)
    try:
        text = dialog.summary.toPlainText()
        assert "[SCENARIO]" in text
        assert "test case  : 2" in text
        assert "screenshot : 1" in text
        assert "1.2  Case Dua" in text

        assert dialog.selected_modules() == ["Scenario"]
        dialog.combined.setChecked(True)
        dialog.template_edit.setText(str(tmp_path / "t.docx"))
        dialog.apply()

        assert project.output_mode == "combined"
        assert project.template == str(tmp_path / "t.docx")
    finally:
        dialog.close()


def test_export_dialog_refuses_an_empty_selection(qapp, project):
    dialog = ExportDialog(project)
    try:
        dialog._module_checks["Scenario"].setChecked(False)
        assert dialog.selected_modules() == []
        assert "Belum ada module dipilih" in dialog.summary.toPlainText()
    finally:
        dialog.close()


def test_orphan_dialog_restores_to_another_test_case(qapp, project, tmp_path, excel):
    project.add_screenshot("Scenario::1.2", make_png(tmp_path / "a.png"), caption="Pindah")
    simple_workbook(excel, rows=ROWS[:2])
    project.reimport()

    dialog = OrphanDialog(project)
    try:
        assert dialog.table.rowCount() == 1
        assert dialog.table.item(0, 0).text() == "Scenario::1.2"

        dialog.table.setCurrentCell(0, 0)
        index = dialog.target.findData("Scenario::1.1")
        dialog.target.setCurrentIndex(index)
        dialog._restore()

        assert project.orphan_count() == 0
        assert project.screenshot_count("Scenario::1.1") == 1
        assert dialog.table.rowCount() == 0
    finally:
        dialog.close()
