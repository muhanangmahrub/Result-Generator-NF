"""Offscreen tests for the scenario-change review dialog and its main window hooks."""

from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt                                    # noqa: E402
from PySide6.QtWidgets import QApplication, QPushButton                       # noqa: E402

from rtg.config import load_config                               # noqa: E402
from rtg.gui.dialogs import ScenarioChangesDialog                # noqa: E402
from rtg.gui.main_window import MainWindow                       # noqa: E402
from rtg.project import Project                                  # noqa: E402
from rtg.scenario_diff import EDITED                    # noqa: E402
from tests.fixtures import make_png, simple_workbook             # noqa: E402
from tests.test_project_changes import (                         # noqa: E402
    A, B, C, ORIGINAL, STEPS_A, STEPS_B, STEPS_C, rows,
)

CONFIG = load_config()


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def excel(tmp_path):
    return simple_workbook(tmp_path / "src" / "Scenario.xlsx", rows=ORIGINAL)


@pytest.fixture
def project(tmp_path, excel):
    project = Project.create(tmp_path / "proj", excel, config=CONFIG)
    shot = make_png(tmp_path / "shot.png", 320, 180)
    project.add_screenshot("Scenario::2.1", shot, caption="bukti A")
    project.add_screenshot("Scenario::2.2", shot, caption="bukti B")
    return project


def renumber(excel):
    simple_workbook(excel, rows=rows(("2.1", C, STEPS_C), ("2.2", A, STEPS_A), ("2.3", B, STEPS_B)))


def test_dialog_preselects_suggestions_and_applies_them(qapp, project, excel):
    renumber(excel)
    report = project.reimport()
    dialog = ScenarioChangesDialog(project, report.changes)

    assert dialog.table.rowCount() == 2
    mapping, _steps = dialog.remap_selection()
    assert mapping == {"Scenario::2.1": "Scenario::2.2", "Scenario::2.2": "Scenario::2.3"}

    dialog.apply()
    assert project.entries("Scenario::2.2")[0].caption == "bukti A"
    assert project.entries("Scenario::2.3")[0].caption == "bukti B"
    assert project.needs_review == {}


def test_unticked_row_is_not_moved(qapp, project, excel):
    renumber(excel)
    dialog = ScenarioChangesDialog(project, project.reimport().changes)
    dialog.table.item(0, dialog.COL_APPLY).setCheckState(Qt.CheckState.Unchecked)

    mapping, _ = dialog.remap_selection()
    assert mapping == {"Scenario::2.2": "Scenario::2.3"}


def test_user_can_pick_another_target(qapp, project, excel):
    renumber(excel)
    dialog = ScenarioChangesDialog(project, project.reimport().changes)
    combo = dialog._combos[0]
    combo.setCurrentIndex(combo.findData("Scenario::2.3"))
    dialog.table.item(1, dialog.COL_APPLY).setCheckState(Qt.CheckState.Unchecked)

    mapping, step_maps = dialog.remap_selection()
    assert mapping == {"Scenario::2.1": "Scenario::2.3"}
    assert "Scenario::2.1" in step_maps          # recomputed for the other target


def test_edited_row_is_only_acknowledged_when_ticked(qapp, project, excel):
    simple_workbook(excel, rows=rows(("2.1", A + " (revisi)", STEPS_A), ("2.2", B, STEPS_B)))
    report = project.reimport()
    assert report.changes.get("Scenario::2.1").status == EDITED
    dialog = ScenarioChangesDialog(project, report.changes)

    assert dialog.acknowledged() == []
    dialog.table.item(0, dialog.COL_APPLY).setCheckState(Qt.CheckState.Checked)
    assert dialog.acknowledged() == ["Scenario::2.1"]
    dialog.apply()
    assert project.needs_review == {}


def test_undo_button_only_when_backup_exists(qapp, project, excel):
    renumber(excel)
    report = project.reimport()

    def labels(dialog):
        return [b.text() for b in dialog.findChildren(QPushButton)]

    assert "Undo re-import" in labels(ScenarioChangesDialog(project, report.changes))
    assert "Undo re-import" not in labels(
        ScenarioChangesDialog(project, report.changes, allow_undo=False)
    )


def test_tree_marks_cases_needing_review(qapp, project, excel):
    simple_workbook(excel, rows=rows(("2.1", A + " (revisi)", STEPS_A), ("2.2", B, STEPS_B)))
    project.reimport()
    window = MainWindow(project)

    case_item = window.tree.topLevelItem(0).child(0).child(0)
    assert "perlu dicek" in case_item.text(2)
    assert "Nama" in case_item.toolTip(0)
    window.mark_reviewed("Scenario::2.1")
    case_item = window.tree.topLevelItem(0).child(0).child(0)
    assert "perlu dicek" not in case_item.text(2)


def test_warnings_tab_mentions_review_flags(qapp, project, excel):
    simple_workbook(excel, rows=rows(("2.1", A + " (revisi)", STEPS_A), ("2.2", B, STEPS_B)))
    project.reimport()
    window = MainWindow(project)
    assert "berubah sejak bukti diambil" in window.warnings.toPlainText()
