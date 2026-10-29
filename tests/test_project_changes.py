"""Project behaviour when the scenario workbook is edited after capturing."""

from __future__ import annotations

import json

import pytest

from rtg.config import load_config
from rtg.project import MAX_BACKUPS, PROJECT_FILENAME, Project, ProjectError
from rtg.scenario_diff import EDITED, MISSING, SUSPECT
from tests.fixtures import make_png, simple_workbook

CONFIG = load_config()

A = "Upload file partner ke sharefolder"
B = "Validasi format tanggal salah"
C = "Hapus data customer lama"


def rows(*cases):
    """cases: (case id, case name, [step texts]) -> workbook rows."""
    out = []
    for cid, name, steps in cases:
        for i, item in enumerate(steps):
            # a step is its text, or (number, text) to keep the ID after a delete
            number, text = item if isinstance(item, tuple) else (i + 1, item)
            row = {"test_step_id": f"{cid}.{number}", "test_step": text,
                   "expected_results": f"hasil {text}"}
            if i == 0:
                row.update(scenario_id="SC-1", test_scenario="Alpha",
                           test_case_id=cid, test_case=name)
            out.append(row)
    return out


STEPS_A = ["buka sharefolder", "upload file csv", "cek status"]
STEPS_B = ["isi tanggal salah", "klik submit", "lihat pesan error"]
STEPS_C = ["pilih customer", "klik delete", "konfirmasi hapus"]
ORIGINAL = rows(("2.1", A, STEPS_A), ("2.2", B, STEPS_B))


@pytest.fixture
def excel(tmp_path):
    return simple_workbook(tmp_path / "src" / "Scenario.xlsx", rows=ORIGINAL)


@pytest.fixture
def project(tmp_path, excel):
    return Project.create(tmp_path / "proj", excel, config=CONFIG)


@pytest.fixture
def shot(tmp_path):
    return make_png(tmp_path / "incoming" / "capture.png", 640, 360)


def edit_excel(excel, *cases):
    simple_workbook(excel, rows=rows(*cases))


def test_snapshot_is_taken_with_first_evidence_and_persisted(project, shot):
    project.add_screenshot("Scenario::2.1", shot)
    project.set_result("Scenario::2.2", "2.2.1", {"pass_fail": "P"})

    assert set(project.snapshots) == {"Scenario::2.1", "Scenario::2.2"}
    reopened = Project.open(project.folder, config=CONFIG)
    assert reopened.snapshots["Scenario::2.1"].test_case == A
    assert reopened.excel_sha256 == project.excel_sha256 != ""


def test_snapshot_is_not_overwritten_by_later_evidence(project, shot):
    project.add_screenshot("Scenario::2.1", shot)
    taken = project.snapshots["Scenario::2.1"].taken_at
    project.snapshots["Scenario::2.1"].test_case = "lama"
    project.add_screenshot("Scenario::2.1", shot)
    assert project.snapshots["Scenario::2.1"].test_case == "lama"
    assert project.snapshots["Scenario::2.1"].taken_at == taken


def test_snapshot_is_dropped_once_evidence_is_gone(project, shot):
    project.add_screenshot("Scenario::2.1", shot)
    project.delete_screenshot("Scenario::2.1", 0)
    data = json.loads(project.project_file.read_text(encoding="utf-8"))
    assert data["snapshots"] == {}


def test_old_project_without_snapshots_gets_a_baseline(project, shot):
    project.add_screenshot("Scenario::2.1", shot)
    path = project.project_file
    data = json.loads(path.read_text(encoding="utf-8"))
    for key in ("snapshots", "needs_review", "excel_sha256", "orphan_step_results"):
        data.pop(key, None)
    path.write_text(json.dumps(data), encoding="utf-8")

    reopened = Project.open(project.folder, config=CONFIG)
    assert "Scenario::2.1" in reopened.snapshots
    assert reopened.excel_changed is False


def test_renumber_is_detected_and_remap_moves_evidence_in_a_chain(project, shot, excel):
    project.add_screenshot("Scenario::2.1", shot, caption="bukti A")
    project.add_screenshot("Scenario::2.2", shot, caption="bukti B")
    project.set_result("Scenario::2.1", "2.1.2", {"pass_fail": "P"})

    edit_excel(excel, ("2.1", C, STEPS_C), ("2.2", A, STEPS_A), ("2.3", B, STEPS_B))
    report = project.reimport()

    first = report.changes.get("Scenario::2.1")
    second = report.changes.get("Scenario::2.2")
    assert (first.status, first.suggested_uid) == (SUSPECT, "Scenario::2.2")
    assert (second.status, second.suggested_uid) == (SUSPECT, "Scenario::2.3")
    # nothing moved on its own
    assert project.entries("Scenario::2.1")[0].caption == "bukti A"
    assert set(project.needs_review) == {"Scenario::2.1", "Scenario::2.2"}

    project.apply_remap(
        {c.uid: c.suggested_uid for c in report.changes.changes},
        {c.uid: c.step_map for c in report.changes.changes},
    )

    assert project.screenshot_count("Scenario::2.1") == 0
    assert project.entries("Scenario::2.2")[0].caption == "bukti A"
    assert project.entries("Scenario::2.3")[0].caption == "bukti B"
    assert project.result_for("Scenario::2.2", "2.2.2") == {"pass_fail": "P"}
    assert project.needs_review == {}
    assert project.snapshots["Scenario::2.2"].test_case == A


def test_remap_can_swap_two_test_cases(project, shot, excel):
    project.add_screenshot("Scenario::2.1", shot, caption="bukti A")
    project.add_screenshot("Scenario::2.2", shot, caption="bukti B")

    edit_excel(excel, ("2.1", B, STEPS_B), ("2.2", A, STEPS_A))
    project.reimport()
    project.apply_remap({"Scenario::2.1": "Scenario::2.2", "Scenario::2.2": "Scenario::2.1"})

    assert project.entries("Scenario::2.1")[0].caption == "bukti B"
    assert project.entries("Scenario::2.2")[0].caption == "bukti A"


def test_remap_appends_to_a_target_that_already_has_evidence(project, shot, excel):
    project.add_screenshot("Scenario::2.1", shot, caption="lama")
    edit_excel(excel, ("2.2", B, STEPS_B), ("9.9", A, STEPS_A))
    project.reimport()
    project.add_screenshot("Scenario::9.9", shot, caption="baru")
    project.apply_remap({"Scenario::2.1": "Scenario::9.9"})
    assert [e.caption for e in project.entries("Scenario::9.9")] == ["baru", "lama"]


def test_remap_rejects_unknown_target(project, shot):
    project.add_screenshot("Scenario::2.1", shot)
    with pytest.raises(ProjectError, match="tujuan tidak ada"):
        project.apply_remap({"Scenario::2.1": "Scenario::7.7"})


def test_edited_case_is_flagged_and_acknowledge_clears_it(project, shot, excel):
    project.add_screenshot("Scenario::2.1", shot)
    edit_excel(excel, ("2.1", A + " (revisi)", STEPS_A), ("2.2", B, STEPS_B))
    report = project.reimport()

    assert report.changes.get("Scenario::2.1").status == EDITED
    assert "Scenario::2.1" in project.needs_review
    assert project.screenshot_count("Scenario::2.1") == 1     # evidence stays

    project.acknowledge("Scenario::2.1")
    assert project.needs_review == {}
    assert project.snapshots["Scenario::2.1"].test_case == A + " (revisi)"
    assert not project.check_changes().has_changes


def test_sheet_rename_is_reported_as_missing_with_suggestion(tmp_path, shot):
    excel = simple_workbook(tmp_path / "src" / "S.xlsx", rows=ORIGINAL, sheet_name="Old")
    project = Project.create(tmp_path / "proj", excel, config=CONFIG)
    project.add_screenshot("Old::2.1", shot)

    simple_workbook(excel, rows=ORIGINAL, sheet_name="New")
    report = project.reimport(sheets=["New"])

    change = report.changes.get("Old::2.1")
    assert (change.status, change.suggested_uid) == (MISSING, "New::2.1")
    assert project.orphan_uids == ["Old::2.1"]

    project.apply_remap({"Old::2.1": "New::2.1"})
    assert project.screenshot_count("New::2.1") == 1
    assert project.orphan_count() == 0


def test_result_of_a_deleted_step_is_parked_not_lost(project, excel):
    project.set_result("Scenario::2.1", "2.1.2", {"pass_fail": "F", "actual_results": "x"})
    project.set_result("Scenario::2.1", "2.1.1", {"pass_fail": "P"})

    edit_excel(excel, ("2.1", A, [(1, "buka sharefolder"), (3, "cek status")]), ("2.2", B, STEPS_B))
    project.reimport()

    assert project.result_for("Scenario::2.1", "2.1.2") == {}
    assert project.orphan_step_results["Scenario::2.1"]["2.1.2"]["pass_fail"] == "F"
    assert project.result_for("Scenario::2.1", "2.1.1") == {"pass_fail": "P"}
    # survives a reopen
    reopened = Project.open(project.folder, config=CONFIG)
    assert "2.1.2" in reopened.orphan_step_results["Scenario::2.1"]
    # and comes back when the step does
    edit_excel(excel, ("2.1", A, STEPS_A), ("2.2", B, STEPS_B))
    reopened.reimport()
    assert reopened.result_for("Scenario::2.1", "2.1.2")["pass_fail"] == "F"
    assert reopened.orphan_step_results == {}


def test_reimport_creates_backup_and_restore_undoes_it(project, shot, excel):
    project.add_screenshot("Scenario::2.2", shot)
    edit_excel(excel, ("2.1", A, STEPS_A))                    # 2.2 disappears
    project.reimport()
    assert project.orphan_uids == ["Scenario::2.2"]
    assert project.latest_backup() is not None

    project.restore_backup()
    assert project.orphan_uids == []
    assert project.screenshot_count("Scenario::2.2") == 1


def test_backups_are_capped(project, excel):
    for _ in range(MAX_BACKUPS + 4):
        project.reimport()
    assert len(project.backups()) == MAX_BACKUPS


def test_restore_without_backup_errors(project):
    with pytest.raises(ProjectError, match="backup"):
        project.restore_backup()


def test_open_after_excel_was_edited_reconciles(project, shot, excel):
    project.add_screenshot("Scenario::2.1", shot)
    project.add_screenshot("Scenario::2.2", shot)
    edit_excel(excel, ("2.1", A + " (revisi)", STEPS_A))      # 2.2 deleted, 2.1 edited

    reopened = Project.open(project.folder, config=CONFIG)

    assert reopened.excel_changed is True
    assert reopened.orphan_uids == ["Scenario::2.2"]
    assert "Scenario::2.1" in reopened.needs_review
    assert reopened.last_changes.get("Scenario::2.1").status == EDITED
    assert reopened.backups()
    # the change is only announced once
    again = Project.open(project.folder, config=CONFIG)
    assert again.excel_changed is False
    assert (project.folder / PROJECT_FILENAME).is_file()


# --------------------------------------------------------------------------- #
# Duplicate Test Case IDs stay separate test cases
# --------------------------------------------------------------------------- #

DUPLICATED = rows(
    ("1.1", "Login benar", ["buka halaman", "isi data benar"]),
    ("1.2", "Logout", ["klik logout"]),
    ("1.1", "Login salah", [(3, "isi data salah"), (4, "lihat pesan error")]),
)


def test_duplicate_id_screenshots_go_to_separate_folders(tmp_path, shot):
    excel = simple_workbook(tmp_path / "src" / "Dup.xlsx", rows=DUPLICATED)
    project = Project.create(tmp_path / "proj", excel, config=CONFIG)
    assert [c.uid for c in project.test_cases] == [
        "Scenario::1.1", "Scenario::1.2", "Scenario::1.1~2",
    ]

    first = project.add_screenshot("Scenario::1.1", shot)
    second = project.add_screenshot("Scenario::1.1~2", shot)

    assert project.absolute(first).parent != project.absolute(second).parent
    assert "1.1-2_" in second.file
    assert project.screenshot_count("Scenario::1.1") == 1
    assert project.screenshot_count("Scenario::1.1~2") == 1


def test_results_of_a_formerly_merged_duplicate_move_to_their_own_case(tmp_path):
    excel = simple_workbook(tmp_path / "src" / "Dup.xlsx", rows=DUPLICATED)
    project = Project.create(tmp_path / "proj", excel, config=CONFIG)
    # What an older version saved: every step of both cases under one uid.
    project.results["Scenario::1.1"] = {
        "1.1.1": {"pass_fail": "P"},
        "1.1.3": {"pass_fail": "F"},
    }

    project.reimport()

    assert project.result_for("Scenario::1.1", "1.1.1") == {"pass_fail": "P"}
    assert project.result_for("Scenario::1.1~2", "1.1.3") == {"pass_fail": "F"}
    assert project.result_for("Scenario::1.1", "1.1.3") == {}
    assert project.orphan_step_results == {}
