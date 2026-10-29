"""Unit tests for snapshot / diff of scenario changes."""

from __future__ import annotations

from rtg.models import Module, ScenarioFile, ScenarioGroup, TestCase, TestStep
from rtg.scenario_diff import (
    EDITED,
    MISSING,
    REMOVED,
    SUSPECT,
    CaseSnapshot,
    diff_scenarios,
    match_steps,
    snapshot_of,
)


def make_case(cid, name, steps, module="M"):
    return TestCase(
        module=module,
        test_case_id=cid,
        test_case=name,
        steps=[
            TestStep(row=i + 2, test_step_id=f"{cid}.{i + 1}", test_step=text,
                     expected_results=f"hasil {text}")
            for i, text in enumerate(steps)
        ],
    )


def scenario(*cases):
    group = ScenarioGroup(test_cases=list(cases))
    return ScenarioFile(path="x.xlsx", modules=[Module(name="M", scenarios=[group])])


A = ("Upload file partner ke sharefolder", ["buka sharefolder", "upload file csv", "cek status"])
B = ("Validasi format tanggal salah", ["isi tanggal 99-99-9999", "klik submit", "lihat pesan error"])
C = ("Hapus data customer lama", ["pilih customer", "klik delete", "konfirmasi hapus"])


def snaps(*cases):
    return {c.uid: snapshot_of(c) for c in cases}


def test_unchanged_reports_nothing():
    old = [make_case("1.1", *A), make_case("1.2", *B)]
    result = diff_scenarios(snaps(*old), [c.uid for c in old], scenario(*old))
    assert not result.has_changes


def test_snapshot_roundtrip():
    snap = snapshot_of(make_case("1.1", *A), taken_at="t")
    assert CaseSnapshot.from_dict(snap.to_dict()) == snap


def test_renamed_case_is_edited():
    old = [make_case("1.1", *A)]
    new = [make_case("1.1", "Upload file partner ke sharefolder (revisi)", A[1])]
    result = diff_scenarios(snaps(*old), ["M::1.1"], scenario(*new))
    change = result.get("M::1.1")
    assert change.status == EDITED
    assert any("Nama" in line for line in change.details)


def test_inserted_case_shifts_ids_and_suggests_remap():
    old = [make_case("2.1", *A), make_case("2.2", *B)]
    new = [make_case("2.1", *C), make_case("2.2", *A), make_case("2.3", *B)]
    result = diff_scenarios(snaps(*old), ["M::2.1", "M::2.2"], scenario(*new))

    first, second = result.get("M::2.1"), result.get("M::2.2")
    assert (first.status, first.suggested_uid, first.preselect) == (SUSPECT, "M::2.2", True)
    assert (second.status, second.suggested_uid, second.preselect) == (SUSPECT, "M::2.3", True)


def test_removed_case_without_candidate():
    old = [make_case("1.1", *A), make_case("1.2", *B)]
    new = [make_case("1.1", *A)]
    result = diff_scenarios(snaps(*old), ["M::1.2"], scenario(*new))
    assert result.get("M::1.2").status == REMOVED


def test_missing_case_with_renumbered_candidate():
    old = [make_case("1.1", *A)]
    new = [make_case("9.9", *A)]
    result = diff_scenarios(snaps(*old), ["M::1.1"], scenario(*new))
    change = result.get("M::1.1")
    assert (change.status, change.suggested_uid) == (MISSING, "M::9.9")


def test_ambiguous_candidates_are_not_preselected():
    old = [make_case("1.1", *A)]
    new = [make_case("5.1", *A), make_case("5.2", *A)]
    result = diff_scenarios(snaps(*old), ["M::1.1"], scenario(*new))
    change = result.get("M::1.1")
    assert change.status == MISSING
    assert change.preselect is False


def test_sheet_rename_suggests_other_module():
    old = [make_case("1.1", *A, module="Old")]
    new = [make_case("1.1", *A, module="New")]
    result = diff_scenarios(snaps(*old), ["Old::1.1"], scenario(*new))
    assert result.get("Old::1.1").suggested_uid == "New::1.1"


def test_target_with_unchanged_evidence_is_not_preselected():
    old = [make_case("1.1", *A), make_case("1.2", *B)]
    # 1.1 now reads like 1.2 (which is itself unchanged and has evidence).
    new = [make_case("1.1", *B), make_case("1.2", *B)]
    result = diff_scenarios(snaps(*old), ["M::1.1", "M::1.2"], scenario(*new))
    change = result.get("M::1.1")
    assert change is not None and change.preselect is False


def test_case_without_snapshot_is_skipped():
    new = [make_case("1.1", *A)]
    assert not diff_scenarios({}, ["M::1.1"], scenario(*new)).has_changes


def test_match_steps_inserted_in_middle():
    old = snapshot_of(make_case("1.1", *A)).steps
    new = make_case("1.1", "x", ["buka sharefolder", "langkah baru", "upload file csv", "cek status"])
    mapping = match_steps(old, new.steps)
    assert mapping == {"1.1.1": "1.1.1", "1.1.2": "1.1.3", "1.1.3": "1.1.4"}


def test_match_steps_removed_step():
    old = snapshot_of(make_case("1.1", *A)).steps
    new = make_case("1.1", "x", ["buka sharefolder", "cek status"])
    mapping = match_steps(old, new.steps)
    assert mapping == {"1.1.1": "1.1.1", "1.1.3": "1.1.2"}
    assert "1.1.2" not in mapping


def test_step_details_flag_removed_step_with_result():
    old = [make_case("1.1", *A)]
    new = [make_case("1.1", A[0], ["buka sharefolder", "cek status"])]
    result = diff_scenarios(
        snaps(*old), ["M::1.1"], scenario(*new),
        results={"M::1.1": {"1.1.2": {"pass_fail": "P"}}},
    )
    lines = " ".join(result.get("M::1.1").details)
    assert "1.1.2 dihapus (punya hasil eksekusi)" in lines
