"""Detect what changed in a scenario workbook since evidence was captured.

When the first screenshot or result of a test case is stored, a *snapshot* of
that test case (name, steps, expected results) is saved in ``project.json``.
When the workbook is read again, the snapshots are compared with the current
test cases so that renumbered, renamed, edited or removed test cases are
reported instead of silently keeping evidence on the wrong test case.

Nothing here touches the project or the disk; ``Project`` decides what to do
with the result. Suggestions are only suggestions: the user confirms them.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from .models import ScenarioFile, TestCase, TestStep

#: Below this similarity an existing uid is treated as a different test case.
SIMILAR_THRESHOLD = 0.6
#: Minimum similarity before another test case is suggested as the new home.
CANDIDATE_THRESHOLD = 0.8
#: Two candidates closer than this are ambiguous and are not pre-selected.
AMBIGUOUS_MARGIN = 0.05

UNCHANGED = "unchanged"
EDITED = "edited"
SUSPECT = "suspect"
MISSING = "missing"
REMOVED = "removed"

_SPACE = re.compile(r"\s+")


def step_key(step: TestStep, index: int) -> str:
    """Stable identity of a step inside its test case.

    The Test Step ID ("2.1.1") survives rows being inserted above it, so it is
    preferred; a sheet without that column falls back to the position.
    """
    return step.test_step_id.strip() or f"#{index}"


def _norm(text: Any) -> str:
    return _SPACE.sub(" ", str(text or "")).strip().lower()


# --------------------------------------------------------------------------- #
# Snapshots
# --------------------------------------------------------------------------- #

@dataclass
class StepSnapshot:
    key: str
    test_step: str = ""
    expected_results: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"key": self.key}
        if self.test_step:
            data["test_step"] = self.test_step
        if self.expected_results:
            data["expected_results"] = self.expected_results
        return data

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "StepSnapshot":
        return cls(
            key=str(data.get("key", "")),
            test_step=str(data.get("test_step", "")),
            expected_results=str(data.get("expected_results", "")),
        )


@dataclass
class CaseSnapshot:
    module: str
    test_case_id: str
    test_case: str = ""
    steps: List[StepSnapshot] = field(default_factory=list)
    taken_at: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "module": self.module,
            "test_case_id": self.test_case_id,
            "test_case": self.test_case,
            "steps": [step.to_dict() for step in self.steps],
            "taken_at": self.taken_at,
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "CaseSnapshot":
        return cls(
            module=str(data.get("module", "")),
            test_case_id=str(data.get("test_case_id", "")),
            test_case=str(data.get("test_case", "")),
            steps=[StepSnapshot.from_dict(item) for item in data.get("steps") or []],
            taken_at=str(data.get("taken_at", "")),
        )


def snapshot_of(case: TestCase, *, taken_at: str = "") -> CaseSnapshot:
    return CaseSnapshot(
        module=case.module,
        test_case_id=case.test_case_id,
        test_case=case.test_case,
        steps=[
            StepSnapshot(
                key=step_key(step, index),
                test_step=step.test_step,
                expected_results=step.expected_results,
            )
            for index, step in enumerate(case.steps, start=1)
        ],
        taken_at=taken_at,
    )


def case_text(item: Any) -> str:
    """Normalised name + steps + expected results of a snapshot or test case."""
    parts = [_norm(item.test_case)]
    for step in item.steps:
        parts.append(_norm(step.test_step))
        parts.append(_norm(step.expected_results))
    return "\n".join(parts)


def similarity(a: Any, b: Any) -> float:
    """0..1 text similarity between two snapshots / test cases."""
    return difflib.SequenceMatcher(None, case_text(a), case_text(b)).ratio()


def _text_ratio(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, _norm(a), _norm(b)).ratio()


# --------------------------------------------------------------------------- #
# Step matching
# --------------------------------------------------------------------------- #

def match_steps(
    old: Sequence[StepSnapshot], new: Sequence[TestStep]
) -> Dict[str, str]:
    """Pair old steps with the steps of the new test case: ``old_key -> new_key``.

    Order of attempts: identical step text, then the same key with similar
    text, then position when the number of leftovers is equal.
    """
    new_keys = [step_key(step, index) for index, step in enumerate(new, start=1)]
    mapping: Dict[str, str] = {}
    used_old: set = set()
    used_new: set = set()

    for oi, old_step in enumerate(old):
        text = _norm(old_step.test_step)
        if not text:
            continue
        for ni, step in enumerate(new):
            if ni not in used_new and _norm(step.test_step) == text:
                mapping[old_step.key] = new_keys[ni]
                used_old.add(oi)
                used_new.add(ni)
                break

    for oi, old_step in enumerate(old):
        if oi in used_old:
            continue
        for ni, step in enumerate(new):
            if ni in used_new or new_keys[ni] != old_step.key:
                continue
            if _text_ratio(old_step.test_step, step.test_step) >= SIMILAR_THRESHOLD:
                mapping[old_step.key] = new_keys[ni]
                used_old.add(oi)
                used_new.add(ni)
                break

    left_old = [oi for oi in range(len(old)) if oi not in used_old]
    left_new = [ni for ni in range(len(new)) if ni not in used_new]
    if left_old and len(left_old) == len(left_new):
        for oi, ni in zip(left_old, left_new):
            mapping[old[oi].key] = new_keys[ni]
    return mapping


# --------------------------------------------------------------------------- #
# Diff
# --------------------------------------------------------------------------- #

@dataclass
class CaseChange:
    uid: str
    status: str
    suggested_uid: Optional[str] = None
    score: float = 0.0
    preselect: bool = False
    details: List[str] = field(default_factory=list)
    step_map: Dict[str, str] = field(default_factory=dict)
    candidates: List[Tuple[str, float]] = field(default_factory=list)

    @property
    def needs_attention(self) -> bool:
        return self.status != UNCHANGED


@dataclass
class ScenarioChanges:
    changes: List[CaseChange] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return bool(self.changes)

    @property
    def needs_attention(self) -> bool:
        return any(
            change.status not in (UNCHANGED, REMOVED) for change in self.changes
        )

    def by_status(self, status: str) -> List[CaseChange]:
        return [change for change in self.changes if change.status == status]

    def get(self, uid: str) -> Optional[CaseChange]:
        for change in self.changes:
            if change.uid == uid:
                return change
        return None

    def summary_lines(self) -> List[str]:
        labels = (
            (EDITED, "test case berubah isinya (cek ulang buktinya)"),
            (SUSPECT, "test case kemungkinan bergeser/ter-renumber"),
            (MISSING, "test case hilang tapi ada kandidat pengganti"),
            (REMOVED, "test case hilang tanpa pengganti"),
        )
        lines = []
        for status, label in labels:
            count = len(self.by_status(status))
            if count:
                lines.append(f"{count} {label}")
        return lines


def _step_details(
    snapshot: CaseSnapshot, case: TestCase, results: Mapping[str, Any]
) -> Tuple[List[str], Dict[str, str]]:
    details: List[str] = []
    if _norm(snapshot.test_case) != _norm(case.test_case):
        details.append(f"Nama: '{snapshot.test_case}' -> '{case.test_case}'")

    step_map = match_steps(snapshot.steps, case.steps)
    new_keys = [step_key(s, i) for i, s in enumerate(case.steps, start=1)]
    new_by_key = dict(zip(new_keys, case.steps))

    for step in snapshot.steps:
        suffix = " (punya hasil eksekusi)" if results.get(step.key) else ""
        target = step_map.get(step.key)
        if target is None:
            details.append(f"Step {step.key} dihapus{suffix}")
        elif target != step.key:
            details.append(f"Step {step.key} sekarang bernomor {target}{suffix}")
        else:
            new_step = new_by_key[target]
            if (_norm(new_step.test_step) != _norm(step.test_step)
                    or _norm(new_step.expected_results) != _norm(step.expected_results)):
                details.append(f"Step {step.key} diubah isinya{suffix}")

    mapped = set(step_map.values())
    added = [key for key in new_keys if key not in mapped]
    if added:
        details.append("Step baru: " + ", ".join(added))
    return details, step_map


def diff_scenarios(
    snapshots: Mapping[str, CaseSnapshot],
    evidence_uids: Iterable[str],
    scenario_file: ScenarioFile,
    results: Optional[Mapping[str, Mapping[str, Any]]] = None,
) -> ScenarioChanges:
    """Compare snapshots of evidence-bearing test cases with the new scenario.

    Unchanged test cases are not reported. Test cases without a snapshot are
    skipped (the caller creates a baseline for them).
    """
    results = results or {}
    cases = {case.uid: case for case in scenario_file.test_cases}
    sources = [uid for uid in dict.fromkeys(evidence_uids) if uid in snapshots]

    unchanged = {
        uid for uid in sources
        if uid in cases and case_text(snapshots[uid]) == case_text(cases[uid])
    }

    # Score every candidate target for the sources that did change.
    scored: Dict[str, List[Tuple[str, float]]] = {}
    for uid in sources:
        if uid in unchanged:
            continue
        ranked = sorted(
            (
                (other_uid, similarity(snapshots[uid], other))
                for other_uid, other in cases.items()
                if other_uid != uid
            ),
            key=lambda item: item[1],
            reverse=True,
        )
        scored[uid] = [item for item in ranked if item[1] >= CANDIDATE_THRESHOLD][:3]

    # Greedy one-to-one assignment, best score first.
    pairs = sorted(
        ((score, uid, target) for uid, items in scored.items() for target, score in items),
        reverse=True,
    )
    taken_sources: set = set()
    taken_targets: set = set()
    assigned: Dict[str, Tuple[str, float]] = {}
    for score, uid, target in pairs:
        if uid in taken_sources or target in taken_targets:
            continue
        assigned[uid] = (target, score)
        taken_sources.add(uid)
        taken_targets.add(target)

    changes: List[CaseChange] = []
    for uid in sources:
        if uid in unchanged:
            continue
        snapshot = snapshots[uid]
        stored = results.get(uid) or {}
        candidates = scored.get(uid, [])
        suggestion = assigned.get(uid)
        exists = uid in cases
        own_score = similarity(snapshot, cases[uid]) if exists else 0.0

        if exists and (own_score >= SIMILAR_THRESHOLD or suggestion is None):
            details, step_map = _step_details(snapshot, cases[uid], stored)
            if own_score < SIMILAR_THRESHOLD:
                details.insert(
                    0, "Isi test case berubah drastis dan tidak ada pengganti yang mirip"
                )
            changes.append(CaseChange(
                uid=uid, status=EDITED, score=own_score,
                details=details, step_map=step_map, candidates=candidates,
            ))
            continue

        if suggestion is None:
            changes.append(CaseChange(
                uid=uid, status=REMOVED,
                details=[f"Test case '{snapshot.test_case}' tidak ada lagi di Excel"],
            ))
            continue

        target, score = suggestion
        runner_up = next((s for t, s in candidates if t != target), 0.0)
        ambiguous = score - runner_up < AMBIGUOUS_MARGIN
        target_in_use = target in unchanged
        details, step_map = _step_details(snapshot, cases[target], stored)
        if exists:
            details.insert(0, f"Isi '{uid}' sekarang berbeda; yang mirip: '{target}'")
        else:
            details.insert(0, f"'{uid}' hilang; yang mirip: '{target}'")
        if ambiguous:
            details.append("Ada kandidat lain yang sama miripnya - pilih manual")
        if target_in_use:
            details.append(f"'{target}' sudah punya bukti sendiri yang tidak berubah")
        changes.append(CaseChange(
            uid=uid,
            status=SUSPECT if exists else MISSING,
            suggested_uid=target,
            score=score,
            preselect=not ambiguous and not target_in_use,
            details=details,
            step_map=step_map,
            candidates=candidates,
        ))
    return ScenarioChanges(changes=changes)
