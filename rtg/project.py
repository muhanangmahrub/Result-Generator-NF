"""Project folder handling: project.json, screenshots/ and output/.

One project is one folder:

    <project>/
      project.json          path of the Excel file, selected sheets, manual
                            column mapping, metadata, module titles and the
                            screenshot mapping
      screenshots/          <module-slug>/<TestCaseID>_<test-case-slug>/*.png
      output/               generated .docx files

Screenshot files are stored relative to the project folder, so the whole folder
can be moved, copied or zipped and keeps working.

The scenario workbook itself is only ever read; re-importing refreshes the
scenarios without touching the screenshots already taken. A screenshot whose
test case disappeared from the workbook is not deleted but marked *orphan*, and
is picked up again automatically if that test case comes back.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

from . import __version__
from .capture import CAPTURE_MODES, MODE_FULLSCREEN
from .config import FieldConfig, load_config
from .docx_builder import ExportResult, ModuleReport, Screenshot
from .docx_builder import export as export_documents
from .docx_builder import resolve_naming, summarize
from .errors import ResultGeneratorError
from .excel_writer import RESULT_FIELDS, RowUpdate, WriteReport
from .excel_writer import default_result_name, write_results
from .excel_parser import SheetOverride, WorkbookInspection, inspect_workbook, parse_workbook
from .models import ScenarioFile, TestCase, TestStep
from .normalize import slugify
from .scenario_diff import (
    EDITED, MISSING, REMOVED, SUSPECT, CaseSnapshot, ScenarioChanges,
    diff_scenarios, snapshot_of, step_key,
)

PROJECT_FILENAME = "project.json"
SCREENSHOTS_DIRNAME = "screenshots"
OUTPUT_DIRNAME = "output"
SCHEMA_VERSION = 1
DEFAULT_HOTKEY = "Ctrl+Shift+S"
RESULTS_DIRNAME = "excel"
BACKUPS_DIRNAME = "backups"
MAX_BACKUPS = 10


class ProjectError(ResultGeneratorError):
    """Raised for problems with the project folder itself."""


def _now() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")


def _timestamp() -> str:
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def _file_sha256(path: Path) -> str:
    """Hash of a file, or "" when it cannot be read."""
    digest = hashlib.sha256()
    try:
        with open(path, "rb") as handle:
            for block in iter(lambda: handle.read(1 << 20), b""):
                digest.update(block)
    except OSError:
        return ""
    return digest.hexdigest()


def _valid_capture_mode(value: Any) -> str:
    """Fall back to full screen rather than failing on an unknown mode."""
    text = str(value or "")
    return text if text in CAPTURE_MODES else MODE_FULLSCREEN


# --------------------------------------------------------------------------- #
# Screenshot entries
# --------------------------------------------------------------------------- #

@dataclass
class ScreenshotEntry:
    """One screenshot stored inside the project folder."""

    file: str                      # path relative to the project folder
    caption: str = ""
    captured_at: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"file": self.file}
        if self.caption:
            data["caption"] = self.caption
        if self.captured_at:
            data["captured_at"] = self.captured_at
        return data

    @classmethod
    def from_dict(cls, data: Any) -> "ScreenshotEntry":
        if isinstance(data, str):
            return cls(file=data)
        return cls(
            file=str(data["file"]),
            caption=str(data.get("caption", "")),
            captured_at=str(data.get("captured_at", "")),
        )


# --------------------------------------------------------------------------- #
# Re-import report
# --------------------------------------------------------------------------- #

@dataclass
class ReimportReport:
    """What changed when the Excel file was re-read."""

    added_modules: List[str] = field(default_factory=list)
    removed_modules: List[str] = field(default_factory=list)
    added_test_cases: List[str] = field(default_factory=list)
    removed_test_cases: List[str] = field(default_factory=list)
    orphaned: Dict[str, int] = field(default_factory=dict)
    restored: Dict[str, int] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    changes: Optional[ScenarioChanges] = None

    @property
    def orphaned_count(self) -> int:
        return sum(self.orphaned.values())

    @property
    def restored_count(self) -> int:
        return sum(self.restored.values())

    @property
    def has_changes(self) -> bool:
        return bool(
            self.added_modules or self.removed_modules
            or self.added_test_cases or self.removed_test_cases
            or self.orphaned or self.restored
            or (self.changes is not None and self.changes.has_changes)
        )

    def summary_lines(self) -> List[str]:
        lines: List[str] = []
        if self.added_modules:
            lines.append(f"Module baru: {', '.join(self.added_modules)}")
        if self.removed_modules:
            lines.append(f"Module hilang: {', '.join(self.removed_modules)}")
        if self.added_test_cases:
            lines.append(f"Test case baru: {len(self.added_test_cases)}")
        if self.removed_test_cases:
            lines.append(f"Test case hilang: {len(self.removed_test_cases)}")
        if self.orphaned:
            lines.append(
                f"{self.orphaned_count} screenshot jadi orphan "
                f"(dari {len(self.orphaned)} test case yang hilang)"
            )
        if self.restored:
            lines.append(
                f"{self.restored_count} screenshot orphan dipulihkan "
                "(test case-nya muncul lagi)"
            )
        if self.changes is not None:
            lines.extend(self.changes.summary_lines())
        if not lines:
            lines.append("Tidak ada perubahan skenario.")
        return lines

    def to_dict(self) -> Dict[str, Any]:
        return {
            "added_modules": list(self.added_modules),
            "removed_modules": list(self.removed_modules),
            "added_test_cases": list(self.added_test_cases),
            "removed_test_cases": list(self.removed_test_cases),
            "orphaned": dict(self.orphaned),
            "restored": dict(self.restored),
            "warnings": list(self.warnings),
            "changes": self.changes.summary_lines() if self.changes else [],
        }


# --------------------------------------------------------------------------- #
# Project
# --------------------------------------------------------------------------- #

_INVALID_PATH_CHARS = set('<>:"/\\|?*')


def _safe_component(value: str, *, fallback: str = "x") -> str:
    """Make one path component safe without mangling the Test Case ID."""
    text = "".join(
        "-" if (ch in _INVALID_PATH_CHARS or ord(ch) < 32) else ch
        for ch in str(value)
    ).strip(" .")
    return text or fallback


@dataclass
class Project:
    """A project folder, its settings and its screenshot mapping."""

    folder: Path
    excel_path: Path
    name: str = ""
    selected_sheets: List[str] = field(default_factory=list)
    column_overrides: Dict[str, SheetOverride] = field(default_factory=dict)
    metadata: Dict[str, str] = field(default_factory=dict)
    module_titles: Dict[str, str] = field(default_factory=dict)
    screenshots: Dict[str, List[ScreenshotEntry]] = field(default_factory=dict)
    orphans: Dict[str, List[ScreenshotEntry]] = field(default_factory=dict)
    results: Dict[str, Dict[str, Dict[str, str]]] = field(default_factory=dict)
    #: Results whose step disappeared from the Excel; kept, never written back.
    orphan_step_results: Dict[str, Dict[str, Dict[str, str]]] = field(default_factory=dict)
    #: Test case contents at the moment evidence was first stored.
    snapshots: Dict[str, CaseSnapshot] = field(default_factory=dict)
    #: uid -> why the evidence should be checked again.
    needs_review: Dict[str, List[str]] = field(default_factory=dict)
    excel_sha256: str = ""
    hotkey: str = DEFAULT_HOTKEY
    capture_mode: str = MODE_FULLSCREEN
    capture_monitor: int = 0
    output_mode: str = "per_module"
    template: Optional[str] = None
    created_at: str = ""
    updated_at: str = ""
    config: FieldConfig = field(default_factory=load_config)
    scenario_file: Optional[ScenarioFile] = None
    warnings: List[str] = field(default_factory=list)
    #: True when the workbook differs from the one read last time (not saved).
    excel_changed: bool = False
    last_changes: Optional[ScenarioChanges] = None

    # -- paths ------------------------------------------------------------- #

    @property
    def project_file(self) -> Path:
        return self.folder / PROJECT_FILENAME

    @property
    def screenshots_dir(self) -> Path:
        return self.folder / SCREENSHOTS_DIRNAME

    @property
    def output_dir(self) -> Path:
        return self.folder / OUTPUT_DIRNAME

    @property
    def template_path(self) -> Optional[Path]:
        if not self.template:
            return None
        path = Path(self.template)
        return path if path.is_absolute() else (self.folder / path)

    @property
    def backups_dir(self) -> Path:
        return self.folder / BACKUPS_DIRNAME

    def absolute(self, entry: ScreenshotEntry) -> Path:
        return self.folder / entry.file

    def _ensure_dirs(self) -> None:
        self.folder.mkdir(parents=True, exist_ok=True)
        self.screenshots_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)

    # -- construction ------------------------------------------------------ #

    @classmethod
    def create(
        cls,
        folder: Path,
        excel_path: Path,
        *,
        sheets: Optional[Iterable[str]] = None,
        overrides: Optional[Mapping[str, Any]] = None,
        name: str = "",
        config: Optional[FieldConfig] = None,
        exist_ok: bool = False,
        hotkey: Optional[str] = None,
        capture_mode: Optional[str] = None,
        capture_monitor: Optional[int] = None,
    ) -> "Project":
        """Create a new project folder and parse the workbook once.

        ``hotkey`` and the capture settings are user preferences rather than
        project data: pass the ones the user last chose so a new project does
        not silently fall back to the factory default.
        """
        folder = Path(folder)
        config = config or load_config(search_from=folder)
        if (folder / PROJECT_FILENAME).exists() and not exist_ok:
            raise ProjectError(
                f"Folder '{folder}' sudah berisi project. Pakai Open Project, "
                "atau pilih folder lain."
            )
        excel_path = Path(excel_path)
        if not excel_path.is_file():
            raise ProjectError(f"File Excel tidak ditemukan: {excel_path}")

        if sheets is None:
            sheets = inspect_workbook(
                excel_path, config=config, overrides=overrides
            ).scenario_sheet_names

        project = cls(
            folder=folder,
            excel_path=excel_path,
            name=name or folder.name,
            selected_sheets=list(sheets),
            column_overrides={
                str(key): SheetOverride.coerce(value)
                for key, value in (overrides or {}).items()
            },
            config=config,
            hotkey=hotkey or DEFAULT_HOTKEY,
            capture_mode=_valid_capture_mode(capture_mode),
            capture_monitor=int(capture_monitor or 0),
            created_at=_now(),
        )
        project._ensure_dirs()
        project.load_scenarios()
        if project.scenario_file is not None:
            project.metadata = dict(project.scenario_file.merged_metadata())
        project.save()
        return project

    @classmethod
    def open(
        cls,
        folder: Path,
        *,
        config: Optional[FieldConfig] = None,
        parse: bool = True,
    ) -> "Project":
        """Open an existing project folder."""
        folder = Path(folder)
        if folder.is_file() and folder.name == PROJECT_FILENAME:
            folder = folder.parent
        path = folder / PROJECT_FILENAME
        if not path.is_file():
            raise ProjectError(
                f"'{folder}' bukan folder project (tidak ada {PROJECT_FILENAME})."
            )
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError as exc:
            raise ProjectError(
                f"{PROJECT_FILENAME} rusak dan tidak bisa dibaca: {exc}",
                detail=str(exc),
            ) from exc

        version = int(data.get("schema_version", SCHEMA_VERSION))
        if version > SCHEMA_VERSION:
            raise ProjectError(
                f"Project ini dibuat oleh versi aplikasi yang lebih baru "
                f"(schema {version} > {SCHEMA_VERSION}). Perbarui aplikasinya."
            )

        excel_path = Path(str(data.get("excel_path", "")))
        if not excel_path.is_absolute():
            excel_path = (folder / excel_path).resolve()

        output = data.get("output") or {}
        project = cls(
            folder=folder,
            excel_path=excel_path,
            name=str(data.get("name") or folder.name),
            selected_sheets=list(data.get("selected_sheets", [])),
            column_overrides={
                str(key): SheetOverride.coerce(value)
                for key, value in (data.get("column_overrides") or {}).items()
            },
            metadata=dict(data.get("metadata") or {}),
            module_titles=dict(data.get("module_titles") or {}),
            screenshots={
                str(uid): [ScreenshotEntry.from_dict(item) for item in items]
                for uid, items in (data.get("screenshots") or {}).items()
            },
            orphans={
                str(uid): [ScreenshotEntry.from_dict(item) for item in items]
                for uid, items in (data.get("orphan_screenshots") or {}).items()
            },
            results={
                str(uid): {
                    str(key): {
                        name: str(value)
                        for name, value in (entry or {}).items()
                        if name in RESULT_FIELDS and str(value)
                    }
                    for key, entry in (steps or {}).items()
                }
                for uid, steps in (data.get("results") or {}).items()
            },
            orphan_step_results={
                str(uid): {
                    str(key): {str(n): str(v) for n, v in (entry or {}).items()}
                    for key, entry in (steps or {}).items()
                }
                for uid, steps in (data.get("orphan_step_results") or {}).items()
            },
            snapshots={
                str(uid): CaseSnapshot.from_dict(item)
                for uid, item in (data.get("snapshots") or {}).items()
            },
            needs_review={
                str(uid): [str(line) for line in lines]
                for uid, lines in (data.get("needs_review") or {}).items()
            },
            excel_sha256=str(data.get("excel_sha256") or ""),
            hotkey=str(data.get("hotkey") or DEFAULT_HOTKEY),
            capture_mode=_valid_capture_mode(data.get("capture_mode")),
            capture_monitor=int(data.get("capture_monitor") or 0),
            output_mode=str(output.get("mode") or "per_module"),
            template=output.get("template"),
            created_at=str(data.get("created_at", "")),
            updated_at=str(data.get("updated_at", "")),
            config=config or load_config(search_from=folder),
        )
        project._ensure_dirs()
        if parse:
            project.load_scenarios()
            project._after_open()
        return project

    def _after_open(self) -> None:
        """Detect scenario edits made while the project was closed.

        Backs up project.json first when the workbook changed, then runs the
        same reconcile as a re-import (orphans, parked step results, review
        flags) and keeps the report in ``last_changes``.
        """
        dirty = False
        if self.excel_changed:
            self._backup()
        for uid in self._evidence_uids():
            if uid not in self.snapshots and self.test_case(uid) is not None:
                self._touch_snapshot(uid)
                dirty = True
        report = ReimportReport()
        before = (self.total_screenshots(), self.orphan_count(), self.results_count(),
                  len(self.orphan_step_results), dict(self.needs_review))
        self._reconcile(report, self.analyze())
        after = (self.total_screenshots(), self.orphan_count(), self.results_count(),
                 len(self.orphan_step_results), dict(self.needs_review))
        if dirty or self.excel_changed or before != after:
            self.save()

    # -- persistence ------------------------------------------------------- #

    def to_dict(self) -> Dict[str, Any]:
        try:
            excel = str(self.excel_path.relative_to(self.folder))
        except ValueError:
            excel = str(self.excel_path)
        evidence = set(self._evidence_uids())
        return {
            "schema_version": SCHEMA_VERSION,
            "app_version": __version__,
            "name": self.name,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "excel_path": excel,
            "selected_sheets": list(self.selected_sheets),
            "column_overrides": {
                name: override.to_dict()
                for name, override in self.column_overrides.items()
            },
            "metadata": dict(self.metadata),
            "module_titles": dict(self.module_titles),
            "screenshots": {
                uid: [entry.to_dict() for entry in entries]
                for uid, entries in self.screenshots.items()
                if entries
            },
            "orphan_screenshots": {
                uid: [entry.to_dict() for entry in entries]
                for uid, entries in self.orphans.items()
                if entries
            },
            "results": {
                uid: {key: dict(values) for key, values in steps.items() if values}
                for uid, steps in self.results.items()
                if any(steps.values())
            },
            "orphan_step_results": {
                uid: {key: dict(values) for key, values in steps.items() if values}
                for uid, steps in self.orphan_step_results.items()
                if any(steps.values())
            },
            "snapshots": {
                uid: snapshot.to_dict()
                for uid, snapshot in self.snapshots.items()
                if uid in evidence
            },
            "needs_review": {
                uid: list(lines)
                for uid, lines in self.needs_review.items()
                if lines and uid in evidence
            },
            "excel_sha256": self.excel_sha256,
            "hotkey": self.hotkey,
            "capture_mode": self.capture_mode,
            "capture_monitor": self.capture_monitor,
            "output": {"mode": self.output_mode, "template": self.template},
        }

    def save(self) -> Path:
        """Write project.json atomically, so a crash cannot truncate it."""
        self._ensure_dirs()
        self.updated_at = _now()
        if not self.created_at:
            self.created_at = self.updated_at
        payload = json.dumps(self.to_dict(), indent=2, ensure_ascii=False)
        temp = self.project_file.with_name(PROJECT_FILENAME + ".tmp")
        temp.write_text(payload, encoding="utf-8")
        temp.replace(self.project_file)
        return self.project_file

    # -- scenarios --------------------------------------------------------- #

    def overrides_dict(self) -> Dict[str, Any]:
        return {name: ov.to_dict() for name, ov in self.column_overrides.items()}

    def inspect(self) -> WorkbookInspection:
        """Inspect the workbook for the sheet picker."""
        return inspect_workbook(
            self.excel_path, config=self.config, overrides=self.overrides_dict()
        )

    def load_scenarios(self) -> ScenarioFile:
        """(Re-)read the selected sheets from the Excel file."""
        scenario_file = parse_workbook(
            self.excel_path,
            self.selected_sheets or None,
            config=self.config,
            overrides=self.overrides_dict(),
        )
        self.scenario_file = scenario_file
        digest = _file_sha256(self.excel_path)
        self.excel_changed = bool(self.excel_sha256) and digest != self.excel_sha256
        self.excel_sha256 = digest
        self.warnings = list(scenario_file.warnings)
        if not self.selected_sheets:
            self.selected_sheets = list(scenario_file.module_names)
        return scenario_file

    @property
    def test_cases(self) -> List[TestCase]:
        return self.scenario_file.test_cases if self.scenario_file else []

    @property
    def known_uids(self) -> List[str]:
        return [case.uid for case in self.test_cases]

    def test_case(self, uid: str) -> Optional[TestCase]:
        return self.scenario_file.test_case(uid) if self.scenario_file else None

    def effective_metadata(self) -> Dict[str, str]:
        """Parsed metadata with the user's manual edits on top."""
        merged: Dict[str, str] = {}
        if self.scenario_file is not None:
            merged.update(self.scenario_file.merged_metadata())
        merged.update({k: v for k, v in self.metadata.items() if v})
        return merged

    # -- re-import --------------------------------------------------------- #

    def reimport(
        self,
        *,
        excel_path: Optional[Path] = None,
        sheets: Optional[Iterable[str]] = None,
        overrides: Optional[Mapping[str, Any]] = None,
    ) -> ReimportReport:
        """Re-read the Excel file, keeping the screenshots already taken.

        Screenshots of test cases that no longer exist become orphans instead
        of being deleted; orphans whose test case reappears are restored.
        """
        before_uids = set(self.known_uids)
        before_modules = set(self.scenario_file.module_names) if self.scenario_file else set()

        if excel_path is not None:
            excel_path = Path(excel_path)
            if not excel_path.is_file():
                raise ProjectError(f"File Excel tidak ditemukan: {excel_path}")
            self.excel_path = excel_path
        if sheets is not None:
            self.selected_sheets = list(sheets)
        if overrides is not None:
            self.column_overrides = {
                str(key): SheetOverride.coerce(value)
                for key, value in overrides.items()
            }

        self.save()
        self._backup()
        scenario_file = self.load_scenarios()
        after_uids = set(case.uid for case in scenario_file.test_cases)
        after_modules = set(scenario_file.module_names)

        report = ReimportReport(
            added_modules=sorted(after_modules - before_modules),
            removed_modules=sorted(before_modules - after_modules),
            added_test_cases=sorted(after_uids - before_uids),
            removed_test_cases=sorted(before_uids - after_uids),
            warnings=list(scenario_file.warnings),
        )

        # Compare with the snapshots *before* anything is moved.
        report.changes = self.analyze()
        self._reconcile(report, report.changes)

        # Drop module titles of modules that are gone.
        for name in report.removed_modules:
            self.module_titles.pop(name, None)

        self.save()
        return report

    def _reconcile(self, report: ReimportReport, changes: ScenarioChanges) -> None:
        """Park evidence that lost its test case / step, restore what came back."""
        after_uids = set(self.known_uids)

        # Screenshots of vanished test cases -> orphan.
        for uid in list(self.screenshots):
            if uid in after_uids or not self.screenshots[uid]:
                continue
            entries = self.screenshots.pop(uid)
            self.orphans.setdefault(uid, []).extend(entries)
            report.orphaned[uid] = len(entries)

        # Orphans whose test case came back -> restored.
        for uid in list(self.orphans):
            if uid not in after_uids or not self.orphans[uid]:
                continue
            entries = self.orphans.pop(uid)
            self.screenshots.setdefault(uid, []).extend(entries)
            report.restored[uid] = len(entries)

        # Results of steps that no longer exist are kept aside, not dropped.
        # Test cases waiting for a remap decision are left alone.
        pending = {c.uid for c in changes.changes if c.status == SUSPECT}
        for uid, steps in list(self.results.items()):
            case = self.test_case(uid)
            if case is None or uid in pending:
                continue
            keys = {step_key(step, i) for i, step in enumerate(case.steps, start=1)}
            for key in [k for k in steps if k not in keys]:
                sibling = self._sibling_with_step(case, key)
                if sibling is not None:
                    # A project saved while a duplicated Test Case ID was still
                    # merged: the step now belongs to the other test case.
                    self.results.setdefault(sibling.uid, {})[key] = steps.pop(key)
                else:
                    self.orphan_step_results.setdefault(uid, {})[key] = steps.pop(key)
            if not steps:
                self.results.pop(uid, None)

        # Step results come back when the step does.
        for uid, steps in list(self.orphan_step_results.items()):
            case = self.test_case(uid)
            if case is None:
                continue
            keys = {step_key(step, i) for i, step in enumerate(case.steps, start=1)}
            for key in [k for k in steps if k in keys]:
                self.results.setdefault(uid, {}).setdefault(key, steps.pop(key))
            if not steps:
                self.orphan_step_results.pop(uid, None)

        self.needs_review = {
            change.uid: list(change.details)
            for change in changes.changes
            if change.status in (EDITED, SUSPECT)
        }

    def _sibling_with_step(self, case: TestCase, key: str) -> Optional[TestCase]:
        """Another test case with the same ID that owns step ``key`` and has no result yet.

        Positional keys (``#2``) exist in every case, so they are never moved.
        """
        if key.startswith("#"):
            return None
        for other in self.test_cases:
            if (other.uid == case.uid or other.module != case.module
                    or other.test_case_id != case.test_case_id):
                continue
            has_step = any(
                step_key(step, i) == key for i, step in enumerate(other.steps, start=1)
            )
            if has_step and key not in self.results.get(other.uid, {}):
                return other
        return None

    # -- change detection --------------------------------------------------- #

    def _evidence_uids(self) -> List[str]:
        """Test cases that hold screenshots, orphans or results."""
        uids: List[str] = []
        for source in (self.screenshots, self.orphans, self.results):
            for uid, items in source.items():
                if any(items.values()) if isinstance(items, dict) else items:
                    if uid not in uids:
                        uids.append(uid)
        return uids

    def _touch_snapshot(self, uid: str) -> None:
        """Remember what the test case looked like when evidence was first stored."""
        case = self.test_case(uid)
        if case is not None and uid not in self.snapshots:
            self.snapshots[uid] = snapshot_of(case, taken_at=_now())

    def analyze(self) -> ScenarioChanges:
        """Compare the stored snapshots with the scenario that is loaded now."""
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        changes = diff_scenarios(
            self.snapshots, self._evidence_uids(), self.scenario_file, self.results
        )
        self.last_changes = changes
        return changes

    def check_changes(self) -> ScenarioChanges:
        """Re-run the comparison without moving anything (Review menu)."""
        changes = self.analyze()
        self.needs_review = {
            change.uid: list(change.details)
            for change in changes.changes
            if change.status in (EDITED, SUSPECT)
        }
        return changes

    def acknowledge(self, uid: str, *, save: bool = True) -> None:
        """The user checked this test case: take a fresh snapshot, clear the flag."""
        self.needs_review.pop(uid, None)
        self.snapshots.pop(uid, None)
        self._touch_snapshot(uid)
        if self.last_changes is not None:
            self.last_changes.changes = [
                c for c in self.last_changes.changes if c.uid != uid
            ]
        if save:
            self.save()

    def apply_remap(
        self,
        mapping: Mapping[str, str],
        step_maps: Optional[Mapping[str, Mapping[str, str]]] = None,
        *,
        save: bool = True,
    ) -> Dict[str, int]:
        """Move the evidence of ``source`` test cases onto ``target`` ones.

        Works as one batch so chains (2.1 -> 2.2, 2.2 -> 2.3) and swaps are
        safe. Evidence already on a target is kept; the moved one is appended.
        Step results are renamed with ``step_maps``; steps without a partner
        go to ``orphan_step_results``. Files on disk are not touched.
        Returns ``{source uid: screenshots moved}``.
        """
        step_maps = step_maps or {}
        pairs = {src: dst for src, dst in mapping.items() if src != dst}
        for src, dst in pairs.items():
            if self.test_case(dst) is None:
                raise ProjectError(f"Test case tujuan tidak ada: {dst}")

        detached = {}
        for src in pairs:
            shots = self.screenshots.pop(src, []) + self.orphans.pop(src, [])
            results = dict(self.results.pop(src, {}))
            results.update(self.orphan_step_results.pop(src, {}))
            detached[src] = (shots, results)

        moved: Dict[str, int] = {}
        for src, dst in pairs.items():
            shots, results = detached[src]
            if shots:
                self.screenshots.setdefault(dst, []).extend(shots)
            step_map = step_maps.get(src, {})
            for key, values in results.items():
                new_key = step_map.get(key)
                target = self.results.setdefault(dst, {})
                if new_key is not None and new_key not in target:
                    target[new_key] = values
                else:
                    self.orphan_step_results.setdefault(dst, {})[key] = values
            if dst in self.results and not self.results[dst]:
                self.results.pop(dst)
            moved[src] = len(shots)

        for src, dst in pairs.items():
            for uid in (src, dst):
                self.needs_review.pop(uid, None)
                self.snapshots.pop(uid, None)
        for dst in pairs.values():
            self._touch_snapshot(dst)
        if self.last_changes is not None:
            handled = set(pairs) | set(pairs.values())
            self.last_changes.changes = [
                c for c in self.last_changes.changes if c.uid not in handled
            ]
        if save:
            self.save()
        return moved

    # -- backups ------------------------------------------------------------ #

    def _backup(self) -> Optional[Path]:
        """Copy project.json to backups/ so a re-import can be undone."""
        if not self.project_file.is_file():
            return None
        self.backups_dir.mkdir(exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        target = self.backups_dir / f"project-{stamp}.json"
        shutil.copy2(self.project_file, target)
        for old in self.backups()[:-MAX_BACKUPS]:
            try:
                old.unlink()
            except OSError:
                pass
        return target

    def backups(self) -> List[Path]:
        """Backups of project.json, oldest first."""
        if not self.backups_dir.is_dir():
            return []
        return sorted(self.backups_dir.glob("project-*.json"))

    def latest_backup(self) -> Optional[Path]:
        found = self.backups()
        return found[-1] if found else None

    def restore_backup(self, path: Optional[Path] = None) -> Path:
        """Put a backup back as project.json and reload it into this object."""
        path = Path(path) if path else self.latest_backup()
        if path is None or not path.is_file():
            raise ProjectError("Tidak ada backup project.json untuk dikembalikan.")
        shutil.copy2(path, self.project_file)
        fresh = Project.open(self.folder, config=self.config, parse=False)
        self.__dict__.update(fresh.__dict__)
        self.load_scenarios()
        self.last_changes = None
        return path

    # -- screenshot mapping ------------------------------------------------ #

    def entries(self, uid: str) -> List[ScreenshotEntry]:
        return self.screenshots.get(uid, [])

    def screenshot_count(self, uid: str) -> int:
        return len(self.screenshots.get(uid, []))

    def total_screenshots(self) -> int:
        return sum(len(items) for items in self.screenshots.values())

    @property
    def orphan_uids(self) -> List[str]:
        return [uid for uid, items in self.orphans.items() if items]

    def orphan_count(self) -> int:
        return sum(len(items) for items in self.orphans.values())

    def target_dir(self, uid: str) -> Path:
        """``screenshots/<module-slug>/<TestCaseID>_<test-case-slug>``."""
        case = self.test_case(uid)
        if case is None:
            raise ProjectError(
                f"Test case '{uid}' tidak ada di skenario yang sedang dibuka."
            )
        return (
            self.screenshots_dir
            / slugify(case.module, fallback="module")
            / _safe_component(case.folder_name)
        )

    def next_sequence(self, uid: str) -> int:
        return len(self.screenshots.get(uid, [])) + 1

    def add_screenshot(
        self,
        uid: str,
        image: Path,
        *,
        caption: str = "",
        move: bool = False,
        save: bool = True,
    ) -> ScreenshotEntry:
        """Store an image file as the next screenshot of ``uid``.

        File name: ``<TestCaseID>_<urutan>_<timestamp>.png``
        """
        case = self.test_case(uid)
        if case is None:
            raise ProjectError(
                f"Test case '{uid}' tidak ada di skenario yang sedang dibuka."
            )
        image = Path(image)
        if not image.is_file():
            raise ProjectError(f"File gambar tidak ditemukan: {image}")

        directory = self.target_dir(uid)
        directory.mkdir(parents=True, exist_ok=True)
        suffix = image.suffix.lower() or ".png"
        base = _safe_component(case.test_case_id, fallback="tc")
        target = directory / f"{base}_{self.next_sequence(uid):03d}_{_timestamp()}{suffix}"
        counter = 1
        while target.exists():  # same second, same test case
            counter += 1
            target = target.with_name(
                f"{target.stem}-{counter}{target.suffix}"
            )
        if move:
            shutil.move(str(image), target)
        else:
            shutil.copy2(image, target)

        entry = ScreenshotEntry(
            file=str(target.relative_to(self.folder)).replace("\\", "/"),
            caption=caption,
            captured_at=_now(),
        )
        self.screenshots.setdefault(uid, []).append(entry)
        self._touch_snapshot(uid)
        if save:
            self.save()
        return entry

    def set_caption(self, uid: str, index: int, caption: str, *, save: bool = True) -> None:
        self.screenshots[uid][index].caption = caption
        if save:
            self.save()

    def reorder_screenshot(
        self, uid: str, from_index: int, to_index: int, *, save: bool = True
    ) -> None:
        """Move one screenshot inside its test case (drag and drop)."""
        entries = self.screenshots.get(uid)
        if not entries:
            raise ProjectError(f"Test case '{uid}' belum punya screenshot.")
        if not 0 <= from_index < len(entries):
            raise ProjectError(f"Indeks screenshot di luar jangkauan: {from_index}")
        entry = entries.pop(from_index)
        entries.insert(max(0, min(to_index, len(entries))), entry)
        if save:
            self.save()

    def move_screenshot(
        self, uid: str, index: int, target_uid: str, *, save: bool = True
    ) -> ScreenshotEntry:
        """Move a screenshot to another test case, file included."""
        entries = self.screenshots.get(uid)
        if not entries or not 0 <= index < len(entries):
            raise ProjectError(f"Screenshot tidak ditemukan: {uid}[{index}]")
        if self.test_case(target_uid) is None:
            raise ProjectError(f"Test case tujuan tidak ada: {target_uid}")

        entry = entries[index]
        source = self.absolute(entry)
        directory = self.target_dir(target_uid)
        directory.mkdir(parents=True, exist_ok=True)
        target_case = self.test_case(target_uid)
        assert target_case is not None
        base = _safe_component(target_case.test_case_id, fallback="tc")
        target = directory / (
            f"{base}_{self.next_sequence(target_uid):03d}_{_timestamp()}{source.suffix}"
        )
        counter = 1
        while target.exists():
            counter += 1
            target = target.with_name(f"{target.stem}-{counter}{target.suffix}")

        if source.is_file():
            shutil.move(str(source), target)
            entry.file = str(target.relative_to(self.folder)).replace("\\", "/")
        entries.pop(index)
        if not entries:
            self.screenshots.pop(uid, None)
        self.screenshots.setdefault(target_uid, []).append(entry)
        self._touch_snapshot(target_uid)
        if save:
            self.save()
        return entry

    def delete_screenshot(
        self, uid: str, index: int, *, delete_file: bool = True, save: bool = True
    ) -> ScreenshotEntry:
        entries = self.screenshots.get(uid)
        if not entries or not 0 <= index < len(entries):
            raise ProjectError(f"Screenshot tidak ditemukan: {uid}[{index}]")
        entry = entries.pop(index)
        if not entries:
            self.screenshots.pop(uid, None)
        if delete_file:
            path = self.absolute(entry)
            if path.is_file():
                path.unlink()
        if save:
            self.save()
        return entry

    def restore_orphan(
        self, uid: str, target_uid: str, *, save: bool = True
    ) -> List[ScreenshotEntry]:
        """Re-attach every orphan screenshot of ``uid`` to another test case."""
        entries = self.orphans.get(uid)
        if not entries:
            raise ProjectError(f"Tidak ada screenshot orphan untuk '{uid}'.")
        if self.test_case(target_uid) is None:
            raise ProjectError(f"Test case tujuan tidak ada: {target_uid}")
        self.orphans.pop(uid)
        restored: List[ScreenshotEntry] = []
        for entry in entries:
            self.screenshots.setdefault(target_uid, []).append(entry)
            restored.append(entry)
        self._touch_snapshot(target_uid)
        if save:
            self.save()
        return restored

    def discard_orphans(self, uid: str, *, delete_files: bool = False,
                        save: bool = True) -> int:
        entries = self.orphans.pop(uid, [])
        if delete_files:
            for entry in entries:
                path = self.absolute(entry)
                if path.is_file():
                    path.unlink()
        if save:
            self.save()
        return len(entries)

    def missing_files(self) -> List[Tuple[str, ScreenshotEntry]]:
        """Screenshots recorded in project.json whose file is gone."""
        missing: List[Tuple[str, ScreenshotEntry]] = []
        for uid, entries in self.screenshots.items():
            for entry in entries:
                if not self.absolute(entry).is_file():
                    missing.append((uid, entry))
        return missing

    # -- execution results -------------------------------------------------- #

    @property
    def results_dir(self) -> Path:
        return self.folder / RESULTS_DIRNAME

    def result_for(self, uid: str, key: str) -> Dict[str, str]:
        return dict(self.results.get(uid, {}).get(key, {}))

    def set_result(
        self, uid: str, key: str, values: Mapping[str, Any], *, save: bool = True
    ) -> Dict[str, str]:
        """Record Actual Results / Pass-Fail / Date / Executed by for one step.

        An empty value clears that field rather than storing a blank, so the
        write-back never overwrites what is already in the workbook with "".
        """
        unknown = [name for name in values if name not in RESULT_FIELDS]
        if unknown:
            raise ProjectError(
                "Field hasil tidak dikenali: " + ", ".join(sorted(unknown))
                + ". Yang tersedia: " + ", ".join(RESULT_FIELDS) + "."
            )
        entry = self.results.setdefault(uid, {}).setdefault(key, {})
        for name in RESULT_FIELDS:
            if name not in values:
                continue
            text = "" if values[name] is None else str(values[name]).strip()
            if text:
                entry[name] = text
            else:
                entry.pop(name, None)
        if not entry:
            self.results[uid].pop(key, None)
            if not self.results[uid]:
                self.results.pop(uid, None)
        else:
            self._touch_snapshot(uid)
        if save:
            self.save()
        return dict(entry)

    def clear_result(self, uid: str, key: str, *, save: bool = True) -> None:
        if uid in self.results:
            self.results[uid].pop(key, None)
            if not self.results[uid]:
                self.results.pop(uid, None)
        if save:
            self.save()

    def set_results_for_test_case(
        self, uid: str, values: Mapping[str, Any], *, save: bool = True
    ) -> int:
        """Apply the same values to every step of a test case."""
        case = self.test_case(uid)
        if case is None:
            raise ProjectError(f"Test case '{uid}' tidak ada di skenario yang dibuka.")
        for index, step in enumerate(case.steps, start=1):
            self.set_result(uid, step_key(step, index), values, save=False)
        if save:
            self.save()
        return len(case.steps)

    def results_count(self) -> int:
        return sum(len(steps) for steps in self.results.values())

    def steps_with_results(self, uid: str) -> int:
        return len(self.results.get(uid, {}))

    def orphan_results(self) -> List[str]:
        """Test cases that have stored results but no longer exist in the Excel."""
        known = set(self.known_uids)
        return sorted(uid for uid, steps in self.results.items() if steps and uid not in known)

    def result_updates(self) -> Dict[str, List[RowUpdate]]:
        """Turn the stored results into row updates for the current workbook.

        The row numbers come from the freshly parsed scenario, so a re-import
        that moved rows around is picked up automatically.
        """
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        updates: Dict[str, List[RowUpdate]] = {}
        for module in self.scenario_file.modules:
            rows: List[RowUpdate] = []
            for case in module.test_cases:
                stored = self.results.get(case.uid)
                if not stored:
                    continue
                for index, step in enumerate(case.steps, start=1):
                    values = stored.get(step_key(step, index))
                    if values:
                        rows.append(RowUpdate(row=step.row, values=dict(values)))
            if rows:
                updates[module.name] = rows
        return updates

    def write_result_workbook(
        self,
        destination: Optional[Path] = None,
        *,
        overwrite: bool = False,
    ) -> WriteReport:
        """Write the results into a copy of the Excel file. Never the original."""
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        updates = self.result_updates()
        if not updates:
            raise ProjectError(
                "Belum ada hasil eksekusi yang disimpan, jadi tidak ada yang "
                "ditulis ke Excel."
            )
        if destination is None:
            self.results_dir.mkdir(parents=True, exist_ok=True)
            destination = self.results_dir / default_result_name(self.excel_path)
            overwrite = True
        columns = {
            module.name: module.column_map for module in self.scenario_file.modules
        }
        return write_results(
            self.excel_path, Path(destination), updates, columns, overwrite=overwrite
        )

    # -- export ------------------------------------------------------------ #

    def export_screenshots(self) -> Dict[str, List[Screenshot]]:
        """The screenshot mapping in the shape docx_builder expects."""
        return {
            uid: [
                Screenshot(path=self.absolute(entry), caption=entry.caption)
                for entry in entries
            ]
            for uid, entries in self.screenshots.items()
            if entries
        }

    def naming(self):
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        return resolve_naming(self.scenario_file, titles=self.module_titles)

    def summarize(self, *, modules: Optional[Iterable[str]] = None) -> List[ModuleReport]:
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        return summarize(
            self.scenario_file,
            self.export_screenshots(),
            modules=modules,
            titles=self.module_titles,
        )

    def generate(
        self,
        *,
        modules: Optional[Iterable[str]] = None,
        combined: Optional[bool] = None,
        output_dir: Optional[Path] = None,
    ) -> List[ExportResult]:
        """Write the Word document(s) into the project's output folder."""
        if self.scenario_file is None:
            raise ProjectError("Skenario belum dibaca; jalankan load_scenarios().")
        template = self.template_path
        if template is not None and not template.is_file():
            raise ProjectError(f"File template tidak ditemukan: {template}")
        if combined is None:
            combined = self.output_mode == "combined"
        return export_documents(
            self.scenario_file,
            self.export_screenshots(),
            output_dir=Path(output_dir) if output_dir else self.output_dir,
            combined=combined,
            modules=modules,
            titles=self.module_titles,
            template=template,
            project_name=self.effective_metadata().get("project_name"),
        )
