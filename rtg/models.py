"""Data model for a parsed scenario workbook.

Hierarchy: ScenarioFile -> Module (one Excel sheet) -> ScenarioGroup -> TestCase
-> TestStep.

The unique identity of a test case is ``module + test_case_id`` (see
:attr:`TestCase.uid`): the same Test Case ID legitimately appears in several
sheets, so the sheet name must be part of the key used for screenshot mapping.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple

from .normalize import slugify

UID_SEPARATOR = "::"

#: Appended to the uid of the 2nd, 3rd... test case that reuses a Test Case ID.
OCCURRENCE_SEPARATOR = "~"

#: Separator between Test Case ID and Test Case in a Word Heading 2. The
#: reference result document uses a tab, which is also what Word's heading
#: styles are set up for.
HEADING_SEPARATOR = "\t"

#: Optional per-step fields, in the order they appear in the scenario template.
STEP_FIELDS: Tuple[str, ...] = (
    "test_step_id",
    "test_step",
    "positive_negative",
    "pre_condition",
    "test_data",
    "expected_results",
    "actual_results",
    "pass_fail",
    "date_executed",
    "test_executed_by",
    "notes",
)


@dataclass
class TestStep:
    """One row of the scenario table."""

    row: int
    test_step: str = ""
    test_step_id: str = ""
    positive_negative: str = ""
    pre_condition: str = ""
    test_data: str = ""
    expected_results: str = ""
    actual_results: str = ""
    pass_fail: str = ""
    date_executed: str = ""
    test_executed_by: str = ""
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"row": self.row}
        for name in STEP_FIELDS:
            value = getattr(self, name)
            if value:
                data[name] = value
        return data


@dataclass
class TestCase:
    """A test case: the unit a screenshot is attached to."""

    module: str
    test_case_id: str
    test_case: str = ""
    scenario_id: str = ""
    test_scenario: str = ""
    test_scenario_description: str = ""
    row: int = 0
    steps: List[TestStep] = field(default_factory=list)
    #: 1 for the first test case with this ID in the sheet, 2 for the next...
    occurrence: int = 1

    @property
    def uid(self) -> str:
        """Stable identity across re-imports: ``<module>::<test case id>``.

        A repeated Test Case ID keeps its first uid unchanged; later ones get
        ``~2``, ``~3``... so screenshots and results never collide.
        """
        base = f"{self.module}{UID_SEPARATOR}{self.test_case_id}"
        if self.occurrence > 1:
            return f"{base}{OCCURRENCE_SEPARATOR}{self.occurrence}"
        return base

    @property
    def heading(self) -> str:
        """Word Heading 2 text: ``<Test Case ID><TAB><Test Case>``."""
        if self.test_case and self.test_case_id:
            return f"{self.test_case_id}{HEADING_SEPARATOR}{self.test_case}"
        return self.test_case_id or self.test_case

    @property
    def display_label(self) -> str:
        """Same text as :attr:`heading` but for the GUI tree and the CLI."""
        label = self.heading.replace(HEADING_SEPARATOR, "  ")
        if self.occurrence > 1:
            label += f"  [ID ganda #{self.occurrence}]"
        return label

    @property
    def folder_name(self) -> str:
        """``<TestCaseID>_<test-case-slug>`` - the screenshot subfolder."""
        if self.occurrence > 1:
            return f"{self.test_case_id}-{self.occurrence}_{slugify(self.test_case)}"
        return f"{self.test_case_id}_{slugify(self.test_case)}"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "uid": self.uid,
            "module": self.module,
            "test_case_id": self.test_case_id,
            "test_case": self.test_case,
            "scenario_id": self.scenario_id,
            "test_scenario": self.test_scenario,
            "test_scenario_description": self.test_scenario_description,
            "row": self.row,
            "occurrence": self.occurrence,
            "steps": [step.to_dict() for step in self.steps],
        }


@dataclass
class ScenarioGroup:
    """Test cases that share a Scenario ID / Test Scenario."""

    scenario_id: str = ""
    test_scenario: str = ""
    test_scenario_description: str = ""
    test_cases: List[TestCase] = field(default_factory=list)

    @property
    def title(self) -> str:
        if self.scenario_id and self.test_scenario:
            return f"{self.scenario_id}  {self.test_scenario}"
        return self.test_scenario or self.scenario_id or "(tanpa skenario)"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "scenario_id": self.scenario_id,
            "test_scenario": self.test_scenario,
            "test_scenario_description": self.test_scenario_description,
            "test_cases": [tc.to_dict() for tc in self.test_cases],
        }


@dataclass
class Module:
    """One scenario sheet. The module name is always the sheet name."""

    name: str
    sheet_index: int = 0
    header_row: int = 0
    column_map: Dict[str, int] = field(default_factory=dict)
    metadata: Dict[str, str] = field(default_factory=dict)
    scenarios: List[ScenarioGroup] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    unmapped_columns: List[Tuple[int, str]] = field(default_factory=list)

    @property
    def slug(self) -> str:
        return slugify(self.name, fallback="module")

    @property
    def title(self) -> str:
        """Default Heading 1 text: sheet name in upper case."""
        return self.name.upper()

    @property
    def test_cases(self) -> List[TestCase]:
        return [tc for group in self.scenarios for tc in group.test_cases]

    @property
    def step_count(self) -> int:
        return sum(len(tc.steps) for tc in self.test_cases)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "sheet_index": self.sheet_index,
            "header_row": self.header_row,
            "column_map": dict(self.column_map),
            "metadata": dict(self.metadata),
            "scenarios": [group.to_dict() for group in self.scenarios],
            "warnings": list(self.warnings),
            "unmapped_columns": [list(item) for item in self.unmapped_columns],
        }


@dataclass
class ScenarioFile:
    """The result of parsing one workbook."""

    path: Path
    modules: List[Module] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def test_cases(self) -> List[TestCase]:
        return [tc for module in self.modules for tc in module.test_cases]

    @property
    def module_names(self) -> List[str]:
        return [module.name for module in self.modules]

    def module(self, name: str) -> Optional[Module]:
        for module in self.modules:
            if module.name == name:
                return module
        return None

    def test_case(self, uid: str) -> Optional[TestCase]:
        for case in self.test_cases:
            if case.uid == uid:
                return case
        return None

    def iter_test_cases(self) -> Iterator[Tuple[Module, ScenarioGroup, TestCase]]:
        for module in self.modules:
            for group in module.scenarios:
                for case in group.test_cases:
                    yield module, group, case

    def merged_metadata(self) -> Dict[str, str]:
        """Metadata from all sheets; the first sheet that defines a key wins.

        Lets a module without its own metadata block fall back to another
        sheet's values, as required by the spec.
        """
        merged: Dict[str, str] = {}
        for module in self.modules:
            for key, value in module.metadata.items():
                if value and not merged.get(key):
                    merged[key] = value
        return merged

    def to_dict(self) -> Dict[str, Any]:
        return {
            "path": str(self.path),
            "modules": [module.to_dict() for module in self.modules],
            "warnings": list(self.warnings),
        }
