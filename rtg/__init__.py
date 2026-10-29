"""Result Test Generator - build UAT result documents from scenario workbooks."""

from __future__ import annotations

__version__ = "0.1.0"
__app_name__ = "Result Test Generator"

from .capture import (
    CAPTURE_MODES,
    MODE_FULLSCREEN,
    MODE_MONITOR,
    MODE_REGION,
    CaptureError,
    Monitor,
    capture_to_file,
    crop,
    grab,
    list_monitors,
    save_png,
)
from .config import FieldConfig, load_config, write_default_config
from .errors import (
    CorruptWorkbookError,
    FileLockedError,
    MissingRequiredColumnsError,
    NoScenarioSheetError,
    ResultGeneratorError,
    ScenarioFileNotFoundError,
    SheetNotFoundError,
    UnsupportedFileFormatError,
)
from .docx_builder import (
    ExportResult,
    ModuleNaming,
    ModuleReport,
    Screenshot,
    build_document,
    export,
    export_combined,
    export_per_module,
    resolve_naming,
    summarize,
)
from .excel_parser import (
    SheetInspection,
    SheetOverride,
    WorkbookInspection,
    inspect_workbook,
    parse_workbook,
)
from .hotkey import HotkeyError, HotkeyListener, format_hotkey, parse_hotkey
from .models import Module, ScenarioFile, ScenarioGroup, TestCase, TestStep
from .project import Project, ProjectError, ReimportReport, ScreenshotEntry

__all__ = [
    "__version__",
    "__app_name__",
    "FieldConfig",
    "load_config",
    "write_default_config",
    "ResultGeneratorError",
    "ScenarioFileNotFoundError",
    "UnsupportedFileFormatError",
    "FileLockedError",
    "CorruptWorkbookError",
    "NoScenarioSheetError",
    "MissingRequiredColumnsError",
    "SheetNotFoundError",
    "ExportResult",
    "ModuleNaming",
    "ModuleReport",
    "Screenshot",
    "build_document",
    "export",
    "export_combined",
    "export_per_module",
    "resolve_naming",
    "summarize",
    "SheetInspection",
    "SheetOverride",
    "WorkbookInspection",
    "inspect_workbook",
    "parse_workbook",
    "CAPTURE_MODES",
    "MODE_FULLSCREEN",
    "MODE_MONITOR",
    "MODE_REGION",
    "CaptureError",
    "Monitor",
    "capture_to_file",
    "crop",
    "grab",
    "list_monitors",
    "save_png",
    "HotkeyError",
    "HotkeyListener",
    "format_hotkey",
    "parse_hotkey",
    "Project",
    "ProjectError",
    "ReimportReport",
    "ScreenshotEntry",
    "Module",
    "ScenarioFile",
    "ScenarioGroup",
    "TestCase",
    "TestStep",
]
