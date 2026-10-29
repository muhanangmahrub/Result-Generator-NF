"""Exceptions raised by the Result Test Generator core.

Every exception carries a `message` written in Indonesian so the GUI can show it
to the user verbatim, plus an optional `detail` with technical context for logs.
"""

from __future__ import annotations

from typing import Any, List, Optional, Sequence


class ResultGeneratorError(Exception):
    """Base class for all errors raised by this application."""

    def __init__(self, message: str, *, detail: Optional[str] = None) -> None:
        super().__init__(message)
        self.message = message
        self.detail = detail

    def __str__(self) -> str:  # pragma: no cover - trivial
        if self.detail:
            return f"{self.message} ({self.detail})"
        return self.message


class ScenarioFileNotFoundError(ResultGeneratorError):
    """The scenario workbook does not exist."""


class UnsupportedFileFormatError(ResultGeneratorError):
    """The file is not an .xlsx/.xlsm workbook (e.g. legacy .xls or .csv)."""


class FileLockedError(ResultGeneratorError):
    """The workbook cannot be read because another process holds a lock."""


class CorruptWorkbookError(ResultGeneratorError):
    """The file has an .xlsx extension but cannot be opened as a workbook."""


class NoScenarioSheetError(ResultGeneratorError):
    """No sheet in the workbook contains the mandatory scenario columns."""

    def __init__(self, message: str, *, sheets: Optional[Sequence[Any]] = None,
                 detail: Optional[str] = None) -> None:
        super().__init__(message, detail=detail)
        self.sheets: List[Any] = list(sheets or [])


class MissingRequiredColumnsError(ResultGeneratorError):
    """A sheet the user explicitly selected is missing mandatory columns.

    The GUI catches this to open the manual column-mapping dialog, so the
    detected headers travel with the exception.
    """

    def __init__(self, message: str, *, sheet_name: str,
                 missing: Sequence[str], headers: Optional[Sequence[Any]] = None,
                 header_row: Optional[int] = None,
                 detail: Optional[str] = None) -> None:
        super().__init__(message, detail=detail)
        self.sheet_name = sheet_name
        self.missing: List[str] = list(missing)
        self.headers: List[Any] = list(headers or [])
        self.header_row = header_row


class SheetNotFoundError(ResultGeneratorError):
    """A sheet name stored in project.json no longer exists in the workbook."""
