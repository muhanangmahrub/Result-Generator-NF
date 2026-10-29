"""Convert the generated .docx into PDF.

Two converters, tried in that order unless one is named explicitly:

* **Microsoft Word** through COM automation. Best fidelity, because the file is
  rendered by the same program the document was written for. Needs ``pywin32``
  and Word installed.
* **LibreOffice** (``soffice --headless``). Works without Word, but the layout
  can drift slightly.

Nothing else in the application depends on this module, so a machine without
either converter still produces .docx files normally.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence

from .errors import ResultGeneratorError

WORD = "word"
LIBREOFFICE = "libreoffice"
CONVERTERS = (WORD, LIBREOFFICE)

CONVERTER_LABELS = {
    WORD: "Microsoft Word",
    LIBREOFFICE: "LibreOffice",
}

#: Word's wdFormatPDF.
_WD_FORMAT_PDF = 17
_SOFFICE_TIMEOUT = 180

#: Where LibreOffice usually lives on Windows when it is not on PATH.
_SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


class PdfExportError(ResultGeneratorError):
    """Raised when the PDF cannot be produced."""


@dataclass(frozen=True)
class ConverterInfo:
    name: str
    available: bool
    detail: str = ""

    @property
    def label(self) -> str:
        return CONVERTER_LABELS.get(self.name, self.name)


# --------------------------------------------------------------------------- #
# Detection
# --------------------------------------------------------------------------- #

def find_soffice() -> Optional[Path]:
    """Locate the LibreOffice executable, on PATH or in the usual places."""
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return Path(found)
    for candidate in _SOFFICE_CANDIDATES:
        path = Path(candidate)
        if path.is_file():
            return path
    return None


def word_available() -> bool:
    """True when pywin32 is installed and Word is registered for COM.

    Deliberately a registry lookup rather than starting Word: launching and
    quitting Word just to answer "is it there?" costs seconds, and doing it
    repeatedly makes the COM server fail with RPC errors.
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client  # noqa: F401,PLC0415
    except ImportError:
        return False
    try:
        import winreg  # noqa: PLC0415

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
    except OSError:
        return False
    return True


def available_converters() -> List[ConverterInfo]:
    """Report which converters this machine can use."""
    infos: List[ConverterInfo] = []

    if not sys.platform.startswith("win"):
        infos.append(ConverterInfo(WORD, False, "hanya tersedia di Windows"))
    else:
        try:
            import win32com.client  # noqa: F401,PLC0415
        except ImportError:
            infos.append(
                ConverterInfo(WORD, False, "pywin32 belum terpasang (pip install pywin32)")
            )
        else:
            ok = word_available()
            infos.append(
                ConverterInfo(
                    WORD, ok, "" if ok else "Microsoft Word tidak terpasang"
                )
            )

    soffice = find_soffice()
    infos.append(
        ConverterInfo(
            LIBREOFFICE, soffice is not None,
            str(soffice) if soffice else "soffice tidak ditemukan",
        )
    )
    return infos


def default_converter() -> Optional[str]:
    for info in available_converters():
        if info.available:
            return info.name
    return None


# --------------------------------------------------------------------------- #
# Conversion
# --------------------------------------------------------------------------- #

def _convert_with_word(source: Path, destination: Path) -> Path:
    try:
        import pythoncom  # noqa: PLC0415
        import win32com.client  # noqa: PLC0415
    except ImportError as exc:
        raise PdfExportError(
            "Export PDF lewat Word butuh pywin32. Jalankan: pip install pywin32",
            detail=str(exc),
        ) from exc

    # The GUI may call this from a worker thread, which needs its own COM
    # apartment. Only uninitialise one we actually created.
    initialised = False
    try:
        pythoncom.CoInitialize()
        initialised = True
    except Exception:  # pragma: no cover - already initialised on this thread
        pass

    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        document = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
        document.SaveAs(str(destination), FileFormat=_WD_FORMAT_PDF)
    except Exception as exc:
        raise PdfExportError(
            f"Word gagal mengubah '{source.name}' menjadi PDF: {exc}",
            detail=str(exc),
        ) from exc
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:  # pragma: no cover
            pass
        document = None
        try:
            if word is not None:
                word.Quit()
        except Exception:  # pragma: no cover
            pass
        word = None
        if initialised:
            # Every COM proxy must be released before the apartment is torn
            # down, otherwise the teardown raises RPC_S_CALL_FAILED.
            import gc  # noqa: PLC0415

            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:  # pragma: no cover
                pass

    if not destination.is_file():  # pragma: no cover - Word reported success
        raise PdfExportError(f"Word selesai tetapi {destination.name} tidak terbentuk.")
    return destination


def _convert_with_libreoffice(source: Path, destination: Path) -> Path:
    soffice = find_soffice()
    if soffice is None:
        raise PdfExportError(
            "LibreOffice (soffice) tidak ditemukan. Pasang LibreOffice atau "
            "pakai Microsoft Word."
        )
    outdir = destination.parent
    outdir.mkdir(parents=True, exist_ok=True)
    command: Sequence[str] = (
        str(soffice), "--headless", "--norestore",
        "--convert-to", "pdf", "--outdir", str(outdir), str(source),
    )
    try:
        completed = subprocess.run(
            command, capture_output=True, timeout=_SOFFICE_TIMEOUT, check=False
        )
    except subprocess.TimeoutExpired as exc:
        raise PdfExportError(
            f"LibreOffice tidak selesai dalam {_SOFFICE_TIMEOUT} detik.",
            detail=str(exc),
        ) from exc
    except OSError as exc:
        raise PdfExportError(
            f"Gagal menjalankan LibreOffice: {exc}", detail=str(exc)
        ) from exc

    produced = outdir / f"{source.stem}.pdf"
    if completed.returncode != 0 or not produced.is_file():
        detail = (completed.stderr or completed.stdout or b"").decode(
            "utf-8", "replace"
        ).strip()
        raise PdfExportError(
            f"LibreOffice gagal mengubah '{source.name}' menjadi PDF.",
            detail=detail or f"exit code {completed.returncode}",
        )
    if produced != destination:
        shutil.move(str(produced), destination)
    return destination


def docx_to_pdf(
    source: Path,
    destination: Optional[Path] = None,
    *,
    converter: Optional[str] = None,
    overwrite: bool = True,
) -> Path:
    """Convert one .docx to PDF.

    Args:
        source: the .docx to convert.
        destination: output path; defaults to the same name with a .pdf suffix.
        converter: ``"word"`` or ``"libreoffice"``; default picks what is
            available, Word first.
    """
    source = Path(source)
    if not source.is_file():
        raise PdfExportError(f"File .docx tidak ditemukan: {source}")
    if source.suffix.lower() not in (".docx", ".doc"):
        raise PdfExportError(
            f"Hanya file Word yang bisa diubah ke PDF, diterima: {source.suffix}"
        )

    destination = Path(destination) if destination else source.with_suffix(".pdf")
    if destination.exists():
        if not overwrite:
            raise PdfExportError(f"File PDF sudah ada: {destination}")
        try:
            destination.unlink()
        except OSError as exc:
            raise PdfExportError(
                f"File '{destination.name}' sedang dibuka aplikasi lain. "
                "Tutup dulu lalu coba lagi.",
                detail=str(exc),
            ) from exc
    destination.parent.mkdir(parents=True, exist_ok=True)

    if converter is None:
        converter = default_converter()
        if converter is None:
            raise PdfExportError(
                "Tidak ada converter PDF di komputer ini. Pasang Microsoft Word "
                "beserta pywin32 (pip install pywin32), atau pasang LibreOffice."
            )
    if converter not in CONVERTERS:
        raise PdfExportError(
            f"Converter tidak dikenali: {converter!r}. "
            f"Pilihan: {', '.join(CONVERTERS)}."
        )

    if converter == WORD:
        return _convert_with_word(source, destination)
    return _convert_with_libreoffice(source, destination)


def convert_many(
    sources: Sequence[Path],
    *,
    converter: Optional[str] = None,
    output_dir: Optional[Path] = None,
) -> List[Path]:
    """Convert several documents, keeping their names."""
    produced: List[Path] = []
    for source in sources:
        source = Path(source)
        target = (
            Path(output_dir) / f"{source.stem}.pdf" if output_dir
            else source.with_suffix(".pdf")
        )
        produced.append(docx_to_pdf(source, target, converter=converter))
    return produced
