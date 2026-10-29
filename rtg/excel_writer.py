"""Write execution results back into a *copy* of the scenario workbook.

The original file is never touched: it is copied byte for byte, and only the
one worksheet XML part inside the copy is edited.

Why not openpyxl? Loading and saving the real scenario file with openpyxl
loses content: EMF images are dropped, JPEG logos are re-encoded as PNG, and
``xl/richData``, ``customXml``, ``metadata.xml`` and the printer settings
disappear (603 KB -> 525 KB on the reference file). A result workbook is a
deliverable, so that is not acceptable. Editing the sheet XML inside the zip
leaves every other part exactly as Excel wrote it.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

from .errors import ResultGeneratorError
from .normalize import cell_text

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CALC_CHAIN = "xl/calcChain.xml"

#: Columns this module knows how to fill in.
RESULT_FIELDS: Tuple[str, ...] = (
    "actual_results",
    "pass_fail",
    "date_executed",
    "test_executed_by",
)

#: Excel's day zero. 1900 is treated as a leap year, hence the two-day offset.
_EXCEL_EPOCH = dt.date(1899, 12, 30)

_XMLNS_RE = re.compile(rb'xmlns:([A-Za-z0-9_.-]+)\s*=\s*"([^"]+)"')
_CELL_REF_RE = re.compile(r"^([A-Z]+)([0-9]+)$")
_ROOT_TAG_RE = re.compile(rb"<[A-Za-z_][\w.:-]*[^>]*>")
_DECL_RE = re.compile(rb'xmlns(:[A-Za-z0-9_.-]+)?="[^"]*"')
_XML_HEADER_RE = re.compile(rb"^\s*<\?xml[^>]*\?>")
DEFAULT_XML_HEADER = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
)


class ExcelWriteError(ResultGeneratorError):
    """Raised when the result workbook cannot be produced."""


# --------------------------------------------------------------------------- #
# Column letters and dates
# --------------------------------------------------------------------------- #

def column_letter(index: int) -> str:
    """1 -> ``A``, 27 -> ``AA``."""
    if index < 1:
        raise ValueError(f"Nomor kolom harus >= 1, diterima {index}")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def column_index(letters: str) -> int:
    """``A`` -> 1, ``AA`` -> 27."""
    value = 0
    for char in letters.upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def to_excel_serial(value: dt.date) -> float:
    """Convert a date/datetime to the serial number Excel stores."""
    if isinstance(value, dt.datetime):
        days = (value.date() - _EXCEL_EPOCH).days
        seconds = value.hour * 3600 + value.minute * 60 + value.second
        return days + seconds / 86400.0
    return float((value - _EXCEL_EPOCH).days)


def parse_date(value: Any) -> Optional[dt.date]:
    """Best effort date parsing for what the GUI or a project file holds."""
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return value
    text = cell_text(value)
    if not text:
        return None
    for pattern in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            parsed = dt.datetime.strptime(text, pattern)
        except ValueError:
            continue
        return parsed if pattern.endswith("%H:%M") else parsed.date()
    return None


# --------------------------------------------------------------------------- #
# Input
# --------------------------------------------------------------------------- #

@dataclass
class RowUpdate:
    """The values to write into one sheet row."""

    row: int
    values: Dict[str, str] = field(default_factory=dict)

    def cleaned(self) -> Dict[str, str]:
        return {
            name: value
            for name, value in self.values.items()
            if name in RESULT_FIELDS and cell_text(value) != ""
        }


@dataclass
class WriteReport:
    """What the write-back actually did."""

    path: Path
    sheets: List[str] = field(default_factory=list)
    cells_written: int = 0
    rows_written: int = 0
    skipped_fields: List[str] = field(default_factory=list)
    dropped_calc_chain: bool = False

    def summary_lines(self) -> List[str]:
        lines = [
            f"File hasil : {self.path.name}",
            f"Sheet      : {', '.join(self.sheets) or '-'}",
            f"Terisi     : {self.cells_written} sel di {self.rows_written} baris",
        ]
        if self.skipped_fields:
            lines.append(
                "Kolom tidak ada di sheet (dilewati): "
                + ", ".join(sorted(set(self.skipped_fields)))
            )
        if self.dropped_calc_chain:
            lines.append(
                "Cache perhitungan (calcChain) dibuang karena ada rumus yang ditimpa; "
                "Excel akan membangunnya ulang saat file dibuka."
            )
        return lines

    def to_dict(self) -> Dict[str, Any]:
        return {
            "path": str(self.path),
            "sheets": list(self.sheets),
            "cells_written": self.cells_written,
            "rows_written": self.rows_written,
            "skipped_fields": list(self.skipped_fields),
            "dropped_calc_chain": self.dropped_calc_chain,
        }


# --------------------------------------------------------------------------- #
# Sheet XML surgery
# --------------------------------------------------------------------------- #

def _register_namespaces(xml_bytes: bytes) -> None:
    """Keep the prefixes Excel used.

    ``mc:Ignorable="x14ac xr xr2"`` names prefixes literally, so letting
    ElementTree rename them to ns0/ns1 would leave that attribute pointing at
    prefixes that no longer exist.
    """
    ET.register_namespace("", MAIN_NS)
    for prefix, uri in _XMLNS_RE.findall(xml_bytes[:4096]):
        ET.register_namespace(prefix.decode("ascii"), uri.decode("ascii"))


def _namespace_decls(tag: bytes) -> Dict[bytes, bytes]:
    """``{prefix: full xmlns declaration}`` found in an element start tag."""
    return {match.group(1) or b"": match.group(0) for match in _DECL_RE.finditer(tag)}


def _restore_root(original: bytes, produced: bytes) -> bytes:
    """Put back what ElementTree drops when re-serialising a sheet.

    Two things matter to Excel:

    * every ``xmlns`` declaration on the root. ElementTree only emits the
      namespaces a tree actually uses, but ``mc:Ignorable="x14ac xr xr2 xr3"``
      names prefixes literally - dropping their declarations makes Excel refuse
      to open the workbook;
    * the ``standalone="yes"`` XML declaration Excel itself writes.
    """
    original_match = _ROOT_TAG_RE.search(original)
    produced_match = _ROOT_TAG_RE.search(produced)
    if original_match and produced_match:
        original_tag = original_match.group(0)
        produced_tag = produced_match.group(0)
        have = _namespace_decls(produced_tag)
        missing = [
            decl for prefix, decl in _namespace_decls(original_tag).items()
            if prefix not in have
        ]
        if missing:
            name_end = produced_tag.find(b" ")
            if name_end == -1:
                name_end = len(produced_tag) - (2 if produced_tag.endswith(b"/>") else 1)
            patched = (
                produced_tag[:name_end] + b" " + b" ".join(missing)
                + produced_tag[name_end:]
            )
            produced = produced.replace(produced_tag, patched, 1)

    header_match = _XML_HEADER_RE.search(original)
    header = header_match.group(0) + b"\r\n" if header_match else DEFAULT_XML_HEADER
    produced = _XML_HEADER_RE.sub(b"", produced).lstrip(b"\r\n")
    return header + produced


def _tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def _row_number(element) -> int:
    try:
        return int(element.get("r", "0"))
    except ValueError:  # pragma: no cover - malformed file
        return 0


def _cell_column(element) -> int:
    match = _CELL_REF_RE.match(element.get("r", ""))
    return column_index(match.group(1)) if match else 0


def _find_or_create_row(sheet_data, row_number: int):
    insert_at = len(sheet_data)
    for index, element in enumerate(sheet_data):
        number = _row_number(element)
        if number == row_number:
            return element, False
        if number > row_number:
            insert_at = index
            break
    row = ET.Element(_tag("row"), {"r": str(row_number)})
    sheet_data.insert(insert_at, row)
    return row, True


def _find_or_create_cell(row, column: int, row_number: int):
    reference = f"{column_letter(column)}{row_number}"
    insert_at = len(row)
    for index, element in enumerate(row):
        if element.tag != _tag("c"):
            continue
        current = _cell_column(element)
        if current == column:
            return element
        if current > column:
            insert_at = index
            break
    cell = ET.Element(_tag("c"), {"r": reference})
    row.insert(insert_at, cell)
    return cell


def _looks_numeric(cell) -> bool:
    """True when the cell currently holds a plain number (so: a date cell)."""
    if cell.get("t") not in (None, "n"):
        return False
    value = cell.find(_tag("v"))
    if value is None or not (value.text or "").strip():
        return False
    try:
        float(value.text)
    except ValueError:
        return False
    return True


def _set_cell(cell, value: Any, *, as_serial: bool) -> bool:
    """Write ``value`` into a cell, keeping its style. True if it had a formula.

    Text is written as an inline string so ``sharedStrings.xml`` never has to
    be touched. ``as_serial`` is decided by the caller *before* the old
    children are dropped: a date is stored as a serial number only when the
    cell already held a number, which means it already carries a date format -
    otherwise the user would see 46022 instead of a date.
    """
    had_formula = cell.find(_tag("f")) is not None
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)

    if as_serial:
        date_value = parse_date(value)
        if date_value is not None:
            node = ET.SubElement(cell, _tag("v"))
            node.text = repr(to_excel_serial(date_value)).rstrip("0").rstrip(".")
            return had_formula

    cell.set("t", "inlineStr")
    is_node = ET.SubElement(cell, _tag("is"))
    t_node = ET.SubElement(is_node, _tag("t"))
    t_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_node.text = cell_text(value)
    return had_formula


def _edit_sheet_xml(
    xml_bytes: bytes,
    updates: Sequence[RowUpdate],
    columns: Mapping[str, int],
) -> Tuple[bytes, int, int, bool]:
    """Apply the updates to one worksheet part."""
    _register_namespaces(xml_bytes)
    root = ET.fromstring(xml_bytes)
    sheet_data = root.find(_tag("sheetData"))
    if sheet_data is None:
        raise ExcelWriteError("Sheet tidak punya <sheetData>; file mungkin rusak.")

    cells_written = 0
    rows_written = 0
    formula_replaced = False

    for update in updates:
        values = update.cleaned()
        if not values:
            continue
        row_element, _created = _find_or_create_row(sheet_data, update.row)
        touched = False
        for field_name, value in values.items():
            column = columns.get(field_name)
            if not column:
                continue
            cell = _find_or_create_cell(row_element, column, update.row)
            # Checked before the cell is cleared: a date column already holds
            # numbers, an empty or textual one does not.
            as_serial = _looks_numeric(cell) and parse_date(value) is not None
            formula_replaced |= _set_cell(cell, value, as_serial=as_serial)
            cells_written += 1
            touched = True
        if touched:
            rows_written += 1

    body = _restore_root(xml_bytes, ET.tostring(root, encoding="UTF-8"))
    return body, cells_written, rows_written, formula_replaced


# --------------------------------------------------------------------------- #
# Package handling
# --------------------------------------------------------------------------- #

def _sheet_part_names(zf: zipfile.ZipFile) -> Dict[str, str]:
    """Map sheet name -> part name, via workbook.xml and its relationships."""
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read(WORKBOOK_RELS))
    targets = {
        rel.get("Id"): rel.get("Target", "")
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    r_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

    mapping: Dict[str, str] = {}
    sheets = workbook.find(_tag("sheets"))
    for sheet in (sheets if sheets is not None else []):
        target = targets.get(sheet.get(r_id), "")
        if not target:
            continue
        part = target[1:] if target.startswith("/") else f"xl/{target.lstrip('/')}"
        part = part.replace("/./", "/")
        mapping[sheet.get("name", "")] = part
    return mapping


def _strip_calc_chain(name: str, data: bytes) -> bytes:
    """Remove the calcChain references from a package part."""
    if name == CONTENT_TYPES:
        text = data.decode("utf-8")
        text = re.sub(r"<Override[^>]*calcChain\.xml[^>]*/>", "", text)
        return text.encode("utf-8")
    if name == WORKBOOK_RELS:
        text = data.decode("utf-8")
        text = re.sub(r"<Relationship[^>]*calcChain\.xml[^>]*/>", "", text)
        return text.encode("utf-8")
    return data


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #

def write_results(
    source: Path,
    destination: Path,
    updates: Mapping[str, Sequence[RowUpdate]],
    columns: Mapping[str, Mapping[str, int]],
    *,
    overwrite: bool = False,
) -> WriteReport:
    """Copy ``source`` to ``destination`` and fill in the result columns.

    Args:
        source: the original scenario workbook. Only read.
        destination: where the filled-in copy goes. Must differ from source.
        updates: ``{sheet name: [RowUpdate, ...]}``.
        columns: ``{sheet name: {field: 1-based column}}`` - normally
            ``Module.column_map`` from the parser, so the target columns are
            found by header name just like everything else.
    """
    source = Path(source)
    destination = Path(destination)
    if not source.is_file():
        raise ExcelWriteError(f"File Excel asli tidak ditemukan: {source}")
    try:
        same = source.resolve() == destination.resolve()
    except OSError:  # pragma: no cover - unresolvable path
        same = str(source) == str(destination)
    if same:
        raise ExcelWriteError(
            "File hasil tidak boleh menimpa file skenario asli. "
            "Pilih nama atau folder lain."
        )
    if destination.exists() and not overwrite:
        raise ExcelWriteError(f"File sudah ada: {destination}")

    report = WriteReport(path=destination)
    for sheet_name, rows in updates.items():
        mapping = columns.get(sheet_name, {})
        for row in rows:
            for field_name in row.cleaned():
                if not mapping.get(field_name):
                    report.skipped_fields.append(field_name)

    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(source, destination)
    except OSError as exc:
        raise ExcelWriteError(
            f"Gagal menyalin file Excel ke {destination}: {exc}", detail=str(exc)
        ) from exc

    try:
        with zipfile.ZipFile(destination) as zf:
            parts = _sheet_part_names(zf)
            originals = {item.filename: zf.read(item.filename) for item in zf.infolist()}
            infos = list(zf.infolist())
    except (zipfile.BadZipFile, KeyError) as exc:
        destination.unlink(missing_ok=True)
        raise ExcelWriteError(
            f"File '{source.name}' tidak bisa dibaca sebagai workbook Excel.",
            detail=str(exc),
        ) from exc

    edited: Dict[str, bytes] = {}
    formula_replaced = False
    for sheet_name, rows in updates.items():
        if not rows:
            continue
        part = parts.get(sheet_name)
        if part is None:
            destination.unlink(missing_ok=True)
            raise ExcelWriteError(
                f"Sheet '{sheet_name}' tidak ada di file Excel. "
                "Tersedia: " + ", ".join(sorted(parts)) + "."
            )
        if part not in originals:
            destination.unlink(missing_ok=True)
            raise ExcelWriteError(f"Bagian '{part}' tidak ada di dalam file Excel.")
        body, cells, written_rows, had_formula = _edit_sheet_xml(
            originals[part], rows, columns.get(sheet_name, {})
        )
        if cells:
            edited[part] = body
            report.sheets.append(sheet_name)
            report.cells_written += cells
            report.rows_written += written_rows
            formula_replaced |= had_formula

    report.dropped_calc_chain = formula_replaced and CALC_CHAIN in originals

    temp = destination.with_name(destination.name + ".tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in infos:
                name = info.filename
                if report.dropped_calc_chain and name == CALC_CHAIN:
                    continue
                data = edited.get(name, originals[name])
                if report.dropped_calc_chain:
                    data = _strip_calc_chain(name, data)
                item = zipfile.ZipInfo(name, date_time=info.date_time)
                item.compress_type = info.compress_type
                item.external_attr = info.external_attr
                item.internal_attr = info.internal_attr
                item.create_system = info.create_system
                out.writestr(item, data)
        temp.replace(destination)
    except OSError as exc:
        temp.unlink(missing_ok=True)
        destination.unlink(missing_ok=True)
        raise ExcelWriteError(
            f"Gagal menulis file hasil {destination}: {exc}", detail=str(exc)
        ) from exc

    return report


def default_result_name(source: Path) -> str:
    """``Scenario.xlsx`` -> ``Scenario - Result.xlsx``."""
    source = Path(source)
    return f"{source.stem} - Result{source.suffix or '.xlsx'}"
