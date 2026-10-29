"""Unit tests for writing results into a copy of the scenario workbook."""

from __future__ import annotations

import datetime as dt
import hashlib
import re
import zipfile

import pytest
from openpyxl import load_workbook

from rtg.excel_parser import parse_workbook
from rtg.excel_writer import (
    RESULT_FIELDS,
    ExcelWriteError,
    RowUpdate,
    column_index,
    column_letter,
    default_result_name,
    parse_date,
    to_excel_serial,
    write_results,
)
from rtg.config import load_config
from tests.fixtures import simple_workbook

CONFIG = load_config()

ROWS = [
    {"test_case_id": "1.1", "test_case": "Case Satu", "test_step_id": "1.1.1",
     "test_step": "buka aplikasi", "expected_results": "halaman tampil"},
    {"test_step_id": "1.1.2", "test_step": "login", "expected_results": "masuk dashboard"},
    {"test_case_id": "1.2", "test_case": "Case Dua", "test_step_id": "1.2.1",
     "test_step": "cek laporan", "expected_results": "laporan tampil"},
]


@pytest.fixture
def source(tmp_path):
    return simple_workbook(tmp_path / "src" / "Scenario.xlsx", rows=ROWS)


@pytest.fixture
def parsed(source):
    return parse_workbook(source, config=CONFIG).modules[0]


@pytest.fixture
def columns(parsed):
    return {parsed.name: parsed.column_map}


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize(
    "index, letters",
    [(1, "A"), (2, "B"), (26, "Z"), (27, "AA"), (28, "AB"), (52, "AZ"),
     (702, "ZZ"), (703, "AAA"), (1523, "BFO")],
)
def test_column_letter_round_trip(index, letters):
    assert column_letter(index) == letters
    assert column_index(letters) == index


def test_column_letter_rejects_zero():
    with pytest.raises(ValueError):
        column_letter(0)


@pytest.mark.parametrize(
    "value, serial",
    [
        (dt.date(1900, 1, 1), 2.0),
        (dt.date(2026, 10, 27), 46322.0),
        (dt.datetime(2026, 10, 27, 12, 0), 46322.5),
    ],
)
def test_to_excel_serial(value, serial):
    assert to_excel_serial(value) == pytest.approx(serial)


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("2026-10-27", dt.date(2026, 10, 27)),
        ("27/10/2026", dt.date(2026, 10, 27)),
        ("27-10-2026", dt.date(2026, 10, 27)),
        (dt.date(2026, 10, 27), dt.date(2026, 10, 27)),
        ("bukan tanggal", None),
        ("", None),
        (None, None),
    ],
)
def test_parse_date(raw, expected):
    assert parse_date(raw) == expected


def test_default_result_name():
    assert default_result_name("Scenario.xlsx") == "Scenario - Result.xlsx"
    assert default_result_name("a/b/Test Case.xlsm") == "Test Case - Result.xlsm"


# --------------------------------------------------------------------------- #
# Writing
# --------------------------------------------------------------------------- #

def _read(path, sheet, row, column):
    wb = load_workbook(path, data_only=True)
    try:
        return wb[sheet].cell(row=row, column=column).value
    finally:
        wb.close()


def test_write_results_fills_the_columns(source, parsed, columns, tmp_path):
    destination = tmp_path / "out" / "Result.xlsx"
    step = parsed.test_cases[0].steps[0]
    report = write_results(
        source, destination,
        {parsed.name: [RowUpdate(step.row, {
            "actual_results": "sesuai",
            "pass_fail": "P",
            "date_executed": "2026-10-27",
            "test_executed_by": "Mahrub",
        })]},
        columns,
    )

    assert report.cells_written == 4
    assert report.rows_written == 1
    assert report.sheets == [parsed.name]
    assert destination.is_file()

    cm = parsed.column_map
    assert _read(destination, parsed.name, step.row, cm["actual_results"]) == "sesuai"
    assert _read(destination, parsed.name, step.row, cm["pass_fail"]) == "P"
    assert _read(destination, parsed.name, step.row, cm["test_executed_by"]) == "Mahrub"


def test_original_workbook_is_never_touched(source, parsed, columns, tmp_path):
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    mtime = source.stat().st_mtime_ns

    write_results(
        source, tmp_path / "out.xlsx",
        {parsed.name: [RowUpdate(parsed.test_cases[0].steps[0].row, {"pass_fail": "P"})]},
        columns,
    )

    assert hashlib.sha256(source.read_bytes()).hexdigest() == before
    assert source.stat().st_mtime_ns == mtime


def test_untouched_rows_keep_their_values(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    first, second = parsed.test_cases[0].steps
    write_results(
        source, destination,
        {parsed.name: [RowUpdate(first.row, {"pass_fail": "P"})]},
        columns,
    )
    cm = parsed.column_map
    assert _read(destination, parsed.name, second.row, cm["test_step"]) == "login"
    assert _read(destination, parsed.name, second.row, cm["expected_results"]) == "masuk dashboard"
    assert _read(destination, parsed.name, second.row, cm["pass_fail"]) is None


def test_empty_values_are_not_written(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    step = parsed.test_cases[0].steps[0]
    report = write_results(
        source, destination,
        {parsed.name: [RowUpdate(step.row, {
            "pass_fail": "P", "actual_results": "", "test_executed_by": "   ",
        })]},
        columns,
    )
    assert report.cells_written == 1
    cm = parsed.column_map
    assert _read(destination, parsed.name, step.row, cm["actual_results"]) is None


def test_only_result_fields_are_written(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    step = parsed.test_cases[0].steps[0]
    report = write_results(
        source, destination,
        {parsed.name: [RowUpdate(step.row, {
            "pass_fail": "P",
            "test_step": "JANGAN DITIMPA",
            "expected_results": "JANGAN DITIMPA",
        })]},
        columns,
    )
    assert report.cells_written == 1
    cm = parsed.column_map
    assert _read(destination, parsed.name, step.row, cm["test_step"]) == "buka aplikasi"
    assert set(RESULT_FIELDS) == {
        "actual_results", "pass_fail", "date_executed", "test_executed_by"
    }


def test_date_written_into_an_empty_cell_stays_readable(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    step = parsed.test_cases[0].steps[0]
    write_results(
        source, destination,
        {parsed.name: [RowUpdate(step.row, {"date_executed": "2026-10-27"})]},
        columns,
    )
    # The source cell was empty and carries no date format, so the value goes
    # in as text rather than as a bare 46322.
    value = _read(destination, parsed.name, step.row, parsed.column_map["date_executed"])
    assert value == "2026-10-27"


def test_date_written_over_an_existing_date_keeps_the_date_format(tmp_path):
    rows = [
        {"test_case_id": "1.1", "test_case": "Case Satu", "test_step_id": "1.1.1",
         "test_step": "langkah", "date_executed": dt.datetime(2026, 7, 27)},
    ]
    source = simple_workbook(tmp_path / "dated.xlsx", rows=rows)
    module = parse_workbook(source, config=CONFIG).modules[0]
    destination = tmp_path / "out.xlsx"

    write_results(
        source, destination,
        {module.name: [RowUpdate(module.test_cases[0].steps[0].row,
                                 {"date_executed": "2026-10-27"})]},
        {module.name: module.column_map},
    )
    value = _read(destination, module.name, module.test_cases[0].steps[0].row,
                  module.column_map["date_executed"])
    assert isinstance(value, dt.datetime)
    assert value.date() == dt.date(2026, 10, 27)


def test_missing_target_column_is_reported_not_fatal(tmp_path):
    rows = [{"test_case_id": "1.1", "test_case": "A", "test_step": "x"}]
    source = simple_workbook(
        tmp_path / "minimal.xlsx",
        fields=("test_case_id", "test_case", "test_step", "pass_fail"),
        rows=rows,
    )
    module = parse_workbook(source, config=CONFIG).modules[0]
    destination = tmp_path / "out.xlsx"

    report = write_results(
        source, destination,
        {module.name: [RowUpdate(module.test_cases[0].steps[0].row, {
            "pass_fail": "P", "actual_results": "tidak ada kolomnya",
        })]},
        {module.name: module.column_map},
    )
    assert report.cells_written == 1
    assert "actual_results" in report.skipped_fields
    assert "tidak ada di sheet" in " ".join(report.summary_lines())


def test_cell_is_created_when_the_row_had_none(source, parsed, columns, tmp_path):
    # The second step leaves every result column empty in the source.
    destination = tmp_path / "out.xlsx"
    step = parsed.test_cases[0].steps[1]
    write_results(
        source, destination,
        {parsed.name: [RowUpdate(step.row, {"actual_results": "dibuat baru"})]},
        columns,
    )
    assert _read(destination, parsed.name, step.row,
                 parsed.column_map["actual_results"]) == "dibuat baru"


def test_writes_to_several_test_cases_and_rows(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    updates = [
        RowUpdate(step.row, {"pass_fail": "P"})
        for case in parsed.test_cases for step in case.steps
    ]
    report = write_results(source, destination, {parsed.name: updates}, columns)
    assert report.rows_written == 3
    assert report.cells_written == 3


# --------------------------------------------------------------------------- #
# Package integrity
# --------------------------------------------------------------------------- #

def test_every_other_part_is_byte_identical(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    write_results(
        source, destination,
        {parsed.name: [RowUpdate(parsed.test_cases[0].steps[0].row, {"pass_fail": "P"})]},
        columns,
    )
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(destination) as b:
        names_a, names_b = set(a.namelist()), set(b.namelist())
        assert names_a == names_b
        changed = [n for n in sorted(names_a) if a.read(n) != b.read(n)]
        assert len(changed) == 1
        assert changed[0].startswith("xl/worksheets/")


def test_root_namespace_declarations_survive(source, parsed, columns, tmp_path):
    """Excel refuses a sheet whose mc:Ignorable names undeclared prefixes."""
    part = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(source) as zf:
        original = {name: zf.read(name) for name in zf.namelist()}

    # Give the sheet the namespace soup a real Excel file has.
    decorated = original[part].replace(
        b"<worksheet ",
        b'<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        b' mc:Ignorable="x14ac xr xr2"'
        b' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
        b' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'
        b' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" ',
        1,
    )
    with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in original.items():
            zf.writestr(name, decorated if name == part else data)

    destination = tmp_path / "out.xlsx"
    write_results(
        source, destination,
        {parsed.name: [RowUpdate(parsed.test_cases[0].steps[0].row, {"pass_fail": "P"})]},
        columns,
    )

    with zipfile.ZipFile(destination) as zf:
        produced = zf.read(part)
    root = re.search(rb"<worksheet\b[^>]*>", produced).group(0).decode()
    for prefix in ("mc", "x14ac", "xr", "xr2"):
        assert f"xmlns:{prefix}=" in root, f"xmlns:{prefix} hilang"
    assert 'mc:Ignorable="x14ac xr xr2"' in root
    assert produced.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')


# --------------------------------------------------------------------------- #
# Errors
# --------------------------------------------------------------------------- #

def test_refuses_to_overwrite_the_source(source, parsed, columns):
    with pytest.raises(ExcelWriteError, match="tidak boleh menimpa"):
        write_results(source, source, {parsed.name: [RowUpdate(8, {"pass_fail": "P"})]},
                      columns)


def test_refuses_an_existing_destination_unless_told(source, parsed, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    destination.write_text("sudah ada", encoding="utf-8")
    updates = {parsed.name: [RowUpdate(parsed.test_cases[0].steps[0].row, {"pass_fail": "P"})]}

    with pytest.raises(ExcelWriteError, match="sudah ada"):
        write_results(source, destination, updates, columns)

    report = write_results(source, destination, updates, columns, overwrite=True)
    assert report.cells_written == 1


def test_missing_source(tmp_path):
    with pytest.raises(ExcelWriteError, match="tidak ditemukan"):
        write_results(tmp_path / "ghost.xlsx", tmp_path / "out.xlsx", {}, {})


def test_unknown_sheet_leaves_no_half_written_file(source, columns, tmp_path):
    destination = tmp_path / "out.xlsx"
    with pytest.raises(ExcelWriteError, match="tidak ada di file Excel"):
        write_results(
            source, destination,
            {"Sheet Hantu": [RowUpdate(8, {"pass_fail": "P"})]},
            columns,
        )
    assert not destination.exists()


def test_corrupt_source(tmp_path):
    broken = tmp_path / "broken.xlsx"
    broken.write_text("bukan zip", encoding="utf-8")
    destination = tmp_path / "out.xlsx"
    with pytest.raises(ExcelWriteError, match="tidak bisa dibaca"):
        write_results(broken, destination, {"A": [RowUpdate(1, {"pass_fail": "P"})]}, {})
    assert not destination.exists()


def test_no_updates_still_produces_a_faithful_copy(source, parsed, tmp_path):
    destination = tmp_path / "out.xlsx"
    report = write_results(source, destination, {}, {})
    assert report.cells_written == 0
    assert destination.is_file()
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(destination) as b:
        assert all(a.read(n) == b.read(n) for n in a.namelist())
