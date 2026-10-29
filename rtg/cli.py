"""Command line front-end used to verify the parser without the GUI.

    python -m rtg.cli inspect  "Scenario UAT.xlsx"
    python -m rtg.cli parse    "Scenario UAT.xlsx" --sheet "AMF Segmentation"
    python -m rtg.cli parse    "Scenario UAT.xlsx" --json out.json
    python -m rtg.cli generate "Scenario UAT.xlsx" --out output
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List, Optional

from .config import load_config, write_default_config
from .docx_builder import export, summarize
from .errors import ResultGeneratorError
from .excel_parser import inspect_workbook, parse_workbook


def _print_inspection(path: Path, config) -> int:
    inspection = inspect_workbook(path, config=config)
    print(f"File   : {inspection.path}")
    print(f"Sheet  : {len(inspection.sheets)} total, "
          f"{len(inspection.scenario_sheets)} sheet skenario terdeteksi")
    print()
    for sheet in inspection.sheets:
        if sheet.is_scenario:
            mark = "[x]"
        elif sheet.is_mapping_candidate:
            mark = "[?]"
        else:
            mark = "[ ]"
        print(f"{mark} {sheet.name}")
        if sheet.header_row:
            print(f"      header baris {sheet.header_row}, "
                  f"{len(sheet.column_map)} kolom dikenali, "
                  f"~{sheet.data_row_estimate} baris step")
            mapped = ", ".join(
                f"{name}=kol{col}" for name, col in sorted(
                    sheet.column_map.items(), key=lambda kv: kv[1]
                )
            )
            print(f"      {mapped}")
            if sheet.unmapped_columns:
                extras = ", ".join(text for _c, text in sheet.unmapped_columns)
                print(f"      kolom tidak dikenali: {extras}")
            if sheet.metadata:
                meta = ", ".join(f"{k}={v}" for k, v in sheet.metadata.items())
                print(f"      metadata: {meta}")
        if sheet.reason:
            print(f"      -> {sheet.reason}")
        print()

    print("Keterangan: [x] sheet skenario  [?] perlu mapping manual  [ ] bukan sheet skenario")
    return 0 if inspection.scenario_sheets else 1


def _print_tree(scenario_file) -> None:
    for module in scenario_file.modules:
        print(f"[MODULE] {module.name}  "
              f"({len(module.test_cases)} test case, {module.step_count} step)")
        if module.metadata:
            for key, value in module.metadata.items():
                print(f"    meta: {key} = {value}")
        for group in module.scenarios:
            print(f"  - {group.title}")
            for case in group.test_cases:
                print(f"      * {case.display_label}   [uid: {case.uid}]")
                for index, step in enumerate(case.steps, start=1):
                    text = step.test_step.replace("\n", " ")
                    if len(text) > 70:
                        text = text[:67] + "..."
                    label = step.test_step_id or str(index)
                    print(f"          {label}. {text}")
        print()
    if scenario_file.warnings:
        print("Peringatan:")
        for warning in scenario_file.warnings:
            print(f"  ! {warning}")


def _print_reports(reports) -> int:
    """Print the per-module summary required before exporting."""
    incomplete = 0
    for report in reports:
        print(f"[MODULE] {report.title}")
        print(f"    test case     : {report.test_case_count}")
        print(f"    screenshot    : {report.screenshot_count}")
        if report.without_screenshot:
            incomplete += len(report.without_screenshot)
            print(f"    belum ada screenshot : {len(report.without_screenshot)} test case")
            for label in report.without_screenshot:
                print(f"        - {label}")
        if report.missing_files:
            print(f"    file screenshot hilang : {len(report.missing_files)}")
            for name in report.missing_files:
                print(f"        ! {name}")
        print()
    return incomplete


def _generate(args, config) -> int:
    scenario_file = parse_workbook(args.excel, args.sheets, config=config)

    screenshots = None
    if args.screenshots:
        screenshots = json.loads(Path(args.screenshots).read_text(encoding="utf-8"))

    reports = summarize(scenario_file, screenshots)
    incomplete = _print_reports(reports)

    if args.dry_run:
        print("--dry-run: tidak ada file yang ditulis.")
        return 0

    results = export(
        scenario_file,
        screenshots,
        output_dir=args.out,
        combined=args.combined,
        template=args.template,
    )
    for result in results:
        print(f"Dokumen ditulis : {result.path}")
        print(f"    {result.test_case_count} test case, "
              f"{result.screenshot_count} screenshot")
    if incomplete:
        print(f"\nCatatan: {incomplete} test case masih bertanda "
              "'[Belum ada screenshot]' di dokumen.")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="rtg", description="Result Test Generator - parser skenario Excel"
    )
    parser.add_argument("--config", type=Path, default=None,
                        help="file alias kolom (JSON); default: bawaan aplikasi")
    sub = parser.add_subparsers(dest="command", required=True)

    p_inspect = sub.add_parser("inspect", help="deteksi sheet skenario & mapping kolom")
    p_inspect.add_argument("excel", type=Path)

    p_parse = sub.add_parser("parse", help="baca skenario dan tampilkan tree")
    p_parse.add_argument("excel", type=Path)
    p_parse.add_argument("--sheet", action="append", dest="sheets", default=None,
                         help="nama sheet (boleh diulang); default: semua sheet skenario")
    p_parse.add_argument("--json", type=Path, default=None,
                         help="simpan hasil parse ke file JSON")

    p_gen = sub.add_parser("generate", help="buat dokumen Result Test UAT (.docx)")
    p_gen.add_argument("excel", type=Path)
    p_gen.add_argument("--sheet", action="append", dest="sheets", default=None,
                       help="nama sheet/module (boleh diulang)")
    p_gen.add_argument("--out", type=Path, default=Path("output"),
                       help="folder output (default: ./output)")
    p_gen.add_argument("--combined", action="store_true",
                       help="satu file gabungan untuk semua module (default: satu file per module)")
    p_gen.add_argument("--template", type=Path, default=None,
                       help="file .docx milik user sebagai template style")
    p_gen.add_argument("--screenshots", type=Path, default=None,
                       help='file JSON {"Module::1.1": ["a.png", {"path": "b.png", "caption": "..."}]}')
    p_gen.add_argument("--dry-run", action="store_true",
                       help="hanya tampilkan ringkasan per module, tanpa menulis file")

    p_config = sub.add_parser("write-config", help="salin config alias bawaan agar bisa diedit")
    p_config.add_argument("destination", type=Path, nargs="?", default=Path("config"))
    p_config.add_argument("--overwrite", action="store_true")

    args = parser.parse_args(argv)

    try:
        if args.command == "write-config":
            written = write_default_config(args.destination, overwrite=args.overwrite)
            print(f"Config ditulis ke: {written}")
            return 0

        config = load_config(args.config)
        if args.command == "inspect":
            return _print_inspection(args.excel, config)
        if args.command == "generate":
            return _generate(args, config)

        scenario_file = parse_workbook(args.excel, args.sheets, config=config)
        _print_tree(scenario_file)
        if args.json:
            args.json.parent.mkdir(parents=True, exist_ok=True)
            args.json.write_text(
                json.dumps(scenario_file.to_dict(), indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            print(f"JSON ditulis ke: {args.json}")
        return 0
    except ResultGeneratorError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except (FileNotFoundError, FileExistsError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
