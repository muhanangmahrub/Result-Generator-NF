"""Dialogs: sheet picker, manual column mapping, metadata, re-import, export."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..config import FieldConfig
from ..excel_parser import SheetInspection, SheetOverride, WorkbookInspection
from ..project import Project
from ..scenario_diff import (
    EDITED,
    MISSING,
    REMOVED,
    SUSPECT,
    CaseChange,
    ScenarioChanges,
    match_steps,
)

FIELD_LABELS: Dict[str, str] = {
    "scenario_id": "Scenario ID",
    "test_scenario": "Test Scenario",
    "test_scenario_description": "Test Scenario Description",
    "test_case_id": "Test Case ID",
    "test_case": "Test Case",
    "test_step_id": "Test Step ID",
    "test_step": "Test Step",
    "positive_negative": "Positive/Negative",
    "pre_condition": "Pre-Condition",
    "test_data": "Test Data",
    "expected_results": "Expected Results",
    "actual_results": "Actual Results",
    "pass_fail": "Pass/Fail",
    "date_executed": "Date Executed",
    "test_executed_by": "Test Executed by",
    "notes": "Notes",
}

METADATA_LABELS: Dict[str, str] = {
    "project_name": "Project Name",
    "module_name": "Module",
    "version": "Version",
    "environment_testing": "Environment Testing",
    "type_testing": "Type Testing",
    "document_title": "Document Title",
    "created_by": "Created By",
    "created_date": "Created Date",
    "reviewed_by": "Reviewed By",
    "approved_by": "Approved By",
}


def field_label(name: str) -> str:
    return FIELD_LABELS.get(name, name.replace("_", " ").title())


# --------------------------------------------------------------------------- #
# Manual column mapping
# --------------------------------------------------------------------------- #

class ColumnMappingDialog(QDialog):
    """Map application fields to the headers actually present in a sheet.

    Shown when a sheet is missing a mandatory column, and reachable at any time
    from the sheet picker. The result is stored in project.json, which pins the
    mapping for that sheet until the user resets it.
    """

    def __init__(
        self,
        inspection: SheetInspection,
        config: FieldConfig,
        *,
        header_choices: Sequence[tuple],
        current: Optional[SheetOverride] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Mapping kolom manual - {inspection.name}")
        self.resize(640, 560)
        self._inspection = inspection
        self._config = config
        self._header_choices = list(header_choices)
        self._reset = False
        self._combos: Dict[str, QComboBox] = {}

        layout = QVBoxLayout(self)

        intro = QLabel(
            f"Sheet <b>{inspection.name}</b>. Pilih kolom untuk setiap field. "
            "Field bertanda <b>*</b> wajib ada.<br>"
            "Mapping ini disimpan di project.json dan mengunci kolom sheet ini."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        row = QHBoxLayout()
        row.addWidget(QLabel("Baris header:"))
        self.header_row = QSpinBox()
        self.header_row.setRange(1, max(1, config.header_scan_rows))
        self.header_row.setValue(inspection.header_row or 1)
        row.addWidget(self.header_row)
        row.addStretch(1)
        layout.addLayout(row)

        box = QGroupBox("Kolom")
        form = QFormLayout(box)
        required = set(config.required_fields)
        for name in config.known_fields:
            combo = QComboBox()
            combo.addItem("(tidak dipakai)", None)
            for column, text in self._header_choices:
                combo.addItem(f"kol {column} - {text}", column)
            column = inspection.column_map.get(name)
            if column is not None:
                index = combo.findData(column)
                combo.setCurrentIndex(index if index >= 0 else 0)
            self._combos[name] = combo
            label = field_label(name) + (" *" if name in required else "")
            form.addRow(label, combo)
        layout.addWidget(box)

        self.warning = QLabel()
        self.warning.setWordWrap(True)
        self.warning.setProperty("role", "danger")
        layout.addWidget(self.warning)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        reset = buttons.addButton(
            "Reset ke deteksi otomatis", QDialogButtonBox.ButtonRole.ResetRole
        )
        reset.clicked.connect(self._on_reset)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _selected(self) -> Dict[str, int]:
        chosen: Dict[str, int] = {}
        for name, combo in self._combos.items():
            value = combo.currentData()
            if value is not None:
                chosen[name] = int(value)
        return chosen

    def _on_reset(self) -> None:
        self._reset = True
        self.accept()

    def _on_accept(self) -> None:
        chosen = self._selected()
        missing = [f for f in self._config.required_fields if f not in chosen]
        if missing:
            self.warning.setText(
                "Kolom wajib belum dipilih: "
                + ", ".join(field_label(f) for f in missing)
            )
            return
        duplicates = [c for c in chosen.values() if list(chosen.values()).count(c) > 1]
        if duplicates:
            self.warning.setText(
                "Satu kolom dipakai untuk lebih dari satu field: kol "
                + ", ".join(str(c) for c in sorted(set(duplicates)))
            )
            return
        self.accept()

    def result_override(self) -> Optional[SheetOverride]:
        """``None`` means: go back to automatic detection."""
        if self._reset:
            return None
        return SheetOverride(
            header_row=int(self.header_row.value()), columns=self._selected()
        )


# --------------------------------------------------------------------------- #
# Sheet picker
# --------------------------------------------------------------------------- #

class SheetPickerDialog(QDialog):
    """Checkbox list of the sheets detected in a workbook."""

    COLUMNS = ("Sheet (module)", "Status", "Header", "Kolom", "Baris step")

    def __init__(
        self,
        inspection: WorkbookInspection,
        config: FieldConfig,
        *,
        reinspect: Callable[[Dict[str, Any]], WorkbookInspection],
        selected: Optional[Iterable[str]] = None,
        overrides: Optional[Mapping[str, SheetOverride]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Pilih module (sheet skenario)")
        self.resize(820, 480)
        self._config = config
        self._reinspect = reinspect
        self._inspection = inspection
        self._auto_columns = {s.name: dict(s.column_map) for s in inspection.sheets}
        self.overrides: Dict[str, SheetOverride] = dict(overrides or {})
        self._preselected = set(selected) if selected is not None else None

        layout = QVBoxLayout(self)
        intro = QLabel(
            "Sheet dianggap sheet skenario kalau punya baris header dengan kolom "
            "<b>Test Case ID</b>, <b>Test Case</b> dan <b>Test Step</b> - nama sheet "
            "tidak dipakai sebagai patokan.<br>"
            "Tanda <b>?</b> berarti sheet itu mirip sheet skenario tetapi ada kolom "
            "wajib yang belum dikenali; pakai <i>Mapping kolom manual</i>."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for index in range(1, len(self.COLUMNS)):
            header.setSectionResizeMode(index, QHeaderView.ResizeMode.ResizeToContents)
        self.table.itemSelectionChanged.connect(self._update_buttons)
        self.table.itemDoubleClicked.connect(lambda _item: self._open_mapping())
        layout.addWidget(self.table)

        row = QHBoxLayout()
        self.select_all = QPushButton("Pilih semua sheet skenario")
        self.select_all.clicked.connect(self._select_all)
        row.addWidget(self.select_all)
        self.clear_all = QPushButton("Kosongkan")
        self.clear_all.clicked.connect(self._clear_all)
        row.addWidget(self.clear_all)
        row.addStretch(1)
        self.map_button = QPushButton("Mapping kolom manual...")
        self.map_button.clicked.connect(self._open_mapping)
        row.addWidget(self.map_button)
        layout.addLayout(row)

        self.detail = QLabel()
        self.detail.setWordWrap(True)
        self.detail.setMinimumHeight(48)
        layout.addWidget(self.detail)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.accepted.connect(self._on_accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self._fill_table()

    # -- table ------------------------------------------------------------- #

    def _fill_table(self) -> None:
        sheets = self._inspection.sheets
        self.table.setRowCount(len(sheets))
        for row, sheet in enumerate(sheets):
            name_item = QTableWidgetItem(sheet.name)
            name_item.setFlags(
                name_item.flags() | Qt.ItemFlag.ItemIsUserCheckable
            )
            selectable = sheet.is_scenario
            if self._preselected is None:
                checked = selectable
            else:
                checked = sheet.name in self._preselected
            name_item.setCheckState(
                Qt.CheckState.Checked if checked and selectable else Qt.CheckState.Unchecked
            )
            if not selectable:
                name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
            name_item.setData(Qt.ItemDataRole.UserRole, sheet.name)
            self.table.setItem(row, 0, name_item)

            if sheet.is_scenario:
                status = "sheet skenario"
            elif sheet.is_mapping_candidate:
                status = "? perlu mapping"
            else:
                status = "bukan skenario"
            if sheet.name in self.overrides:
                status += " (mapping manual)"
            self.table.setItem(row, 1, QTableWidgetItem(status))
            self.table.setItem(
                row, 2, QTableWidgetItem(str(sheet.header_row or "-"))
            )
            self.table.setItem(row, 3, QTableWidgetItem(str(len(sheet.column_map))))
            self.table.setItem(
                row, 4, QTableWidgetItem(str(sheet.data_row_estimate or "-"))
            )
        self._update_buttons()

    def _current_sheet(self) -> Optional[SheetInspection]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        if item is None:
            return None
        return self._inspection.sheet(item.data(Qt.ItemDataRole.UserRole))

    def _update_buttons(self) -> None:
        sheet = self._current_sheet()
        self.map_button.setEnabled(sheet is not None)
        if sheet is None:
            self.detail.setText("")
            return
        lines: List[str] = []
        if sheet.reason:
            lines.append(sheet.reason)
        if sheet.column_map:
            lines.append(
                "Terpetakan: "
                + ", ".join(
                    f"{field_label(name)}=kol{col}"
                    for name, col in sorted(sheet.column_map.items(), key=lambda kv: kv[1])
                )
            )
        if sheet.unmapped_columns:
            lines.append(
                "Kolom tidak dikenali: "
                + ", ".join(text for _c, text in sheet.unmapped_columns)
            )
        if sheet.metadata:
            lines.append(
                "Metadata: " + ", ".join(f"{k}={v}" for k, v in sheet.metadata.items())
            )
        self.detail.setText("<br>".join(lines))

    def _select_all(self) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is not None and item.flags() & Qt.ItemFlag.ItemIsUserCheckable:
                item.setCheckState(Qt.CheckState.Checked)

    def _clear_all(self) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is not None and item.flags() & Qt.ItemFlag.ItemIsUserCheckable:
                item.setCheckState(Qt.CheckState.Unchecked)

    # -- mapping ----------------------------------------------------------- #

    def _header_choices(self, sheet: SheetInspection) -> List[tuple]:
        if sheet.headers:
            return list(sheet.headers)
        return []

    def _open_mapping(self) -> None:
        sheet = self._current_sheet()
        if sheet is None:
            return
        choices = self._header_choices(sheet)
        if not choices:
            QMessageBox.information(
                self,
                "Tidak ada header",
                f"Sheet '{sheet.name}' tidak punya baris header yang bisa dibaca, "
                "jadi tidak ada kolom untuk dipetakan.",
            )
            return
        dialog = ColumnMappingDialog(
            sheet,
            self._config,
            header_choices=choices,
            current=self.overrides.get(sheet.name),
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        override = dialog.result_override()
        checked = self._checked_names()
        if override is None:
            self.overrides.pop(sheet.name, None)
        else:
            self.overrides[sheet.name] = override
        self._refresh(keep_checked=checked | {sheet.name})

    def _refresh(self, *, keep_checked: Optional[set] = None) -> None:
        payload = {name: ov.to_dict() for name, ov in self.overrides.items()}
        self._inspection = self._reinspect(payload)
        if keep_checked is not None:
            self._preselected = keep_checked
        self._fill_table()

    # -- result ------------------------------------------------------------ #

    def _checked_names(self) -> set:
        names = set()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                names.add(item.data(Qt.ItemDataRole.UserRole))
        return names

    def selected_sheets(self) -> List[str]:
        checked = self._checked_names()
        return [s.name for s in self._inspection.sheets if s.name in checked]

    def _on_accept(self) -> None:
        if not self.selected_sheets():
            QMessageBox.warning(
                self, "Belum ada module dipilih",
                "Pilih minimal satu sheet skenario.",
            )
            return
        self.accept()


# --------------------------------------------------------------------------- #
# New project
# --------------------------------------------------------------------------- #

class NewProjectDialog(QDialog):
    """Ask for the Excel file and the project folder."""

    def __init__(self, *, parent: Optional[QWidget] = None,
                 start_dir: Optional[Path] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Project baru")
        self.resize(640, 220)
        self._start_dir = Path(start_dir) if start_dir else Path.home()

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.excel_edit = QLineEdit()
        excel_row = QHBoxLayout()
        excel_row.addWidget(self.excel_edit)
        browse_excel = QPushButton("Pilih...")
        browse_excel.clicked.connect(self._browse_excel)
        excel_row.addWidget(browse_excel)
        excel_widget = QWidget()
        excel_widget.setLayout(excel_row)
        form.addRow("File skenario (.xlsx):", excel_widget)

        self.folder_edit = QLineEdit()
        folder_row = QHBoxLayout()
        folder_row.addWidget(self.folder_edit)
        browse_folder = QPushButton("Pilih...")
        browse_folder.clicked.connect(self._browse_folder)
        folder_row.addWidget(browse_folder)
        folder_widget = QWidget()
        folder_widget.setLayout(folder_row)
        form.addRow("Folder project:", folder_widget)

        self.name_edit = QLineEdit()
        form.addRow("Nama project:", self.name_edit)
        layout.addLayout(form)

        note = QLabel(
            "Folder project akan berisi <code>project.json</code>, "
            "<code>screenshots/</code> dan <code>output/</code>. "
            "File Excel asli tidak pernah diubah."
        )
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addStretch(1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih file skenario", str(self._start_dir),
            "Excel (*.xlsx *.xlsm)",
        )
        if path:
            self.excel_edit.setText(path)
            if not self.folder_edit.text():
                suggestion = Path(path).parent / (Path(path).stem + " - Result")
                self.folder_edit.setText(str(suggestion))
            if not self.name_edit.text():
                self.name_edit.setText(Path(path).stem)

    def _browse_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Pilih folder project", self.folder_edit.text() or str(self._start_dir)
        )
        if path:
            self.folder_edit.setText(path)

    def _on_accept(self) -> None:
        excel = Path(self.excel_edit.text().strip())
        folder = self.folder_edit.text().strip()
        if not excel.is_file():
            QMessageBox.warning(self, "File belum benar",
                                "Pilih file skenario .xlsx yang ada.")
            return
        if not folder:
            QMessageBox.warning(self, "Folder belum dipilih",
                                "Tentukan folder untuk menyimpan project.")
            return
        self.accept()

    @property
    def excel_path(self) -> Path:
        return Path(self.excel_edit.text().strip())

    @property
    def folder(self) -> Path:
        return Path(self.folder_edit.text().strip())

    @property
    def project_name(self) -> str:
        return self.name_edit.text().strip()


# --------------------------------------------------------------------------- #
# Metadata and module titles
# --------------------------------------------------------------------------- #

class MetadataDialog(QDialog):
    """Edit metadata and the Heading 1 title of each module."""

    def __init__(self, project: Project, *, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Metadata & judul module")
        self.resize(640, 560)
        self._project = project

        layout = QVBoxLayout(self)

        meta_box = QGroupBox("Metadata (kosongkan untuk memakai nilai dari Excel)")
        form = QFormLayout(meta_box)
        parsed = project.scenario_file.merged_metadata() if project.scenario_file else {}
        self._meta_edits: Dict[str, QLineEdit] = {}
        keys = list(METADATA_LABELS) + [k for k in parsed if k not in METADATA_LABELS]
        for key in keys:
            edit = QLineEdit(project.metadata.get(key, ""))
            edit.setPlaceholderText(parsed.get(key, "(tidak ada di Excel)"))
            self._meta_edits[key] = edit
            form.addRow(METADATA_LABELS.get(key, key), edit)
        layout.addWidget(meta_box)

        title_box = QGroupBox("Judul module (Heading 1)")
        title_form = QFormLayout(title_box)
        naming = project.naming() if project.scenario_file else {}
        self._title_edits: Dict[str, QLineEdit] = {}
        for module in (project.scenario_file.modules if project.scenario_file else []):
            edit = QLineEdit(project.module_titles.get(module.name, ""))
            item = naming.get(module.name)
            edit.setPlaceholderText(item.title if item else module.title)
            self._title_edits[module.name] = edit
            title_form.addRow(module.name, edit)
        layout.addWidget(title_box)
        layout.addStretch(1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def apply(self) -> None:
        for key, edit in self._meta_edits.items():
            value = edit.text().strip()
            if value:
                self._project.metadata[key] = value
            else:
                self._project.metadata.pop(key, None)
        for name, edit in self._title_edits.items():
            value = edit.text().strip()
            if value:
                self._project.module_titles[name] = value
            else:
                self._project.module_titles.pop(name, None)


# --------------------------------------------------------------------------- #
# Export
# --------------------------------------------------------------------------- #

class ExportDialog(QDialog):
    """Choose per-module or combined output, plus the optional template."""

    def __init__(self, project: Project, *, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Generate dokumen Word")
        self.resize(700, 560)
        self._project = project

        layout = QVBoxLayout(self)

        mode_box = QGroupBox("Bentuk output")
        mode_layout = QVBoxLayout(mode_box)
        self.per_module = QRadioButton("Satu file .docx per module")
        self.combined = QRadioButton("Satu file gabungan untuk semua module terpilih")
        if project.output_mode == "combined":
            self.combined.setChecked(True)
        else:
            self.per_module.setChecked(True)
        mode_layout.addWidget(self.per_module)
        mode_layout.addWidget(self.combined)
        layout.addWidget(mode_box)

        module_box = QGroupBox("Module yang diexport")
        module_layout = QVBoxLayout(module_box)
        self._module_checks: Dict[str, QCheckBox] = {}
        for module in (project.scenario_file.modules if project.scenario_file else []):
            check = QCheckBox(module.name)
            check.setChecked(True)
            self._module_checks[module.name] = check
            module_layout.addWidget(check)
        layout.addWidget(module_box)

        template_row = QHBoxLayout()
        template_row.addWidget(QLabel("Template style (.docx, opsional):"))
        self.template_edit = QLineEdit(project.template or "")
        template_row.addWidget(self.template_edit)
        browse = QPushButton("Pilih...")
        browse.clicked.connect(self._browse_template)
        template_row.addWidget(browse)
        clear = QPushButton("Hapus")
        clear.clicked.connect(lambda: self.template_edit.setText(""))
        template_row.addWidget(clear)
        layout.addLayout(template_row)

        layout.addWidget(QLabel("Ringkasan sebelum export:"))
        self.summary = QPlainTextEdit()
        self.summary.setReadOnly(True)
        layout.addWidget(self.summary, 1)

        for check in self._module_checks.values():
            check.stateChanged.connect(lambda _state: self.refresh_summary())
        self.refresh_summary()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Generate")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih template .docx", self.template_edit.text() or "",
            "Word (*.docx)",
        )
        if path:
            self.template_edit.setText(path)

    def selected_modules(self) -> List[str]:
        return [name for name, check in self._module_checks.items() if check.isChecked()]

    def refresh_summary(self) -> None:
        modules = self.selected_modules()
        if not modules:
            self.summary.setPlainText("Belum ada module dipilih.")
            return
        lines: List[str] = []
        for report in self._project.summarize(modules=modules):
            lines.append(f"[{report.title}]")
            lines.append(f"    test case  : {report.test_case_count}")
            lines.append(f"    screenshot : {report.screenshot_count}")
            if report.without_screenshot:
                lines.append(
                    f"    belum ada screenshot : {len(report.without_screenshot)} test case"
                )
                for label in report.without_screenshot:
                    lines.append(f"        - {label}")
            if report.missing_files:
                lines.append(f"    file hilang : {len(report.missing_files)}")
                for name in report.missing_files:
                    lines.append(f"        ! {name}")
            lines.append("")
        self.summary.setPlainText("\n".join(lines))

    def _on_accept(self) -> None:
        if not self.selected_modules():
            QMessageBox.warning(self, "Belum ada module",
                                "Pilih minimal satu module untuk diexport.")
            return
        self.accept()

    def apply(self) -> None:
        self._project.output_mode = "combined" if self.combined.isChecked() else "per_module"
        text = self.template_edit.text().strip()
        self._project.template = text or None


# --------------------------------------------------------------------------- #
# Orphan screenshots
# --------------------------------------------------------------------------- #

class OrphanDialog(QDialog):
    """Re-attach or discard screenshots whose test case disappeared."""

    def __init__(self, project: Project, *, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Screenshot orphan")
        self.resize(760, 420)
        self._project = project

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Screenshot berikut dulunya menempel pada test case yang sekarang sudah "
            "tidak ada di file Excel. File-nya tidak dihapus.<br>"
            "Kalau test case-nya muncul lagi saat Re-import, screenshot ini "
            "dipulihkan otomatis."
        ))

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(("Test case lama", "Jumlah screenshot"))
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table, 1)

        row = QHBoxLayout()
        row.addWidget(QLabel("Pindahkan ke:"))
        self.target = QComboBox()
        for case in project.test_cases:
            self.target.addItem(f"{case.module} - {case.display_label}", case.uid)
        row.addWidget(self.target, 1)
        restore = QPushButton("Pindahkan")
        restore.clicked.connect(self._restore)
        row.addWidget(restore)
        discard = QPushButton("Buang (hapus file)")
        discard.clicked.connect(self._discard)
        row.addWidget(discard)
        layout.addLayout(row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

        self._fill()

    def _fill(self) -> None:
        uids = self._project.orphan_uids
        self.table.setRowCount(len(uids))
        for row, uid in enumerate(uids):
            item = QTableWidgetItem(uid)
            item.setData(Qt.ItemDataRole.UserRole, uid)
            self.table.setItem(row, 0, item)
            self.table.setItem(
                row, 1, QTableWidgetItem(str(len(self._project.orphans.get(uid, []))))
            )

    def _current_uid(self) -> Optional[str]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _restore(self) -> None:
        uid = self._current_uid()
        target = self.target.currentData()
        if not uid or not target:
            return
        self._project.restore_orphan(uid, target)
        self._fill()

    def _discard(self) -> None:
        uid = self._current_uid()
        if not uid:
            return
        answer = QMessageBox.question(
            self, "Buang screenshot orphan",
            f"Hapus permanen semua screenshot orphan dari '{uid}'?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._project.discard_orphans(uid, delete_files=True)
        self._fill()


_STATUS_LABELS = {
    EDITED: "Berubah isinya",
    SUSPECT: "Bergeser / ter-renumber",
    MISSING: "Hilang, ada pengganti?",
    REMOVED: "Hilang",
}


class ScenarioChangesDialog(QDialog):
    """Review test cases whose scenario changed after evidence was captured.

    Nothing is moved until the user presses *Terapkan*. Ticked remap rows are
    moved to the chosen test case; ticked "edited" rows are marked as checked.
    """

    COL_APPLY, COL_STATUS, COL_OLD, COL_TARGET, COL_SCORE, COL_DETAIL = range(6)

    def __init__(
        self,
        project: Project,
        changes: ScenarioChanges,
        *,
        allow_undo: bool = True,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Review perubahan skenario")
        self.resize(1000, 480)
        self._project = project
        self._rows: List[CaseChange] = list(changes.changes)
        self.undo_requested = False

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "File skenario berubah setelah bukti diambil. Tidak ada yang dipindahkan "
            "sebelum Anda menekan <b>Terapkan</b>.<br>"
            "Centang <i>Pindahkan</i> untuk memindahkan screenshot dan hasil ke test case "
            "baru; centang <i>Sudah dicek</i> kalau bukti lama masih berlaku."
        ))

        self.table = QTableWidget(len(self._rows), 6)
        self.table.setHorizontalHeaderLabels(
            ("Terapkan", "Status", "Test case lama", "Pindahkan ke", "Skor", "Detail")
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(self.COL_DETAIL, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)

        self._combos: Dict[int, QComboBox] = {}
        for row, change in enumerate(self._rows):
            self._fill_row(row, change)
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(self.COL_TARGET, 260)

        buttons = QDialogButtonBox()
        apply_button = buttons.addButton("Terapkan", QDialogButtonBox.ButtonRole.AcceptRole)
        apply_button.setDefault(True)
        buttons.addButton("Nanti saja", QDialogButtonBox.ButtonRole.RejectRole)
        if allow_undo and project.latest_backup() is not None:
            undo = buttons.addButton(
                "Undo re-import", QDialogButtonBox.ButtonRole.ActionRole
            )
            undo.clicked.connect(self._undo)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # -- table ------------------------------------------------------------- #

    def _fill_row(self, row: int, change: CaseChange) -> None:
        snapshot = self._project.snapshots.get(change.uid)
        old_label = change.uid + (f"  ({snapshot.test_case})" if snapshot else "")
        remappable = change.status in (SUSPECT, MISSING)
        checkable = remappable or change.status == EDITED

        apply_item = QTableWidgetItem("Pindahkan" if remappable else "Sudah dicek")
        if checkable:
            apply_item.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled
            )
            apply_item.setCheckState(
                Qt.CheckState.Checked if change.preselect else Qt.CheckState.Unchecked
            )
        else:
            apply_item.setText("Jadi orphan")
            apply_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        self.table.setItem(row, self.COL_APPLY, apply_item)
        self.table.setItem(row, self.COL_STATUS, QTableWidgetItem(_STATUS_LABELS[change.status]))
        self.table.setItem(row, self.COL_OLD, QTableWidgetItem(old_label))
        score = f"{change.score:.0%}" if change.score else ""
        self.table.setItem(row, self.COL_SCORE, QTableWidgetItem(score))
        detail = QTableWidgetItem("; ".join(change.details))
        detail.setToolTip("\n".join(change.details))
        self.table.setItem(row, self.COL_DETAIL, detail)

        if remappable:
            combo = QComboBox()
            combo.addItem("(jangan pindahkan)", None)
            for case in self._project.test_cases:
                combo.addItem(f"{case.module} - {case.display_label}", case.uid)
            index = combo.findData(change.suggested_uid)
            combo.setCurrentIndex(max(index, 0))
            self.table.setCellWidget(row, self.COL_TARGET, combo)
            self._combos[row] = combo
        elif change.status == EDITED:
            self.table.setItem(row, self.COL_TARGET, QTableWidgetItem("(tetap di test case ini)"))

    def _checked(self, row: int) -> bool:
        item = self.table.item(row, self.COL_APPLY)
        return bool(item and item.flags() & Qt.ItemFlag.ItemIsUserCheckable
                    and item.checkState() == Qt.CheckState.Checked)

    # -- result ------------------------------------------------------------ #

    def remap_selection(self):
        """``(mapping, step_maps)`` for the rows the user ticked."""
        mapping: Dict[str, str] = {}
        step_maps: Dict[str, Dict[str, str]] = {}
        for row, change in enumerate(self._rows):
            combo = self._combos.get(row)
            if combo is None or not self._checked(row):
                continue
            target = combo.currentData()
            if not target or target == change.uid:
                continue
            mapping[change.uid] = target
            if target == change.suggested_uid:
                step_maps[change.uid] = dict(change.step_map)
            else:
                snapshot = self._project.snapshots.get(change.uid)
                case = self._project.test_case(target)
                if snapshot is not None and case is not None:
                    step_maps[change.uid] = match_steps(snapshot.steps, case.steps)
        return mapping, step_maps

    def acknowledged(self) -> List[str]:
        return [
            change.uid for row, change in enumerate(self._rows)
            if change.status == EDITED and self._checked(row)
        ]

    def apply(self) -> None:
        mapping, step_maps = self.remap_selection()
        if mapping:
            self._project.apply_remap(mapping, step_maps, save=False)
        for uid in self.acknowledged():
            self._project.acknowledge(uid, save=False)
        self._project.save()

    def _undo(self) -> None:
        answer = QMessageBox.question(
            self, "Undo re-import",
            "Kembalikan project.json ke kondisi sebelum re-import terakhir?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.undo_requested = True
        self.reject()
