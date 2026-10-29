"""Editing of the execution results that get written back to Excel."""

from __future__ import annotations

import datetime as dt
from typing import Dict, List, Optional

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..excel_writer import RESULT_FIELDS
from .icons import icon
from ..project import Project, step_key

#: Table layout: (field, heading, editable).
COLUMNS = (
    ("step", "Step", False),
    ("test_step", "Test Step", False),
    ("actual_results", "Actual Results", True),
    ("pass_fail", "Pass/Fail", True),
    ("date_executed", "Date Executed", True),
    ("test_executed_by", "Test Executed by", True),
)

PASS_FAIL_CHOICES = ("", "P", "F", "N/A")


class BulkResultDialog(QDialog):
    """Fill the same result into every step of a test case."""

    def __init__(self, *, executed_by: str = "", parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Isi hasil untuk semua step")
        self.resize(480, 260)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Nilai yang dikosongkan tidak diubah, jadi Anda bisa mengisi "
            "Pass/Fail saja tanpa menimpa Actual Results."
        ))
        form = QFormLayout()

        self.actual = QLineEdit()
        form.addRow("Actual Results:", self.actual)

        self.pass_fail = QComboBox()
        self.pass_fail.addItems(PASS_FAIL_CHOICES)
        self.pass_fail.setEditable(True)
        form.addRow("Pass/Fail:", self.pass_fail)

        date_row = QHBoxLayout()
        self.use_date = QPushButton("Pakai tanggal")
        self.use_date.setCheckable(True)
        self.use_date.setChecked(True)
        self.date = QDateEdit(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat("yyyy-MM-dd")
        date_row.addWidget(self.date, 1)
        date_row.addWidget(self.use_date)
        date_widget = QWidget()
        date_widget.setLayout(date_row)
        form.addRow("Date Executed:", date_widget)

        self.executed_by = QLineEdit(executed_by)
        form.addRow("Test Executed by:", self.executed_by)
        layout.addLayout(form)
        layout.addStretch(1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> Dict[str, str]:
        """Only the fields the user actually filled in."""
        data: Dict[str, str] = {}
        if self.actual.text().strip():
            data["actual_results"] = self.actual.text().strip()
        if self.pass_fail.currentText().strip():
            data["pass_fail"] = self.pass_fail.currentText().strip()
        if self.use_date.isChecked():
            data["date_executed"] = self.date.date().toString("yyyy-MM-dd")
        if self.executed_by.text().strip():
            data["test_executed_by"] = self.executed_by.text().strip()
        return data


class ResultsPanel(QWidget):
    """Per-step Actual Results / Pass-Fail / Date / Executed by."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._project: Optional[Project] = None
        self._uid: Optional[str] = None
        self._keys: List[str] = []
        self._loading = False

        layout = QVBoxLayout(self)
        self.header = QLabel("Pilih sebuah test case di tree.")
        self.header.setWordWrap(True)
        layout.addWidget(self.header)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels([title for _f, title, _e in COLUMNS])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        self.btn_bulk = QPushButton("Isi semua step...")
        self.btn_bulk.clicked.connect(self._fill_all)
        buttons.addWidget(self.btn_bulk)
        self.btn_today = QPushButton("Tandai semua Pass hari ini")
        self.btn_today.clicked.connect(self._mark_all_pass)
        buttons.addWidget(self.btn_today)
        self.btn_clear = QPushButton("Kosongkan hasil")
        self.btn_clear.clicked.connect(self._clear_all)
        buttons.addWidget(self.btn_clear)
        buttons.addStretch(1)
        layout.addLayout(buttons)
        self.apply_icons()

        self._update_buttons()

    # -- filling ----------------------------------------------------------- #

    def set_test_case(self, project: Optional[Project], uid: Optional[str]) -> None:
        self._project = project
        self._uid = uid
        self.refresh()

    def refresh(self) -> None:
        self._loading = True
        try:
            self.table.setRowCount(0)
            self._keys = []
            project, uid = self._project, self._uid
            if project is None or uid is None:
                self.header.setText("Pilih sebuah test case di tree.")
                return
            case = project.test_case(uid)
            if case is None:
                self.header.setText(f"Test case '{uid}' tidak ada di skenario.")
                return
            filled = project.steps_with_results(uid)
            self.header.setText(
                f"<b>{case.display_label}</b> - {len(case.steps)} step, "
                f"{filled} sudah ada hasilnya.<br>"
                "Nilai ini ditulis ke <i>salinan</i> file Excel, bukan file aslinya."
            )
            self.table.setRowCount(len(case.steps))
            for row, step in enumerate(case.steps, start=0):
                key = step_key(step, row + 1)
                self._keys.append(key)
                stored = project.result_for(uid, key)
                values = {
                    "step": key,
                    "test_step": step.test_step.replace("\n", " "),
                    "actual_results": stored.get("actual_results", step.actual_results),
                    "pass_fail": stored.get("pass_fail", step.pass_fail),
                    "date_executed": stored.get("date_executed", step.date_executed),
                    "test_executed_by": stored.get(
                        "test_executed_by", step.test_executed_by
                    ),
                }
                for column, (field, _title, editable) in enumerate(COLUMNS):
                    item = QTableWidgetItem(values.get(field, ""))
                    if not editable:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    elif field not in stored and values.get(field):
                        # Came from the workbook, not from this project yet.
                        item.setToolTip("Nilai dari file Excel; edit untuk menimpanya.")
                    self.table.setItem(row, column, item)
            self.table.resizeColumnsToContents()
            self.table.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.ResizeMode.Stretch
            )
            self.table.horizontalHeader().setSectionResizeMode(
                2, QHeaderView.ResizeMode.Stretch
            )
        finally:
            self._loading = False
            self._update_buttons()

    def _update_buttons(self) -> None:
        enabled = self._project is not None and self._uid is not None and bool(self._keys)
        for button in (self.btn_bulk, self.btn_today, self.btn_clear):
            button.setEnabled(enabled)

    # -- editing ----------------------------------------------------------- #

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._loading or self._project is None or self._uid is None:
            return
        field, _title, editable = COLUMNS[item.column()]
        if not editable:
            return
        row = item.row()
        if not 0 <= row < len(self._keys):
            return
        self._project.set_result(self._uid, self._keys[row], {field: item.text()})
        self._refresh_header()

    def _refresh_header(self) -> None:
        project, uid = self._project, self._uid
        if project is None or uid is None:
            return
        case = project.test_case(uid)
        if case is None:
            return
        self.header.setText(
            f"<b>{case.display_label}</b> - {len(case.steps)} step, "
            f"{project.steps_with_results(uid)} sudah ada hasilnya.<br>"
            "Nilai ini ditulis ke <i>salinan</i> file Excel, bukan file aslinya."
        )

    def _apply_to_all(self, values: Dict[str, str]) -> None:
        if self._project is None or self._uid is None or not values:
            return
        self._project.set_results_for_test_case(self._uid, values)
        self.refresh()

    def apply_icons(self) -> None:
        """Icons follow the light/dark theme; the text stays (the actions are not obvious)."""
        self.btn_bulk.setIcon(icon("checks"))
        self.btn_today.setIcon(icon("done"))
        self.btn_clear.setIcon(icon("delete"))

    def _fill_all(self) -> None:
        if self._project is None:
            return
        last_by = self._project.effective_metadata().get("created_by", "")
        for steps in self._project.results.values():
            for stored in steps.values():
                if stored.get("test_executed_by"):
                    last_by = stored["test_executed_by"]
        dialog = BulkResultDialog(executed_by=last_by, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self._apply_to_all(dialog.values())

    def _mark_all_pass(self) -> None:
        if self._project is None:
            return
        executed_by = self._project.effective_metadata().get("created_by", "")
        self._apply_to_all({
            "pass_fail": "P",
            "date_executed": dt.date.today().isoformat(),
            "test_executed_by": executed_by,
        })

    def _clear_all(self) -> None:
        if self._project is None or self._uid is None:
            return
        for key in self._keys:
            self._project.clear_result(self._uid, key, save=False)
        self._project.save()
        self.refresh()
