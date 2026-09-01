import os
import tempfile

from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from openpyxl import load_workbook


class ExcelEditorWindow(QWidget):
    """
    Excel editor window.

    The backend remains the source of truth for every structural
    operation. The local workbook/table is updated only after the
    backend operation succeeds.

    Supported operations:
        - Edit cells
        - Save cell changes
        - Add row
        - Delete row
        - Add column
        - Rename column
        - Delete column
        - Create sheet
        - Rename sheet
        - Delete sheet
        - Refresh workbook
    """

    def __init__(
        self,
        file_id: int,
        filename: str,
        api_client,
    ):
        super().__init__()

        self.file_id = file_id
        self.filename = filename
        self.api_client = api_client

        self.workbook = None
        self.temp_file_path = None

        # Unsaved cell changes only.
        # {(sheet_name, "A1"): value}
        self.pending_changes = {}

        # Prevents cellChanged from treating programmatic table
        # updates as user edits.
        self._loading_table = False

        self.setWindowTitle(
            f"Excel Editor - {self.filename}"
        )
        self.setMinimumSize(1200, 700)

        self.setup_ui()
        self.load_excel_file()

        # Buttons
        self.add_row_button.clicked.connect(self.add_row)
        self.delete_row_button.clicked.connect(self.delete_row)
        self.add_column_button.clicked.connect(self.add_column)
        self.rename_column_button.clicked.connect(self.rename_column)
        self.delete_column_button.clicked.connect(self.delete_column)
        self.save_button.clicked.connect(self.save_changes)

        # Sheet management
        self.create_sheet_button.clicked.connect(self.create_sheet)
        self.rename_sheet_button.clicked.connect(self.rename_sheet)
        self.delete_sheet_button.clicked.connect(self.delete_sheet)

    # =========================================================
    # UI
    # =========================================================

    def setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        header_layout = QHBoxLayout()

        title = QLabel(
            f"Excel Editor - {self.filename}"
        )
        title.setStyleSheet(
            """
            QLabel {
                font-size: 24px;
                font-weight: bold;
            }
            """
        )

        header_layout.addWidget(title)
        header_layout.addStretch()

        header_layout.addWidget(QLabel("Sheet:"))

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(180)
        self.sheet_combo.currentTextChanged.connect(
            self.load_sheet
        )
        header_layout.addWidget(self.sheet_combo)

        self.create_sheet_button = QPushButton("New Sheet")
        self.rename_sheet_button = QPushButton("Rename Sheet")
        self.delete_sheet_button = QPushButton("Delete Sheet")
        self.refresh_button = QPushButton("Refresh")

        for button in (
            self.create_sheet_button,
            self.rename_sheet_button,
            self.delete_sheet_button,
            self.refresh_button,
        ):
            button.setMinimumHeight(38)

        header_layout.addWidget(self.create_sheet_button)
        header_layout.addWidget(self.rename_sheet_button)
        header_layout.addWidget(self.delete_sheet_button)
        header_layout.addWidget(self.refresh_button)

        self.refresh_button.clicked.connect(
            self.load_excel_file
        )

        main_layout.addLayout(header_layout)

        # =====================================================
        # Table
        # =====================================================

        self.table = QTableWidget()

        self.table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked
            | QTableWidget.EditTrigger.EditKeyPressed
        )

        self.table.setAlternatingRowColors(True)

        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectItems
        )

        self.table.cellChanged.connect(
            self.handle_cell_changed
        )

        self.table.horizontalHeader().setStretchLastSection(
            True
        )

        main_layout.addWidget(self.table)

        # =====================================================
        # Bottom buttons
        # =====================================================

        bottom_layout = QHBoxLayout()

        self.add_row_button = QPushButton("Add Row")
        self.delete_row_button = QPushButton("Delete Row")
        self.add_column_button = QPushButton("Add Column")
        self.rename_column_button = QPushButton("Rename Column")
        self.delete_column_button = QPushButton("Delete Column")
        self.save_button = QPushButton("Save")

        for button in (
            self.add_row_button,
            self.delete_row_button,
            self.add_column_button,
            self.rename_column_button,
            self.delete_column_button,
            self.save_button,
        ):
            button.setMinimumHeight(40)

        bottom_layout.addWidget(self.add_row_button)
        bottom_layout.addWidget(self.delete_row_button)
        bottom_layout.addWidget(self.add_column_button)
        bottom_layout.addWidget(self.rename_column_button)
        bottom_layout.addWidget(self.delete_column_button)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.save_button)

        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

    # =========================================================
    # Helpers
    # =========================================================

    @staticmethod
    def column_letter(column_number: int) -> str:
        result = ""

        while column_number > 0:
            column_number, remainder = divmod(
                column_number - 1,
                26,
            )
            result = (
                chr(65 + remainder)
                + result
            )

        return result

    def selected_sheet(self) -> str:
        return self.sheet_combo.currentText().strip()

    def set_busy(self, button, text: str, busy: bool):
        button.setEnabled(not busy)
        button.setText(text if busy else button.property("normal_text"))

    def clear_pending_for_sheet(self, sheet_name: str):
        self.pending_changes = {
            key: value
            for key, value in self.pending_changes.items()
            if key[0] != sheet_name
        }

    def rebuild_sheet_combo(
        self,
        selected_sheet: str | None = None,
    ):
        if not self.workbook:
            return

        names = self.workbook.sheetnames

        self.sheet_combo.blockSignals(True)
        try:
            self.sheet_combo.clear()
            self.sheet_combo.addItems(names)

            if selected_sheet in names:
                self.sheet_combo.setCurrentText(selected_sheet)
            elif names:
                self.sheet_combo.setCurrentIndex(0)
        finally:
            self.sheet_combo.blockSignals(False)

    # =========================================================
    # Download + Load Workbook
    # =========================================================

    def load_excel_file(self):
        try:
            self.refresh_button.setEnabled(False)
            self.refresh_button.setText("Loading...")

            self.pending_changes.clear()

            if self.temp_file_path:
                try:
                    os.remove(self.temp_file_path)
                except OSError:
                    pass

            fd, self.temp_file_path = tempfile.mkstemp(
                suffix=".xlsx"
            )
            os.close(fd)

            self.api_client.download_file(
                file_id=self.file_id,
                save_path=self.temp_file_path,
            )

            self.workbook = load_workbook(
                self.temp_file_path
            )

            current_sheet = self.selected_sheet()

            self.rebuild_sheet_combo(
                current_sheet
                if current_sheet in self.workbook.sheetnames
                else None
            )

            if self.workbook.sheetnames:
                self.load_sheet(
                    self.sheet_combo.currentText()
                )
            else:
                self.table.clear()
                self.table.setRowCount(0)
                self.table.setColumnCount(0)

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Excel File",
                str(exc),
            )

        finally:
            self.refresh_button.setEnabled(True)
            self.refresh_button.setText("Refresh")

    # =========================================================
    # Load Sheet
    # =========================================================

    def load_sheet(self, sheet_name: str):
        if not self.workbook or not sheet_name:
            return

        if sheet_name not in self.workbook.sheetnames:
            return

        self._loading_table = True
        self.table.blockSignals(True)

        try:
            worksheet = self.workbook[sheet_name]

            max_row = worksheet.max_row
            max_column = worksheet.max_column

            # A newly created completely empty sheet has
            # max_row/max_column equal to 1 in some openpyxl
            # situations. We still show one editable cell.
            if max_row < 1:
                max_row = 1

            if max_column < 1:
                max_column = 1

            self.table.clear()
            self.table.setRowCount(max_row)
            self.table.setColumnCount(max_column)

            headers = [
                self.column_letter(index + 1)
                for index in range(max_column)
            ]

            self.table.setHorizontalHeaderLabels(headers)

            for row_index in range(max_row):
                excel_row = row_index + 1

                for column_index in range(max_column):
                    excel_column = column_index + 1

                    value = worksheet.cell(
                        row=excel_row,
                        column=excel_column,
                    ).value

                    text = "" if value is None else str(value)

                    self.table.setItem(
                        row_index,
                        column_index,
                        QTableWidgetItem(text),
                    )

            self.table.resizeColumnsToContents()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Sheet",
                str(exc),
            )

        finally:
            self.table.blockSignals(False)
            self._loading_table = False

    # =========================================================
    # Cell Editing
    # =========================================================

    def handle_cell_changed(
        self,
        row: int,
        column: int,
    ):
        if self._loading_table:
            return

        if not self.workbook:
            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        item = self.table.item(row, column)

        value = None if item is None else item.text()

        cell = (
            self.column_letter(column + 1)
            + str(row + 1)
        )

        self.pending_changes[
            (sheet_name, cell)
        ] = value

        try:
            worksheet = self.workbook[sheet_name]
            worksheet[cell] = value
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Update Cell",
                str(exc),
            )

    # =========================================================
    # Save Cell Changes
    # =========================================================

    def save_changes(self):
        if not self.pending_changes:
            QMessageBox.information(
                self,
                "Save",
                "There are no changes to save.",
            )
            return

        self.save_button.setEnabled(False)
        self.save_button.setText("Saving...")

        try:
            successful_changes = []

            for (
                sheet_name,
                cell,
            ), value in list(
                self.pending_changes.items()
            ):
                self.api_client.update_cell(
                    file_id=self.file_id,
                    sheet_name=sheet_name,
                    cell=cell,
                    value=value,
                )

                successful_changes.append(
                    (sheet_name, cell)
                )

            for change in successful_changes:
                self.pending_changes.pop(
                    change,
                    None,
                )

            QMessageBox.information(
                self,
                "Save Successful",
                "Changes saved successfully.",
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Save Failed",
                (
                    "Some changes could not be saved.\n\n"
                    f"{exc}"
                ),
            )

        finally:
            self.save_button.setEnabled(True)
            self.save_button.setText("Save")

    # =========================================================
    # Add Row
    # =========================================================

    def add_row(self):
        sheet_name = self.selected_sheet()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Add Row",
                "Please select a sheet first.",
            )
            return

        column_count = max(
            self.table.columnCount(),
            1,
        )

        # IMPORTANT:
        # Use empty strings instead of a list containing only None.
        # This makes the blank row a real row in openpyxl and prevents
        # max_row from collapsing back to the previous row.
        row_data = [""] * column_count

        self.add_row_button.setEnabled(False)
        self.add_row_button.setText("Adding...")

        try:
            result = self.api_client.add_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_data=row_data,
            )

            # The backend returns the actual Excel row number.
            # Use it instead of assuming table.rowCount() is the
            # correct Excel row.
            backend_row_number = result.get("row_number")

            if backend_row_number is None:
                backend_row_number = (
                    self.table.rowCount() + 1
                )

            # Update local workbook first.
            worksheet = self.workbook[sheet_name]

            # The backend has already added the row. The local
            # workbook is only a UI-side copy, so append one row.
            worksheet.append(row_data)

            # If openpyxl's calculated row differs from the backend
            # row, explicitly write the cells at the backend row.
            # This keeps the local workbook aligned with the server.
            local_row_number = worksheet.max_row

            if local_row_number != backend_row_number:
                for column_number, value in enumerate(
                    row_data,
                    start=1,
                ):
                    worksheet.cell(
                        row=backend_row_number,
                        column=column_number,
                        value=value,
                    )

            # Update the table without triggering cellChanged.
            self._loading_table = True
            self.table.blockSignals(True)

            try:
                target_row = backend_row_number - 1

                while self.table.rowCount() <= target_row:
                    self.table.insertRow(
                        self.table.rowCount()
                    )

                for column in range(column_count):
                    self.table.setItem(
                        target_row,
                        column,
                        QTableWidgetItem(""),
                    )

            finally:
                self.table.blockSignals(False)
                self._loading_table = False

            self.table.resizeColumnsToContents()
            self.table.scrollToBottom()

            QMessageBox.information(
                self,
                "Row Added",
                result.get(
                    "message",
                    "Row added successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Add Row Failed",
                str(exc),
            )

        finally:
            self.add_row_button.setEnabled(True)
            self.add_row_button.setText("Add Row")

    # =========================================================
    # Delete Row
    # =========================================================

    def delete_row(self):
        current_row = self.table.currentRow()

        if current_row < 0:
            QMessageBox.warning(
                self,
                "Delete Row",
                "Please select a row to delete.",
            )
            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Delete Row",
                "Please select a sheet first.",
            )
            return

        # QTableWidget is 0-based.
        # Excel/openpyxl is 1-based.
        row_number = current_row + 1

        answer = QMessageBox.question(
            self,
            "Delete Row",
            (
                f"Are you sure you want to delete "
                f"row {row_number}?"
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if answer != QMessageBox.StandardButton.Yes:
            return

        self.delete_row_button.setEnabled(False)
        self.delete_row_button.setText("Deleting...")

        try:
            result = self.api_client.delete_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_number=row_number,
            )

            # Backend succeeded. Now update the local workbook/table.
            worksheet = self.workbook[sheet_name]

            worksheet.delete_rows(
                row_number,
                1,
            )

            self._loading_table = True
            self.table.blockSignals(True)

            try:
                self.table.removeRow(current_row)
            finally:
                self.table.blockSignals(False)
                self._loading_table = False

            # Row deletion changes Excel coordinates for all rows below
            # the deleted row. Pending cell changes must therefore be
            # rebuilt with shifted row numbers.
            self._shift_pending_changes_after_row_delete(
                sheet_name,
                row_number,
            )

            QMessageBox.information(
                self,
                "Row Deleted",
                result.get(
                    "message",
                    "Row deleted successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Delete Row Failed",
                str(exc),
            )

        finally:
            self.delete_row_button.setEnabled(True)
            self.delete_row_button.setText("Delete Row")

    def _shift_pending_changes_after_row_delete(
        self,
        sheet_name: str,
        deleted_row: int,
    ):
        updated = {}

        for (
            change_sheet,
            cell,
        ), value in self.pending_changes.items():

            if change_sheet != sheet_name:
                updated[(change_sheet, cell)] = value
                continue

            letters = ""
            digits = ""

            for char in cell:
                if char.isalpha():
                    letters += char
                elif char.isdigit():
                    digits += char

            if not digits:
                updated[(change_sheet, cell)] = value
                continue

            row_number = int(digits)

            if row_number == deleted_row:
                # This pending change belonged to the deleted row.
                continue

            if row_number > deleted_row:
                row_number -= 1

            updated[
                (
                    change_sheet,
                    f"{letters}{row_number}",
                )
            ] = value

        self.pending_changes = updated

    # =========================================================
    # Add Column
    # =========================================================

    def add_column(self):
        sheet_name = self.selected_sheet()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Add Column",
                "Please select a sheet first.",
            )
            return

        column_number = self.table.columnCount() + 1

        self.add_column_button.setEnabled(False)
        self.add_column_button.setText("Adding...")

        try:
            result = self.api_client.add_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
            )

            worksheet = self.workbook[sheet_name]

            # Backend succeeded, so update local workbook.
            worksheet.insert_cols(
                column_number,
                1,
            )

            # Update table directly. Do not reload the workbook from
            # the server here, because that would cause the visible
            # table to jump/reload.
            self._loading_table = True
            self.table.blockSignals(True)

            try:
                self.table.insertColumn(
                    column_number - 1
                )

                self.table.setHorizontalHeaderItem(
                    column_number - 1,
                    QTableWidgetItem(
                        self.column_letter(column_number)
                    ),
                )

                for row in range(
                    self.table.rowCount()
                ):
                    self.table.setItem(
                        row,
                        column_number - 1,
                        QTableWidgetItem(""),
                    )

            finally:
                self.table.blockSignals(False)
                self._loading_table = False

            self._shift_pending_changes_after_column_insert(
                sheet_name,
                column_number,
            )

            self.table.resizeColumnsToContents()

            QMessageBox.information(
                self,
                "Column Added",
                result.get(
                    "message",
                    "Column added successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Add Column Failed",
                str(exc),
            )

        finally:
            self.add_column_button.setEnabled(True)
            self.add_column_button.setText("Add Column")

    def _shift_pending_changes_after_column_insert(
        self,
        sheet_name: str,
        inserted_column: int,
    ):
        updated = {}

        for (
            change_sheet,
            cell,
        ), value in self.pending_changes.items():

            if change_sheet != sheet_name:
                updated[(change_sheet, cell)] = value
                continue

            letters = ""
            digits = ""

            for char in cell:
                if char.isalpha():
                    letters += char
                elif char.isdigit():
                    digits += char

            if not letters or not digits:
                updated[(change_sheet, cell)] = value
                continue

            old_column = self.column_number(letters)

            if old_column >= inserted_column:
                old_column += 1

            new_cell = (
                self.column_letter(old_column)
                + digits
            )

            updated[
                (change_sheet, new_cell)
            ] = value

        self.pending_changes = updated

    # =========================================================
    # Rename Column
    # =========================================================

    def rename_column(self):
        current_column = self.table.currentColumn()

        if current_column < 0:
            QMessageBox.warning(
                self,
                "Rename Column",
                "Please select a column first.",
            )
            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        column_number = current_column + 1
        column_letter = self.column_letter(column_number)

        header_item = self.table.horizontalHeaderItem(
            current_column
        )

        current_name = (
            header_item.text()
            if header_item
            else column_letter
        )

        new_name, ok = QInputDialog.getText(
            self,
            "Rename Column",
            (
                f"Enter new name for "
                f"column {column_letter}:"
            ),
            text=current_name,
        )

        if not ok:
            return

        new_name = new_name.strip()

        if not new_name:
            QMessageBox.warning(
                self,
                "Rename Column",
                "Column name cannot be empty.",
            )
            return

        if new_name == current_name:
            QMessageBox.information(
                self,
                "Rename Column",
                "The column name has not changed.",
            )
            return

        self.rename_column_button.setEnabled(False)
        self.rename_column_button.setText("Renaming...")

        try:
            result = self.api_client.update_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
                column_name=new_name,
            )

            worksheet = self.workbook[sheet_name]

            worksheet.cell(
                row=1,
                column=column_number,
                value=new_name,
            )

            self._loading_table = True
            self.table.blockSignals(True)

            try:
                self.table.setHorizontalHeaderItem(
                    current_column,
                    QTableWidgetItem(new_name),
                )
            finally:
                self.table.blockSignals(False)
                self._loading_table = False

            QMessageBox.information(
                self,
                "Column Renamed",
                result.get(
                    "message",
                    "Column renamed successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Rename Column Failed",
                str(exc),
            )

        finally:
            self.rename_column_button.setEnabled(True)
            self.rename_column_button.setText("Rename Column")

    # =========================================================
    # Delete Column
    # =========================================================

    def delete_column(self):
        current_column = self.table.currentColumn()

        if current_column < 0:
            QMessageBox.warning(
                self,
                "Delete Column",
                "Please select a column first.",
            )
            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        column_number = current_column + 1
        column_letter = self.column_letter(column_number)

        answer = QMessageBox.question(
            self,
            "Delete Column",
            (
                f"Are you sure you want to delete "
                f"column {column_letter}?"
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if answer != QMessageBox.StandardButton.Yes:
            return

        self.delete_column_button.setEnabled(False)
        self.delete_column_button.setText("Deleting...")

        try:
            result = self.api_client.delete_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
            )

            worksheet = self.workbook[sheet_name]

            worksheet.delete_cols(
                column_number,
                1,
            )

            self._loading_table = True
            self.table.blockSignals(True)

            try:
                self.table.removeColumn(
                    current_column
                )
            finally:
                self.table.blockSignals(False)
                self._loading_table = False

            self._shift_pending_changes_after_column_delete(
                sheet_name,
                column_number,
            )

            QMessageBox.information(
                self,
                "Column Deleted",
                result.get(
                    "message",
                    "Column deleted successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Delete Column Failed",
                str(exc),
            )

        finally:
            self.delete_column_button.setEnabled(True)
            self.delete_column_button.setText("Delete Column")

    def _shift_pending_changes_after_column_delete(
        self,
        sheet_name: str,
        deleted_column: int,
    ):
        updated = {}

        for (
            change_sheet,
            cell,
        ), value in self.pending_changes.items():

            if change_sheet != sheet_name:
                updated[(change_sheet, cell)] = value
                continue

            letters = ""
            digits = ""

            for char in cell:
                if char.isalpha():
                    letters += char
                elif char.isdigit():
                    digits += char

            if not letters or not digits:
                updated[(change_sheet, cell)] = value
                continue

            old_column = self.column_number(letters)

            if old_column == deleted_column:
                continue

            if old_column > deleted_column:
                old_column -= 1

            new_cell = (
                self.column_letter(old_column)
                + digits
            )

            updated[
                (change_sheet, new_cell)
            ] = value

        self.pending_changes = updated

    @staticmethod
    def column_number(column_letters: str) -> int:
        result = 0

        for char in column_letters.upper():
            if not ("A" <= char <= "Z"):
                continue

            result = (
                result * 26
                + ord(char)
                - ord("A")
                + 1
            )

        return result

    # =========================================================
    # Create Sheet
    # =========================================================

    def create_sheet(self):
        if not self.workbook:
            QMessageBox.warning(
                self,
                "Create Sheet",
                "Workbook is not loaded.",
            )
            return

        sheet_name, ok = QInputDialog.getText(
            self,
            "Create Sheet",
            "Enter new sheet name:",
        )

        if not ok:
            return

        sheet_name = sheet_name.strip()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Create Sheet",
                "Sheet name cannot be empty.",
            )
            return

        if sheet_name in self.workbook.sheetnames:
            QMessageBox.warning(
                self,
                "Create Sheet",
                (
                    f"Sheet '{sheet_name}' "
                    "already exists."
                ),
            )
            return

        self.create_sheet_button.setEnabled(False)
        self.create_sheet_button.setText("Creating...")

        try:
            result = self.api_client.create_sheet(
                file_id=self.file_id,
                sheet_name=sheet_name,
            )

            self.workbook.create_sheet(sheet_name)

            self.rebuild_sheet_combo(sheet_name)
            self.load_sheet(sheet_name)

            QMessageBox.information(
                self,
                "Sheet Created",
                result.get(
                    "message",
                    "Sheet created successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Create Sheet Failed",
                str(exc),
            )

        finally:
            self.create_sheet_button.setEnabled(True)
            self.create_sheet_button.setText("New Sheet")

    # =========================================================
    # Rename Sheet
    # =========================================================

    def rename_sheet(self):
        if not self.workbook:
            QMessageBox.warning(
                self,
                "Rename Sheet",
                "Workbook is not loaded.",
            )
            return

        current_sheet = self.selected_sheet()

        if not current_sheet:
            QMessageBox.warning(
                self,
                "Rename Sheet",
                "Please select a sheet first.",
            )
            return

        new_sheet_name, ok = QInputDialog.getText(
            self,
            "Rename Sheet",
            "Enter new sheet name:",
            text=current_sheet,
        )

        if not ok:
            return

        new_sheet_name = new_sheet_name.strip()

        if not new_sheet_name:
            QMessageBox.warning(
                self,
                "Rename Sheet",
                "Sheet name cannot be empty.",
            )
            return

        if new_sheet_name == current_sheet:
            QMessageBox.information(
                self,
                "Rename Sheet",
                "The sheet name has not changed.",
            )
            return

        if new_sheet_name in self.workbook.sheetnames:
            QMessageBox.warning(
                self,
                "Rename Sheet",
                (
                    f"Sheet '{new_sheet_name}' "
                    "already exists."
                ),
            )
            return

        self.rename_sheet_button.setEnabled(False)
        self.rename_sheet_button.setText("Renaming...")

        try:
            result = self.api_client.rename_sheet(
                file_id=self.file_id,
                sheet_name=current_sheet,
                new_sheet_name=new_sheet_name,
            )

            worksheet = self.workbook[current_sheet]
            worksheet.title = new_sheet_name

            updated_pending = {}

            for (
                key,
                value,
            ) in self.pending_changes.items():

                old_sheet, cell = key

                if old_sheet == current_sheet:
                    updated_pending[
                        (new_sheet_name, cell)
                    ] = value
                else:
                    updated_pending[key] = value

            self.pending_changes = updated_pending

            self.rebuild_sheet_combo(new_sheet_name)
            self.load_sheet(new_sheet_name)

            QMessageBox.information(
                self,
                "Sheet Renamed",
                result.get(
                    "message",
                    "Sheet renamed successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Rename Sheet Failed",
                str(exc),
            )

        finally:
            self.rename_sheet_button.setEnabled(True)
            self.rename_sheet_button.setText("Rename Sheet")

    # =========================================================
    # Delete Sheet
    # =========================================================

    def delete_sheet(self):
        if not self.workbook:
            QMessageBox.warning(
                self,
                "Delete Sheet",
                "Workbook is not loaded.",
            )
            return

        current_sheet = self.selected_sheet()

        if not current_sheet:
            QMessageBox.warning(
                self,
                "Delete Sheet",
                "Please select a sheet first.",
            )
            return

        if len(self.workbook.sheetnames) == 1:
            QMessageBox.warning(
                self,
                "Delete Sheet",
                (
                    "Cannot delete the only "
                    "sheet in the workbook."
                ),
            )
            return

        answer = QMessageBox.question(
            self,
            "Delete Sheet",
            (
                f"Are you sure you want to "
                f"delete sheet '{current_sheet}'?"
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if answer != QMessageBox.StandardButton.Yes:
            return

        self.delete_sheet_button.setEnabled(False)
        self.delete_sheet_button.setText("Deleting...")

        try:
            result = self.api_client.delete_sheet(
                file_id=self.file_id,
                sheet_name=current_sheet,
            )

            worksheet = self.workbook[current_sheet]
            self.workbook.remove(worksheet)

            self.clear_pending_for_sheet(current_sheet)

            remaining = self.workbook.sheetnames

            self.rebuild_sheet_combo(
                remaining[0] if remaining else None
            )

            if remaining:
                self.load_sheet(
                    self.sheet_combo.currentText()
                )
            else:
                self.table.clear()
                self.table.setRowCount(0)
                self.table.setColumnCount(0)

            QMessageBox.information(
                self,
                "Sheet Deleted",
                result.get(
                    "message",
                    "Sheet deleted successfully.",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Delete Sheet Failed",
                str(exc),
            )

        finally:
            self.delete_sheet_button.setEnabled(True)
            self.delete_sheet_button.setText("Delete Sheet")

    # =========================================================
    # Cleanup
    # =========================================================

    def closeEvent(self, event):
        if self.temp_file_path:
            try:
                os.remove(self.temp_file_path)
            except OSError:
                pass

        event.accept()
