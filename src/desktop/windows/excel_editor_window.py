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

        # Store unsaved cell changes
        self.pending_changes = {}

        self.setWindowTitle(
            f"Excel Editor - {self.filename}"
        )

        self.setMinimumSize(1200, 700)

        self.setup_ui()
        self.load_excel_file()

        # =====================================================
        # Button Connections
        # =====================================================

        self.add_row_button.clicked.connect(
            self.add_row
        )

        self.delete_row_button.clicked.connect(
            self.delete_row
        )

        self.add_column_button.clicked.connect(
            self.add_column
        )

        self.rename_column_button.clicked.connect(
            self.rename_column
        )

        self.delete_column_button.clicked.connect(
            self.delete_column
        )

        self.save_button.clicked.connect(
            self.save_changes
        )

    # =========================================================
    # UI Setup
    # =========================================================

    def setup_ui(self):
        main_layout = QVBoxLayout()

        main_layout.setContentsMargins(
            20,
            20,
            20,
            20,
        )

        main_layout.setSpacing(15)

        # =====================================================
        # Header
        # =====================================================

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

        # =====================================================
        # Sheet Selector
        # =====================================================

        sheet_label = QLabel("Sheet:")

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(180)

        self.sheet_combo.currentTextChanged.connect(
            self.load_sheet
        )

        header_layout.addWidget(sheet_label)
        header_layout.addWidget(
            self.sheet_combo
        )

        # =====================================================
        # Refresh
        # =====================================================

        self.refresh_button = QPushButton(
            "Refresh"
        )

        self.refresh_button.setMinimumHeight(38)

        self.refresh_button.clicked.connect(
            self.load_excel_file
        )

        header_layout.addWidget(
            self.refresh_button
        )

        main_layout.addLayout(
            header_layout
        )

        # =====================================================
        # Excel Table
        # =====================================================

        self.table = QTableWidget()

        self.table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked
            | QTableWidget.EditTrigger.EditKeyPressed
        )

        self.table.setAlternatingRowColors(
            True
        )

        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectItems
        )

        self.table.cellChanged.connect(
            self.handle_cell_changed
        )

        self.table.horizontalHeader().setStretchLastSection(
            True
        )

        main_layout.addWidget(
            self.table
        )

        # =====================================================
        # Bottom Buttons
        # =====================================================

        bottom_layout = QHBoxLayout()

        self.add_row_button = QPushButton(
            "Add Row"
        )

        self.delete_row_button = QPushButton(
            "Delete Row"
        )

        self.add_column_button = QPushButton(
            "Add Column"
        )

        self.rename_column_button = QPushButton(
            "Rename Column"
        )

        self.delete_column_button = QPushButton(
            "Delete Column"
        )

        self.save_button = QPushButton(
            "Save"
        )

        for button in (
            self.add_row_button,
            self.delete_row_button,
            self.add_column_button,
            self.rename_column_button,
            self.delete_column_button,
            self.save_button,
        ):
            button.setMinimumHeight(40)

        bottom_layout.addWidget(
            self.add_row_button
        )

        bottom_layout.addWidget(
            self.delete_row_button
        )

        bottom_layout.addWidget(
            self.add_column_button
        )

        bottom_layout.addWidget(
            self.rename_column_button
        )

        bottom_layout.addWidget(
            self.delete_column_button
        )

        bottom_layout.addStretch()

        bottom_layout.addWidget(
            self.save_button
        )

        main_layout.addLayout(
            bottom_layout
        )

        self.setLayout(
            main_layout
        )

    # =========================================================
    # Download + Load Excel
    # =========================================================

    def load_excel_file(self):
        try:
            self.refresh_button.setEnabled(
                False
            )

            self.refresh_button.setText(
                "Loading..."
            )

            # Clear pending changes
            self.pending_changes.clear()

            # Remove previous temporary file
            if self.temp_file_path:
                try:
                    os.remove(
                        self.temp_file_path
                    )
                except OSError:
                    pass

            # Create temporary file
            fd, self.temp_file_path = tempfile.mkstemp(
                suffix=".xlsx"
            )

            os.close(fd)

            # Download Excel file
            self.api_client.download_file(
                file_id=self.file_id,
                save_path=self.temp_file_path,
            )

            # Load workbook
            self.workbook = load_workbook(
                self.temp_file_path
            )

            # Update sheet selector
            self.sheet_combo.blockSignals(
                True
            )

            try:
                self.sheet_combo.clear()

                self.sheet_combo.addItems(
                    self.workbook.sheetnames
                )

            finally:
                self.sheet_combo.blockSignals(
                    False
                )

            # Load first sheet
            if self.workbook.sheetnames:
                self.load_sheet(
                    self.workbook.sheetnames[0]
                )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Excel File",
                str(exc),
            )

        finally:
            self.refresh_button.setEnabled(
                True
            )

            self.refresh_button.setText(
                "Refresh"
            )

    # =========================================================
    # Load Selected Sheet
    # =========================================================

    def load_sheet(
        self,
        sheet_name: str,
    ):
        if not self.workbook:
            return

        if not sheet_name:
            return

        self.table.blockSignals(
            True
        )

        try:
            worksheet = self.workbook[
                sheet_name
            ]

            rows = list(
                worksheet.iter_rows(
                    values_only=True
                )
            )

            # Empty sheet
            if not rows:
                self.table.clear()
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                return

            column_count = max(
                len(row)
                for row in rows
            )

            row_count = len(rows)

            self.table.clear()

            self.table.setRowCount(
                row_count
            )

            self.table.setColumnCount(
                column_count
            )

            # Column Headers
            headers = []

            for column_index in range(
                column_count
            ):
                headers.append(
                    self.column_letter(
                        column_index + 1
                    )
                )

            self.table.setHorizontalHeaderLabels(
                headers
            )

            # Fill Table
            for row_index, row in enumerate(
                rows
            ):
                for column_index in range(
                    column_count
                ):
                    value = ""

                    if column_index < len(row):
                        cell_value = row[
                            column_index
                        ]

                        if cell_value is not None:
                            value = str(
                                cell_value
                            )

                    item = QTableWidgetItem(
                        value
                    )

                    self.table.setItem(
                        row_index,
                        column_index,
                        item,
                    )

            self.table.resizeColumnsToContents()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Sheet",
                str(exc),
            )

        finally:
            self.table.blockSignals(
                False
            )

    # =========================================================
    # Cell Editing
    # =========================================================

    def handle_cell_changed(
        self,
        row: int,
        column: int,
    ):
        if not self.workbook:
            return

        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            return

        item = self.table.item(
            row,
            column,
        )

        if not item:
            value = None
        else:
            value = item.text()

        cell = (
            self.column_letter(column + 1)
            + str(row + 1)
        )

        # Store change locally
        self.pending_changes[
            (sheet_name, cell)
        ] = value

        # Update local workbook
        try:
            worksheet = self.workbook[
                sheet_name
            ]

            worksheet[cell] = value

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Update Cell",
                str(exc),
            )

    # =========================================================
    # Save Changes
    # =========================================================

    def save_changes(self):
        if not self.pending_changes:
            QMessageBox.information(
                self,
                "Save",
                "There are no changes to save.",
            )
            return

        self.save_button.setEnabled(
            False
        )

        self.save_button.setText(
            "Saving..."
        )

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
                    (
                        sheet_name,
                        cell,
                    )
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
            self.save_button.setEnabled(
                True
            )

            self.save_button.setText(
                "Save"
            )

    # =========================================================
    # Add Row
    # =========================================================

    def add_row(self):
        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Add Row",
                "Please select a sheet first.",
            )
            return

        column_count = (
            self.table.columnCount()
        )

        row_data = [
            None
        ] * column_count

        try:
            self.add_row_button.setEnabled(
                False
            )

            self.add_row_button.setText(
                "Adding..."
            )

            result = self.api_client.add_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_data=row_data,
            )

            new_row = (
                self.table.rowCount()
            )

            self.table.blockSignals(
                True
            )

            try:
                self.table.insertRow(
                    new_row
                )

                for column in range(
                    column_count
                ):
                    self.table.setItem(
                        new_row,
                        column,
                        QTableWidgetItem(""),
                    )

            finally:
                self.table.blockSignals(
                    False
                )

            # Update local workbook
            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.append(
                    row_data
                )

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
            self.add_row_button.setEnabled(
                True
            )

            self.add_row_button.setText(
                "Add Row"
            )

    # =========================================================
    # Delete Row
    # =========================================================

    def delete_row(self):
        current_row = (
            self.table.currentRow()
        )

        if current_row < 0:
            QMessageBox.warning(
                self,
                "Delete Row",
                "Please select a row to delete.",
            )
            return

        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Delete Row",
                "Please select a sheet first.",
            )
            return

        row_number = current_row + 1

        answer = QMessageBox.question(
            self,
            "Delete Row",
            f"Are you sure you want to delete row {row_number}?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if answer != QMessageBox.StandardButton.Yes:
            return

        try:
            self.delete_row_button.setEnabled(
                False
            )

            self.delete_row_button.setText(
                "Deleting..."
            )

            result = self.api_client.delete_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_number=row_number,
            )

            self.table.blockSignals(
                True
            )

            try:
                self.table.removeRow(
                    current_row
                )

                if self.workbook:
                    worksheet = self.workbook[
                        sheet_name
                    ]

                    worksheet.delete_rows(
                        row_number,
                        1,
                    )

            finally:
                self.table.blockSignals(
                    False
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
            self.delete_row_button.setEnabled(
                True
            )

            self.delete_row_button.setText(
                "Delete Row"
            )

    # =========================================================
    # Add Column
    # =========================================================

    def add_column(self):
        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Add Column",
                "Please select a sheet first.",
            )
            return

        current_column_count = (
            self.table.columnCount()
        )

        column_number = (
            current_column_count + 1
        )

        try:
            self.add_column_button.setEnabled(
                False
            )

            self.add_column_button.setText(
                "Adding..."
            )

            result = self.api_client.add_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
            )

            self.table.blockSignals(
                True
            )

            try:
                self.table.insertColumn(
                    current_column_count
                )

                self.table.setHorizontalHeaderItem(
                    current_column_count,
                    QTableWidgetItem(
                        self.column_letter(
                            column_number
                        )
                    ),
                )

                for row in range(
                    self.table.rowCount()
                ):
                    self.table.setItem(
                        row,
                        current_column_count,
                        QTableWidgetItem(""),
                    )

            finally:
                self.table.blockSignals(
                    False
                )

            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.insert_cols(
                    column_number,
                    1,
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
            self.add_column_button.setEnabled(
                True
            )

            self.add_column_button.setText(
                "Add Column"
            )

    # =========================================================
    # Rename Column
    # =========================================================

    def rename_column(self):
        current_column = (
            self.table.currentColumn()
        )

        if current_column < 0:
            QMessageBox.warning(
                self,
                "Rename Column",
                "Please select a column first.",
            )
            return

        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Rename Column",
                "Please select a sheet first.",
            )
            return

        column_number = (
            current_column + 1
        )

        column_letter = self.column_letter(
            column_number
        )

        # Get current header value
        header_item = (
            self.table.horizontalHeaderItem(
                current_column
            )
        )

        current_name = ""

        if header_item:
            current_name = header_item.text()

        # Ask for new column name
        new_column_name, ok = QInputDialog.getText(
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

        new_column_name = (
            new_column_name.strip()
        )

        if not new_column_name:
            QMessageBox.warning(
                self,
                "Rename Column",
                "Column name cannot be empty.",
            )
            return

        if new_column_name == current_name:
            QMessageBox.information(
                self,
                "Rename Column",
                "The column name has not changed.",
            )
            return

        try:
            self.rename_column_button.setEnabled(
                False
            )

            self.rename_column_button.setText(
                "Renaming..."
            )

            # -------------------------------------------------
            # Backend
            # -------------------------------------------------

            result = self.api_client.update_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
                column_name=new_column_name,
            )

            # -------------------------------------------------
            # Update UI
            # -------------------------------------------------

            self.table.blockSignals(
                True
            )

            try:
                self.table.setHorizontalHeaderItem(
                    current_column,
                    QTableWidgetItem(
                        new_column_name
                    ),
                )

            finally:
                self.table.blockSignals(
                    False
                )

            # -------------------------------------------------
            # Update local workbook
            # -------------------------------------------------

            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.cell(
                    row=1,
                    column=column_number,
                    value=new_column_name,
                )

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
            self.rename_column_button.setEnabled(
                True
            )

            self.rename_column_button.setText(
                "Rename Column"
            )

    # =========================================================
    # Delete Column
    # =========================================================

    def delete_column(self):
        current_column = (
            self.table.currentColumn()
        )

        if current_column < 0:
            QMessageBox.warning(
                self,
                "Delete Column",
                "Please select a column first.",
            )
            return

        sheet_name = (
            self.sheet_combo.currentText()
        )

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Delete Column",
                "Please select a sheet first.",
            )
            return

        column_number = (
            current_column + 1
        )

        column_letter = self.column_letter(
            column_number
        )

        # Confirmation
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

        try:
            self.delete_column_button.setEnabled(
                False
            )

            self.delete_column_button.setText(
                "Deleting..."
            )

            # Backend
            result = self.api_client.delete_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
            )

            # Update UI
            self.table.blockSignals(
                True
            )

            try:
                self.table.removeColumn(
                    current_column
                )

                # Update local workbook
                if self.workbook:
                    worksheet = self.workbook[
                        sheet_name
                    ]

                    worksheet.delete_cols(
                        column_number,
                        1,
                    )

            finally:
                self.table.blockSignals(
                    False
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
            self.delete_column_button.setEnabled(
                True
            )

            self.delete_column_button.setText(
                "Delete Column"
            )

    # =========================================================
    # Excel Column Letter
    # =========================================================

    def column_letter(
        self,
        column_number: int,
    ) -> str:
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

    # =========================================================
    # Cleanup
    # =========================================================

    def closeEvent(
        self,
        event,
    ):
        if self.temp_file_path:
            try:
                os.remove(
                    self.temp_file_path
                )
            except OSError:
                pass

        event.accept()