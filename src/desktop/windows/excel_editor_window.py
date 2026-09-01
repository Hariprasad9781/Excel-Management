import os
import re
import tempfile

from openpyxl import load_workbook

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


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

        # ---------------------------------------------------------
        # Pending cell changes
        # ---------------------------------------------------------

        self.pending_changes = {}

        # ---------------------------------------------------------
        # Internal state
        # ---------------------------------------------------------

        self._loading_table = False
        self._searching = False

        self.setWindowTitle(
            f"Excel Editor - {self.filename}"
        )

        self.setMinimumSize(
            1200,
            800,
        )

        self.setup_ui()

        # ---------------------------------------------------------
        # Load workbook
        # ---------------------------------------------------------

        self.load_excel_file()

        # ---------------------------------------------------------
        # Button connections
        # ---------------------------------------------------------

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

        # Sheet management

        self.create_sheet_button.clicked.connect(
            self.create_sheet
        )

        self.rename_sheet_button.clicked.connect(
            self.rename_sheet
        )

        self.delete_sheet_button.clicked.connect(
            self.delete_sheet
        )

        # Search

        self.search_button.clicked.connect(
            self.search_excel
        )

        self.clear_search_button.clicked.connect(
            self.clear_search
        )

        self.search_input.returnPressed.connect(
            self.search_excel
        )

        self.search_results_table.cellClicked.connect(
            self.select_search_result
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

        main_layout.setSpacing(12)

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
        # Sheet selector
        # =====================================================

        sheet_label = QLabel("Sheet:")

        self.sheet_combo = QComboBox()

        self.sheet_combo.setMinimumWidth(
            180
        )

        self.sheet_combo.currentTextChanged.connect(
            self.load_sheet
        )

        header_layout.addWidget(
            sheet_label
        )

        header_layout.addWidget(
            self.sheet_combo
        )

        # =====================================================
        # New Sheet
        # =====================================================

        self.create_sheet_button = QPushButton(
            "New Sheet"
        )

        # =====================================================
        # Rename Sheet
        # =====================================================

        self.rename_sheet_button = QPushButton(
            "Rename Sheet"
        )

        # =====================================================
        # Delete Sheet
        # =====================================================

        self.delete_sheet_button = QPushButton(
            "Delete Sheet"
        )

        # =====================================================
        # Refresh
        # =====================================================

        self.refresh_button = QPushButton(
            "Refresh"
        )

        for button in (
            self.create_sheet_button,
            self.rename_sheet_button,
            self.delete_sheet_button,
            self.refresh_button,
        ):
            button.setMinimumHeight(38)

        header_layout.addWidget(
            self.create_sheet_button
        )

        header_layout.addWidget(
            self.rename_sheet_button
        )

        header_layout.addWidget(
            self.delete_sheet_button
        )

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
        # Search Area
        # =====================================================

        search_layout = QHBoxLayout()

        search_label = QLabel(
            "Search:"
        )

        self.search_input = QLineEdit()

        self.search_input.setPlaceholderText(
            "Search in current sheet..."
        )

        self.search_button = QPushButton(
            "Search"
        )

        self.clear_search_button = QPushButton(
            "Clear"
        )

        self.search_status_label = QLabel()

        self.search_status_label.setMinimumWidth(
            100
        )

        search_layout.addWidget(
            search_label
        )

        search_layout.addWidget(
            self.search_input
        )

        search_layout.addWidget(
            self.search_button
        )

        search_layout.addWidget(
            self.clear_search_button
        )

        search_layout.addWidget(
            self.search_status_label
        )

        main_layout.addLayout(
            search_layout
        )

        # =====================================================
        # Search Results Table
        # =====================================================

        self.search_results_table = QTableWidget()

        self.search_results_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )

        self.search_results_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        self.search_results_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )

        self.search_results_table.setAlternatingRowColors(
            True
        )

        self.search_results_table.setMaximumHeight(
            230
        )

        self.search_results_table.horizontalHeader().setStretchLastSection(
            True
        )

        main_layout.addWidget(
            self.search_results_table
        )

        # =====================================================
        # Main Excel Table
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

        self.table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
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
            button.setMinimumHeight(
                40
            )

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
    # Helper - Selected Sheet
    # =========================================================

    def selected_sheet(self) -> str:
        return self.sheet_combo.currentText().strip()

    # =========================================================
    # Load Excel File
    # =========================================================

    def load_excel_file(self):
        try:
            self.refresh_button.setEnabled(
                False
            )

            self.refresh_button.setText(
                "Loading..."
            )

            self.pending_changes.clear()

            # -------------------------------------------------
            # Clear search
            # -------------------------------------------------

            self.clear_search_results_only()

            # -------------------------------------------------
            # Remove previous temp file
            # -------------------------------------------------

            if self.temp_file_path:
                try:
                    os.remove(
                        self.temp_file_path
                    )
                except OSError:
                    pass

            # -------------------------------------------------
            # Create temporary file
            # -------------------------------------------------

            fd, self.temp_file_path = (
                tempfile.mkstemp(
                    suffix=".xlsx"
                )
            )

            os.close(fd)

            # -------------------------------------------------
            # Download latest Excel file
            # -------------------------------------------------

            self.api_client.download_file(
                file_id=self.file_id,
                save_path=self.temp_file_path,
            )

            # -------------------------------------------------
            # Load workbook
            # -------------------------------------------------

            self.workbook = load_workbook(
                self.temp_file_path
            )

            # -------------------------------------------------
            # Update sheet selector
            # -------------------------------------------------

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

            # -------------------------------------------------
            # Load first sheet
            # -------------------------------------------------

            if self.workbook.sheetnames:
                self.sheet_combo.setCurrentIndex(
                    0
                )

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
    # Load Sheet
    # =========================================================

    def load_sheet(
        self,
        sheet_name: str,
    ):
        if not self.workbook:
            return

        if not sheet_name:
            return

        self.clear_search_results_only()

        self._loading_table = True

        self.table.blockSignals(
            True
        )

        try:
            worksheet = self.workbook[
                sheet_name
            ]

            max_row = worksheet.max_row
            max_column = worksheet.max_column

            # -------------------------------------------------
            # Completely empty worksheet
            # -------------------------------------------------

            if (
                max_row == 1
                and max_column == 1
                and worksheet["A1"].value is None
            ):
                self.table.clear()

                self.table.setRowCount(
                    0
                )

                self.table.setColumnCount(
                    0
                )

                return

            # -------------------------------------------------
            # First row = headers
            # -------------------------------------------------

            header_row = 1
            data_start_row = 2

            data_row_count = max(
                0,
                max_row - 1,
            )

            self.table.clear()

            self.table.setRowCount(
                data_row_count
            )

            self.table.setColumnCount(
                max_column
            )

            # -------------------------------------------------
            # Headers
            # -------------------------------------------------

            headers = []

            for column_number in range(
                1,
                max_column + 1,
            ):
                header_value = worksheet.cell(
                    row=header_row,
                    column=column_number,
                ).value

                if header_value is None:
                    header_value = (
                        self.column_letter(
                            column_number
                        )
                    )
                else:
                    header_value = str(
                        header_value
                    )

                headers.append(
                    header_value
                )

            self.table.setHorizontalHeaderLabels(
                headers
            )

            # -------------------------------------------------
            # Data
            # -------------------------------------------------

            for table_row in range(
                data_row_count
            ):
                excel_row = (
                    data_start_row
                    + table_row
                )

                for column_number in range(
                    1,
                    max_column + 1,
                ):
                    value = worksheet.cell(
                        row=excel_row,
                        column=column_number,
                    ).value

                    if value is None:
                        value = ""

                    item = QTableWidgetItem(
                        str(value)
                    )

                    self.table.setItem(
                        table_row,
                        column_number - 1,
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

            self._loading_table = False

    # =========================================================
    # Search Excel
    # =========================================================

    def search_excel(self):
        if self._searching:
            return

        search_term = (
            self.search_input.text().strip()
        )

        if not search_term:
            QMessageBox.warning(
                self,
                "Search",
                "Please enter something to search.",
            )

            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Search",
                "Please select a sheet first.",
            )

            return

        if not self.workbook:
            QMessageBox.warning(
                self,
                "Search",
                "Workbook is not loaded.",
            )

            return

        self._searching = True

        self.search_button.setEnabled(
            False
        )

        self.search_button.setText(
            "Searching..."
        )

        self.search_status_label.setText(
            ""
        )

        self.clear_search_results_only()

        try:
            # -------------------------------------------------
            # Call existing API client search method
            # -------------------------------------------------

            response = (
                self.api_client.search_excel(
                    file_id=self.file_id,
                    sheet_name=sheet_name,
                    search_term=search_term,
                )
            )

            results = response.get(
                "results",
                [],
            )

            # -------------------------------------------------
            # No results
            # -------------------------------------------------

            if not results:
                self.search_status_label.setText(
                    "No results"
                )

                return

            # -------------------------------------------------
            # Build complete rows
            # -------------------------------------------------

            self.populate_search_results(
                results=results,
                search_term=search_term,
                sheet_name=sheet_name,
            )

            self.search_status_label.setText(
                f"{len(results)} match(es)"
            )

        except Exception as exc:
            self.search_status_label.setText(
                "Search failed."
            )

            QMessageBox.critical(
                self,
                "Search Failed",
                str(exc),
            )

        finally:
            self._searching = False

            self.search_button.setEnabled(
                True
            )

            self.search_button.setText(
                "Search"
            )

    # =========================================================
    # Populate Search Results
    # =========================================================

    def populate_search_results(
        self,
        results: list[dict],
        search_term: str,
        sheet_name: str,
    ):
        if not self.workbook:
            return

        worksheet = self.workbook[
            sheet_name
        ]

        max_column = worksheet.max_column

        # -----------------------------------------------------
        # Headers
        # -----------------------------------------------------

        headers = []

        for column_number in range(
            1,
            max_column + 1,
        ):
            header_value = worksheet.cell(
                row=1,
                column=column_number,
            ).value

            if header_value is None:
                header_value = (
                    self.column_letter(
                        column_number
                    )
                )
            else:
                header_value = str(
                    header_value
                )

            headers.append(
                header_value
            )

        # -----------------------------------------------------
        # Group matches by row
        # -----------------------------------------------------

        rows = {}

        for result in results:
            row_number = result.get(
                "row_number"
            )

            column_number = result.get(
                "column_number"
            )

            cell = result.get(
                "cell",
                "",
            )

            if row_number is None:
                continue

            if row_number not in rows:
                rows[row_number] = {
                    "matched_columns": set(),
                    "matched_cells": [],
                }

            if column_number is not None:
                rows[row_number][
                    "matched_columns"
                ].add(
                    column_number
                )

            if cell:
                rows[row_number][
                    "matched_cells"
                ].append(cell)

        sorted_rows = sorted(
            rows.items(),
            key=lambda item: item[0],
        )

        # -----------------------------------------------------
        # Search result columns
        #
        # Row | Excel Columns... | Matched Cell
        # -----------------------------------------------------

        result_headers = [
            "Row"
        ]

        result_headers.extend(
            headers
        )

        result_headers.append(
            "Matched Cell"
        )

        self.search_results_table.setColumnCount(
            len(result_headers)
        )

        self.search_results_table.setHorizontalHeaderLabels(
            result_headers
        )

        self.search_results_table.setRowCount(
            len(sorted_rows)
        )

        # -----------------------------------------------------
        # Fill result rows
        # -----------------------------------------------------

        for result_index, (
            excel_row,
            match_info,
        ) in enumerate(
            sorted_rows
        ):
            # -------------------------------------------------
            # Row number
            # -------------------------------------------------

            row_item = QTableWidgetItem(
                str(excel_row)
            )

            row_item.setData(
                Qt.ItemDataRole.UserRole,
                excel_row,
            )

            self.search_results_table.setItem(
                result_index,
                0,
                row_item,
            )

            # -------------------------------------------------
            # Complete Excel row
            # -------------------------------------------------

            for column_number in range(
                1,
                max_column + 1,
            ):
                value = worksheet.cell(
                    row=excel_row,
                    column=column_number,
                ).value

                if value is None:
                    value = ""

                item = QTableWidgetItem(
                    str(value)
                )

                # -------------------------------------------------
                # Highlight matching column
                # -------------------------------------------------

                if (
                    column_number
                    in match_info[
                        "matched_columns"
                    ]
                ):
                    item.setBackground(
                        QColor(
                            255,
                            235,
                            150,
                        )
                    )

                    item.setToolTip(
                        "Matching value"
                    )

                self.search_results_table.setItem(
                    result_index,
                    column_number,
                    item,
                )

            # -------------------------------------------------
            # Matched cells
            # -------------------------------------------------

            matched_cells = ", ".join(
                match_info[
                    "matched_cells"
                ]
            )

            matched_item = QTableWidgetItem(
                matched_cells
            )

            self.search_results_table.setItem(
                result_index,
                max_column + 1,
                matched_item,
            )

        # -----------------------------------------------------
        # Resize
        # -----------------------------------------------------

        self.search_results_table.resizeColumnsToContents()

        # Keep the table usable if there are many columns.
        self.search_results_table.horizontalHeader().setStretchLastSection(
            False
        )

    # =========================================================
    # Select Search Result
    # =========================================================

    def select_search_result(
        self,
        result_row: int,
        result_column: int,
    ):
        if not self.workbook:
            return

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        # -----------------------------------------------------
        # Excel row number stored in first column
        # -----------------------------------------------------

        row_item = (
            self.search_results_table.item(
                result_row,
                0,
            )
        )

        if not row_item:
            return

        excel_row = row_item.data(
            Qt.ItemDataRole.UserRole
        )

        if excel_row is None:
            try:
                excel_row = int(
                    row_item.text()
                )
            except ValueError:
                return

        # -----------------------------------------------------
        # Matched cell column
        # -----------------------------------------------------

        matched_column_index = (
            self.search_results_table.columnCount()
            - 1
        )

        matched_item = (
            self.search_results_table.item(
                result_row,
                matched_column_index,
            )
        )

        target_cell = None

        if matched_item:
            matched_text = (
                matched_item.text()
            )

            if matched_text:
                target_cell = (
                    matched_text.split(
                        ","
                    )[0].strip()
                )

        # -----------------------------------------------------
        # Select target in main Excel table
        # -----------------------------------------------------

        table_row = (
            excel_row - 2
        )

        if table_row < 0:
            return

        if (
            table_row
            >= self.table.rowCount()
        ):
            return

        table_column = 0

        if target_cell:
            match = re.match(
                r"^([A-Za-z]+)(\d+)$",
                target_cell,
            )

            if match:
                letters = match.group(
                    1
                )

                table_column = (
                    self.column_number(
                        letters
                    )
                    - 1
                )

        if (
            table_column < 0
            or table_column
            >= self.table.columnCount()
        ):
            table_column = 0

        self.table.setCurrentCell(
            table_row,
            table_column,
        )

        self.table.scrollToItem(
            self.table.item(
                table_row,
                table_column,
            )
        )

    # =========================================================
    # Clear Search Results Only
    # =========================================================

    def clear_search_results_only(self):
        self.search_results_table.clear()

        self.search_results_table.setRowCount(
            0
        )

        self.search_results_table.setColumnCount(
            0
        )

        self.search_status_label.setText(
            ""
        )

    # =========================================================
    # Clear Search
    # =========================================================

    def clear_search(self):
        self.search_input.clear()

        self.clear_search_results_only()

        self.search_input.setFocus()

    # =========================================================
    # Cell Changed
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

        item = self.table.item(
            row,
            column,
        )

        if item is None:
            value = None
        else:
            value = item.text()

        # -----------------------------------------------------
        # Table row 0 = Excel row 2
        # -----------------------------------------------------

        excel_row = row + 2

        cell = (
            self.column_letter(
                column + 1
            )
            + str(excel_row)
        )

        self.pending_changes[
            (
                sheet_name,
                cell,
            )
        ] = value

        # -----------------------------------------------------
        # Update local workbook
        # -----------------------------------------------------

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

        successful_changes = []

        try:
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

            for key in successful_changes:
                self.pending_changes.pop(
                    key,
                    None,
                )

            QMessageBox.information(
                self,
                "Save Successful",
                "Changes saved successfully.",
            )

            # Refresh search results if active.
            if self.search_input.text().strip():
                self.search_excel()

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
        sheet_name = self.selected_sheet()

        if not sheet_name:
            QMessageBox.warning(
                self,
                "Add Row",
                "Please select a sheet first.",
            )

            return

        if not self.workbook:
            return

        column_count = max(
            self.table.columnCount(),
            1,
        )

        # IMPORTANT:
        # Empty strings make this a real row in
        # openpyxl and prevent max_row collapsing.
        row_data = [
            ""
        ] * column_count

        self.add_row_button.setEnabled(
            False
        )

        self.add_row_button.setText(
            "Adding..."
        )

        try:
            # -------------------------------------------------
            # Add row in backend
            # -------------------------------------------------

            result = self.api_client.add_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_data=row_data,
            )

            backend_row_number = result.get(
                "row_number"
            )

            if backend_row_number is None:
                backend_row_number = (
                    self.table.rowCount()
                    + 2
                )

            # -------------------------------------------------
            # Update local workbook
            # -------------------------------------------------

            worksheet = self.workbook[
                sheet_name
            ]

            worksheet.append(
                row_data
            )

            # -------------------------------------------------
            # Make sure local workbook row
            # matches backend row
            # -------------------------------------------------

            local_row_number = (
                worksheet.max_row
            )

            if (
                local_row_number
                != backend_row_number
            ):
                for column_number, value in enumerate(
                    row_data,
                    start=1,
                ):
                    worksheet.cell(
                        row=backend_row_number,
                        column=column_number,
                        value=value,
                    )

            # -------------------------------------------------
            # Update UI
            # -------------------------------------------------

            self._loading_table = True

            self.table.blockSignals(
                True
            )

            try:
                target_row = (
                    backend_row_number
                    - 2
                )

                while (
                    self.table.rowCount()
                    <= target_row
                ):
                    self.table.insertRow(
                        self.table.rowCount()
                    )

                for column in range(
                    column_count
                ):
                    self.table.setItem(
                        target_row,
                        column,
                        QTableWidgetItem(
                            ""
                        ),
                    )

            finally:
                self.table.blockSignals(
                    False
                )

                self._loading_table = False

            self.table.resizeColumnsToContents()

            self.table.setCurrentCell(
                target_row,
                0,
            )

            self.table.scrollToBottom()

            # Refresh search result if needed.
            if self.search_input.text().strip():
                self.search_excel()

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

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        # Table row 0 = Excel row 2
        row_number = (
            current_row + 2
        )

        answer = QMessageBox.question(
            self,
            "Delete Row",
            (
                f"Are you sure you want to "
                f"delete Excel row {row_number}?"
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if (
            answer
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.delete_row_button.setEnabled(
            False
        )

        self.delete_row_button.setText(
            "Deleting..."
        )

        try:
            result = self.api_client.delete_row(
                file_id=self.file_id,
                sheet_name=sheet_name,
                row_number=row_number,
            )

            # -------------------------------------------------
            # Update local workbook
            # -------------------------------------------------

            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.delete_rows(
                    row_number,
                    1,
                )

            # -------------------------------------------------
            # Remove pending changes for
            # deleted row and shift rows below
            # -------------------------------------------------

            self._shift_pending_changes_after_row_delete(
                sheet_name,
                row_number,
            )

            # -------------------------------------------------
            # Update UI
            # -------------------------------------------------

            self._loading_table = True

            self.table.blockSignals(
                True
            )

            try:
                self.table.removeRow(
                    current_row
                )

            finally:
                self.table.blockSignals(
                    False
                )

                self._loading_table = False

            # -------------------------------------------------
            # Refresh search results
            # -------------------------------------------------

            if self.search_input.text().strip():
                self.search_excel()

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
    # Shift Pending Changes After Row Delete
    # =========================================================

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
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            match = re.match(
                r"^([A-Za-z]+)(\d+)$",
                cell,
            )

            if not match:
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            letters = match.group(
                1
            )

            row_number = int(
                match.group(
                    2
                )
            )

            if row_number == deleted_row:
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

        if not self.workbook:
            return

        # Add column at the end.
        column_number = (
            self.table.columnCount()
            + 1
        )

        self.add_column_button.setEnabled(
            False
        )

        self.add_column_button.setText(
            "Adding..."
        )

        try:
            # -------------------------------------------------
            # Backend
            # -------------------------------------------------

            result = self.api_client.add_column(
                file_id=self.file_id,
                sheet_name=sheet_name,
                column_number=column_number,
            )

            # -------------------------------------------------
            # Local workbook
            # -------------------------------------------------

            worksheet = self.workbook[
                sheet_name
            ]

            worksheet.insert_cols(
                column_number,
                1,
            )

            # -------------------------------------------------
            # UI
            # -------------------------------------------------

            self._loading_table = True

            self.table.blockSignals(
                True
            )

            try:
                self.table.insertColumn(
                    column_number - 1
                )

                self.table.setHorizontalHeaderItem(
                    column_number - 1,
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
                        column_number - 1,
                        QTableWidgetItem(
                            ""
                        ),
                    )

            finally:
                self.table.blockSignals(
                    False
                )

                self._loading_table = False

            # -------------------------------------------------
            # Shift pending cell changes
            # -------------------------------------------------

            self._shift_pending_changes_after_column_insert(
                sheet_name,
                column_number,
            )

            self.table.resizeColumnsToContents()

            # Refresh search results.
            if self.search_input.text().strip():
                self.search_excel()

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
    # Shift Pending Changes After Column Insert
    # =========================================================

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
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            match = re.match(
                r"^([A-Za-z]+)(\d+)$",
                cell,
            )

            if not match:
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            letters = match.group(
                1
            )

            row_number = match.group(
                2
            )

            old_column = self.column_number(
                letters
            )

            if old_column >= inserted_column:
                old_column += 1

            new_cell = (
                self.column_letter(
                    old_column
                )
                + row_number
            )

            updated[
                (
                    change_sheet,
                    new_cell,
                )
            ] = value

        self.pending_changes = updated

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

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        column_number = (
            current_column + 1
        )

        column_letter = (
            self.column_letter(
                column_number
            )
        )

        header_item = (
            self.table.horizontalHeaderItem(
                current_column
            )
        )

        current_name = (
            header_item.text()
            if header_item
            else column_letter
        )

        new_name, ok = (
            QInputDialog.getText(
                self,
                "Rename Column",
                (
                    f"Enter new name for "
                    f"column {column_letter}:"
                ),
                text=current_name,
            )
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
            return

        self.rename_column_button.setEnabled(
            False
        )

        self.rename_column_button.setText(
            "Renaming..."
        )

        try:
            result = (
                self.api_client.update_column(
                    file_id=self.file_id,
                    sheet_name=sheet_name,
                    column_number=column_number,
                    column_name=new_name,
                )
            )

            # -------------------------------------------------
            # UI
            # -------------------------------------------------

            self._loading_table = True

            self.table.blockSignals(
                True
            )

            try:
                self.table.setHorizontalHeaderItem(
                    current_column,
                    QTableWidgetItem(
                        new_name
                    ),
                )

            finally:
                self.table.blockSignals(
                    False
                )

                self._loading_table = False

            # -------------------------------------------------
            # Local workbook
            # -------------------------------------------------

            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.cell(
                    row=1,
                    column=column_number,
                    value=new_name,
                )

            # Search result must reflect new header.
            if self.search_input.text().strip():
                self.search_excel()

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

        sheet_name = self.selected_sheet()

        if not sheet_name:
            return

        column_number = (
            current_column + 1
        )

        column_letter = (
            self.column_letter(
                column_number
            )
        )

        answer = QMessageBox.question(
            self,
            "Delete Column",
            (
                f"Are you sure you want to "
                f"delete column {column_letter}?"
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if (
            answer
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.delete_column_button.setEnabled(
            False
        )

        self.delete_column_button.setText(
            "Deleting..."
        )

        try:
            result = (
                self.api_client.delete_column(
                    file_id=self.file_id,
                    sheet_name=sheet_name,
                    column_number=column_number,
                )
            )

            # -------------------------------------------------
            # Local workbook
            # -------------------------------------------------

            if self.workbook:
                worksheet = self.workbook[
                    sheet_name
                ]

                worksheet.delete_cols(
                    column_number,
                    1,
                )

            # -------------------------------------------------
            # UI
            # -------------------------------------------------

            self._loading_table = True

            self.table.blockSignals(
                True
            )

            try:
                self.table.removeColumn(
                    current_column
                )

            finally:
                self.table.blockSignals(
                    False
                )

                self._loading_table = False

            # -------------------------------------------------
            # Shift pending changes
            # -------------------------------------------------

            self._shift_pending_changes_after_column_delete(
                sheet_name,
                column_number,
            )

            # -------------------------------------------------
            # Refresh search
            # -------------------------------------------------

            if self.search_input.text().strip():
                self.search_excel()

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
    # Shift Pending Changes After Column Delete
    # =========================================================

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
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            match = re.match(
                r"^([A-Za-z]+)(\d+)$",
                cell,
            )

            if not match:
                updated[
                    (
                        change_sheet,
                        cell,
                    )
                ] = value

                continue

            letters = match.group(
                1
            )

            row_number = match.group(
                2
            )

            old_column = self.column_number(
                letters
            )

            # Change belonged to deleted column.
            if old_column == deleted_column:
                continue

            if old_column > deleted_column:
                old_column -= 1

            new_cell = (
                self.column_letter(
                    old_column
                )
                + row_number
            )

            updated[
                (
                    change_sheet,
                    new_cell,
                )
            ] = value

        self.pending_changes = updated

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

        sheet_name, ok = (
            QInputDialog.getText(
                self,
                "Create Sheet",
                "Enter new sheet name:",
            )
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

        self.create_sheet_button.setEnabled(
            False
        )

        self.create_sheet_button.setText(
            "Creating..."
        )

        try:
            result = (
                self.api_client.create_sheet(
                    file_id=self.file_id,
                    sheet_name=sheet_name,
                )
            )

            # -------------------------------------------------
            # Local workbook
            # -------------------------------------------------

            self.workbook.create_sheet(
                sheet_name
            )

            # -------------------------------------------------
            # Combo box
            # -------------------------------------------------

            self.sheet_combo.blockSignals(
                True
            )

            try:
                self.sheet_combo.addItem(
                    sheet_name
                )

            finally:
                self.sheet_combo.blockSignals(
                    False
                )

            self.sheet_combo.setCurrentText(
                sheet_name
            )

            self.load_sheet(
                sheet_name
            )

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
            self.create_sheet_button.setEnabled(
                True
            )

            self.create_sheet_button.setText(
                "New Sheet"
            )

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

        new_sheet_name, ok = (
            QInputDialog.getText(
                self,
                "Rename Sheet",
                "Enter new sheet name:",
                text=current_sheet,
            )
        )

        if not ok:
            return

        new_sheet_name = (
            new_sheet_name.strip()
        )

        if not new_sheet_name:
            QMessageBox.warning(
                self,
                "Rename Sheet",
                "Sheet name cannot be empty.",
            )

            return

        if new_sheet_name == current_sheet:
            return

        if (
            new_sheet_name
            in self.workbook.sheetnames
        ):
            QMessageBox.warning(
                self,
                "Rename Sheet",
                (
                    f"Sheet '{new_sheet_name}' "
                    "already exists."
                ),
            )

            return

        self.rename_sheet_button.setEnabled(
            False
        )

        self.rename_sheet_button.setText(
            "Renaming..."
        )

        try:
            result = (
                self.api_client.rename_sheet(
                    file_id=self.file_id,
                    sheet_name=current_sheet,
                    new_sheet_name=new_sheet_name,
                )
            )

            # -------------------------------------------------
            # Local workbook
            # -------------------------------------------------

            worksheet = self.workbook[
                current_sheet
            ]

            worksheet.title = (
                new_sheet_name
            )

            # -------------------------------------------------
            # Update combo
            # -------------------------------------------------

            self.sheet_combo.blockSignals(
                True
            )

            try:
                current_index = (
                    self.sheet_combo.currentIndex()
                )

                self.sheet_combo.setItemText(
                    current_index,
                    new_sheet_name,
                )

            finally:
                self.sheet_combo.blockSignals(
                    False
                )

            # -------------------------------------------------
            # Update pending changes
            # -------------------------------------------------

            updated = {}

            for (
                key,
                value,
            ) in self.pending_changes.items():

                old_sheet, cell = key

                if old_sheet == current_sheet:
                    updated[
                        (
                            new_sheet_name,
                            cell,
                        )
                    ] = value
                else:
                    updated[
                        key
                    ] = value

            self.pending_changes = updated

            self.load_sheet(
                new_sheet_name
            )

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
            self.rename_sheet_button.setEnabled(
                True
            )

            self.rename_sheet_button.setText(
                "Rename Sheet"
            )

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
            return

        if len(
            self.workbook.sheetnames
        ) == 1:
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

        if (
            answer
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.delete_sheet_button.setEnabled(
            False
        )

        self.delete_sheet_button.setText(
            "Deleting..."
        )

        try:
            result = (
                self.api_client.delete_sheet(
                    file_id=self.file_id,
                    sheet_name=current_sheet,
                )
            )

            # -------------------------------------------------
            # Remove from local workbook
            # -------------------------------------------------

            worksheet = self.workbook[
                current_sheet
            ]

            self.workbook.remove(
                worksheet
            )

            # -------------------------------------------------
            # Remove pending changes
            # -------------------------------------------------

            self.pending_changes = {
                key: value
                for key, value
                in self.pending_changes.items()
                if key[0] != current_sheet
            }

            # -------------------------------------------------
            # Update combo
            # -------------------------------------------------

            self.sheet_combo.blockSignals(
                True
            )

            try:
                current_index = (
                    self.sheet_combo.currentIndex()
                )

                self.sheet_combo.removeItem(
                    current_index
                )

            finally:
                self.sheet_combo.blockSignals(
                    False
                )

            # -------------------------------------------------
            # Select another sheet
            # -------------------------------------------------

            if self.workbook.sheetnames:
                new_index = min(
                    current_index,
                    len(
                        self.workbook.sheetnames
                    ) - 1,
                )

                self.sheet_combo.setCurrentIndex(
                    new_index
                )

                self.load_sheet(
                    self.sheet_combo.currentText()
                )

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
            self.delete_sheet_button.setEnabled(
                True
            )

            self.delete_sheet_button.setText(
                "Delete Sheet"
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
    # Excel Column Number
    # =========================================================

    def column_number(
        self,
        column_letters: str,
    ) -> int:
        result = 0

        for char in column_letters.upper():
            if not char.isalpha():
                continue

            result = (
                result * 26
                + (
                    ord(char)
                    - ord("A")
                    + 1
                )
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