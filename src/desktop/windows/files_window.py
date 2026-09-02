from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from desktop.services.api_client import ApiClient
from desktop.windows.excel_editor_window import ExcelEditorWindow


class FilesWindow(QWidget):
    def __init__(self, api_client: ApiClient):
        super().__init__()

        self.api_client = api_client
        self.editor_window = None

        self.setWindowTitle(
            "Excel Management - My Files"
        )

        self.setMinimumSize(
            900,
            600,
        )

        self.setup_ui()
        self.load_files()

    # =========================================================
    # UI SETUP
    # =========================================================

    def setup_ui(self):
        main_layout = QVBoxLayout()

        main_layout.setContentsMargins(
            30,
            30,
            30,
            30,
        )

        main_layout.setSpacing(20)

        # =====================================================
        # Header
        # =====================================================

        header_layout = QHBoxLayout()

        title = QLabel("My Files")

        title.setStyleSheet(
            """
            QLabel {
                font-size: 28px;
                font-weight: bold;
            }
            """
        )

        self.refresh_button = QPushButton(
            "Refresh"
        )

        self.refresh_button.setMinimumHeight(
            40
        )

        self.refresh_button.clicked.connect(
            self.refresh_files
        )

        header_layout.addWidget(
            title
        )

        header_layout.addStretch()

        header_layout.addWidget(
            self.refresh_button
        )

        # =====================================================
        # Scroll Area
        # =====================================================

        self.scroll_area = QScrollArea()

        self.scroll_area.setWidgetResizable(
            True
        )

        self.files_container = QWidget()

        self.files_layout = QVBoxLayout()

        self.files_layout.setSpacing(
            12
        )

        self.files_layout.setContentsMargins(
            5,
            5,
            5,
            5,
        )

        self.files_container.setLayout(
            self.files_layout
        )

        self.scroll_area.setWidget(
            self.files_container
        )

        # =====================================================
        # Main Layout
        # =====================================================

        main_layout.addLayout(
            header_layout
        )

        main_layout.addWidget(
            self.scroll_area
        )

        self.setLayout(
            main_layout
        )

    # =========================================================
    # Load Files
    # =========================================================

    def load_files(self):
        """
        Load all Excel files belonging to the
        currently authenticated user.
        """

        try:
            self.refresh_button.setEnabled(
                False
            )

            self.refresh_button.setText(
                "Loading..."
            )

            files = self.api_client.get_files()

            self.clear_files()

            if not files:
                self.show_empty_message()
                return

            for file in files:
                file_card = self.create_file_card(
                    file
                )

                self.files_layout.addWidget(
                    file_card
                )

            self.files_layout.addStretch()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Files",
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
    # Refresh Files
    # =========================================================

    def refresh_files(self):
        """
        Refresh the file list.

        This method is also used by DashboardWindow
        after a successful upload.
        """

        self.load_files()

    # =========================================================
    # Empty State
    # =========================================================

    def show_empty_message(self):
        empty_label = QLabel(
            "No Excel files found."
        )

        empty_label.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        empty_label.setStyleSheet(
            """
            QLabel {
                font-size: 16px;
                padding: 40px;
            }
            """
        )

        self.files_layout.addWidget(
            empty_label
        )

        self.files_layout.addStretch()

    # =========================================================
    # Open File
    # =========================================================

    def open_file(
        self,
        file_id: int,
    ):
        """
        Open an Excel file in the editor.
        """

        try:
            files = self.api_client.get_files()

            selected_file = next(
                (
                    file
                    for file in files
                    if file["id"] == file_id
                ),
                None,
            )

            if not selected_file:
                QMessageBox.warning(
                    self,
                    "File Not Found",
                    "The selected file could not be found.",
                )

                self.refresh_files()
                return

            # -------------------------------------------------
            # If an editor is already open, close it first.
            # -------------------------------------------------

            if (
                self.editor_window is not None
                and self.editor_window.isVisible()
            ):
                self.editor_window.close()

            self.editor_window = (
                ExcelEditorWindow(
                    file_id=file_id,
                    filename=selected_file[
                        "filename"
                    ],
                    api_client=self.api_client,
                )
            )

            self.editor_window.show()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Open File",
                str(exc),
            )

    # =========================================================
    # Download File
    # =========================================================

    def download_file(
        self,
        file_id: int,
        filename: str,
    ):
        """
        Download an Excel file to a location selected
        by the user.
        """

        save_path, _ = (
            QFileDialog.getSaveFileName(
                self,
                "Save Excel File",
                filename,
                "Excel Files (*.xlsx *.xls)",
            )
        )

        if not save_path:
            return

        # -----------------------------------------------------
        # Make sure the correct extension is present.
        # -----------------------------------------------------

        if not save_path.lower().endswith(
            (".xlsx", ".xls")
        ):
            save_path += ".xlsx"

        try:
            self.api_client.download_file(
                file_id=file_id,
                save_path=save_path,
            )

            QMessageBox.information(
                self,
                "Download Successful",
                (
                    f"File downloaded successfully.\n\n"
                    f"Saved to:\n{save_path}"
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Download Failed",
                str(exc),
            )

    # =========================================================
    # Delete File
    # =========================================================

    def delete_file(
        self,
        file_id: int,
        filename: str,
    ):
        """
        Delete an Excel file after confirmation.
        """

        result = QMessageBox.question(
            self,
            "Delete File",
            (
                f"Are you sure you want to delete "
                f"'{filename}'?\n\n"
                "This action cannot be undone."
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if (
            result
            != QMessageBox.StandardButton.Yes
        ):
            return

        try:
            response = (
                self.api_client.delete_file(
                    file_id
                )
            )

            QMessageBox.information(
                self,
                "Delete Successful",
                response.get(
                    "message",
                    "File deleted successfully.",
                ),
            )

            # -------------------------------------------------
            # Refresh after deletion
            # -------------------------------------------------

            self.refresh_files()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Delete Failed",
                str(exc),
            )

    # =========================================================
    # Clear File List
    # =========================================================

    def clear_files(self):
        """
        Remove all existing file cards.
        """

        while self.files_layout.count():
            item = (
                self.files_layout.takeAt(0)
            )

            widget = item.widget()

            if widget is not None:
                widget.deleteLater()

    # =========================================================
    # Create File Card
    # =========================================================

    def create_file_card(
        self,
        file: dict,
    ) -> QFrame:
        """
        Create one UI card for an Excel file.
        """

        card = QFrame()

        card.setFrameShape(
            QFrame.Shape.StyledPanel
        )

        card.setStyleSheet(
            """
            QFrame {
                border: 1px solid #dddddd;
                border-radius: 8px;
                background-color: white;
            }

            QLabel {
                border: none;
            }

            QPushButton {
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 8px 12px;
                background-color: #f5f5f5;
            }

            QPushButton:hover {
                background-color: #e5e5e5;
            }

            QPushButton:pressed {
                background-color: #d5d5d5;
            }
            """
        )

        card_layout = QHBoxLayout()

        card_layout.setContentsMargins(
            20,
            15,
            20,
            15,
        )

        card_layout.setSpacing(
            10
        )

        # =====================================================
        # File Information
        # =====================================================

        info_layout = QVBoxLayout()

        filename = QLabel(
            file["filename"]
        )

        filename.setStyleSheet(
            """
            QLabel {
                font-size: 18px;
                font-weight: bold;
            }
            """
        )

        file_size = QLabel(
            "Size: "
            + self.format_file_size(
                file["file_size"]
            )
        )

        created_at = QLabel(
            f"Created: {file['created_at']}"
        )

        stored_filename = QLabel(
            f"Stored: {file['stored_filename']}"
        )

        stored_filename.setStyleSheet(
            """
            QLabel {
                font-size: 11px;
            }
            """
        )

        info_layout.addWidget(
            filename
        )

        info_layout.addWidget(
            file_size
        )

        info_layout.addWidget(
            created_at
        )

        info_layout.addWidget(
            stored_filename
        )

        # =====================================================
        # Open Button
        # =====================================================

        open_button = QPushButton(
            "Open"
        )

        open_button.setMinimumWidth(
            80
        )

        open_button.setMinimumHeight(
            40
        )

        open_button.clicked.connect(
            lambda checked=False,
            file_id=file["id"]:
            self.open_file(file_id)
        )

        # =====================================================
        # Download Button
        # =====================================================

        download_button = QPushButton(
            "Download"
        )

        download_button.setMinimumWidth(
            90
        )

        download_button.setMinimumHeight(
            40
        )

        download_button.clicked.connect(
            lambda checked=False,
            file_id=file["id"],
            filename=file["filename"]:
            self.download_file(
                file_id,
                filename,
            )
        )

        # =====================================================
        # Delete Button
        # =====================================================

        delete_button = QPushButton(
            "Delete"
        )

        delete_button.setMinimumWidth(
            80
        )

        delete_button.setMinimumHeight(
            40
        )

        delete_button.clicked.connect(
            lambda checked=False,
            file_id=file["id"],
            filename=file["filename"]:
            self.delete_file(
                file_id,
                filename,
            )
        )

        # =====================================================
        # Add to Card
        # =====================================================

        card_layout.addLayout(
            info_layout
        )

        card_layout.addStretch()

        card_layout.addWidget(
            open_button
        )

        card_layout.addWidget(
            download_button
        )

        card_layout.addWidget(
            delete_button
        )

        card.setLayout(
            card_layout
        )

        return card

    # =========================================================
    # File Size Formatting
    # =========================================================

    @staticmethod
    def format_file_size(
        size: int,
    ) -> str:

        if size < 1024:
            return f"{size} B"

        if size < 1024 * 1024:
            return (
                f"{size / 1024:.2f} KB"
            )

        if size < 1024 * 1024 * 1024:
            return (
                f"{size / (1024 * 1024):.2f} MB"
            )

        return (
            f"{size / (1024 * 1024 * 1024):.2f} GB"
        )

    # =========================================================
    # Close Event
    # =========================================================

    def closeEvent(self, event):
        """
        Close the Excel editor when the Files window
        is closed.
        """

        if (
            self.editor_window is not None
            and self.editor_window.isVisible()
        ):
            self.editor_window.close()

        event.accept()