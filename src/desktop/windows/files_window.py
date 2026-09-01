from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
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

        self.setWindowTitle("Excel Management - My Files")
        self.setMinimumSize(900, 600)

        self.setup_ui()
        self.load_files()

    def setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # Header
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

        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.setMinimumHeight(40)
        self.refresh_button.clicked.connect(self.load_files)

        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(self.refresh_button)

        # Scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.files_container = QWidget()
        self.files_layout = QVBoxLayout()
        self.files_layout.setSpacing(12)
        self.files_layout.setContentsMargins(5, 5, 5, 5)

        self.files_container.setLayout(self.files_layout)
        self.scroll_area.setWidget(self.files_container)

        main_layout.addLayout(header_layout)
        main_layout.addWidget(self.scroll_area)

        self.setLayout(main_layout)

    def load_files(self):
        try:
            self.refresh_button.setEnabled(False)
            self.refresh_button.setText("Loading...")

            files = self.api_client.get_files()

            self.clear_files()

            if not files:
                empty_label = QLabel("No Excel files found.")
                empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                empty_label.setStyleSheet(
                    """
                    QLabel {
                        font-size: 16px;
                    }
                    """
                )

                self.files_layout.addWidget(empty_label)
                self.files_layout.addStretch()

                return

            for file in files:
                file_card = self.create_file_card(file)
                self.files_layout.addWidget(file_card)

            self.files_layout.addStretch()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Load Files",
                str(exc),
            )

        finally:
            self.refresh_button.setEnabled(True)
            self.refresh_button.setText("Refresh")

    def open_file(self, file_id: int):
        try:
            # Find the selected file
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
                return

            self.editor_window = ExcelEditorWindow(
                file_id=file_id,
                filename=selected_file["filename"],
                api_client=self.api_client,
            )

            self.editor_window.show()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Failed to Open File",
                str(exc),
            )

    def clear_files(self):
        while self.files_layout.count():
            item = self.files_layout.takeAt(0)

            widget = item.widget()

            if widget:
                widget.deleteLater()

    def create_file_card(self, file: dict) -> QFrame:
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)

        card_layout = QHBoxLayout()
        card_layout.setContentsMargins(20, 15, 20, 15)

        # File information
        info_layout = QVBoxLayout()

        filename = QLabel(file["filename"])
        filename.setStyleSheet(
            """
            QLabel {
                font-size: 18px;
                font-weight: bold;
            }
            """
        )

        file_size = QLabel(
            f"Size: {self.format_file_size(file['file_size'])}"
        )

        created_at = QLabel(
            f"Created: {file['created_at']}"
        )

        info_layout.addWidget(filename)
        info_layout.addWidget(file_size)
        info_layout.addWidget(created_at)

        # Open button
        open_button = QPushButton("Open")
        open_button.setMinimumWidth(100)
        open_button.setMinimumHeight(40)

        open_button.clicked.connect(
            lambda checked=False, file_id=file["id"]:
            self.open_file(file_id)
        )

        card_layout.addLayout(info_layout)
        card_layout.addStretch()
        card_layout.addWidget(open_button)

        card.setLayout(card_layout)

        return card

    def format_file_size(self, size: int) -> str:
        if size < 1024:
            return f"{size} B"

        if size < 1024 * 1024:
            return f"{size / 1024:.2f} KB"

        return f"{size / (1024 * 1024):.2f} MB"

