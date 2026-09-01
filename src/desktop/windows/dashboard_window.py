from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from desktop.windows.files_window import FilesWindow


class DashboardWindow(QWidget):
    def __init__(
        self,
        username: str,
        api_client,
    ):
        super().__init__()

        self.username = username
        self.api_client = api_client

        self.files_window = None

        self.setWindowTitle("Excel Management - Dashboard")
        self.setMinimumSize(1000, 650)

        self.setup_ui()

    def upload_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)",
        )

        if not file_path:
            return

        try:
            self.upload_button.setEnabled(False)
            self.upload_button.setText("Uploading...")

            result = self.api_client.upload_file(file_path)

            QMessageBox.information(
                self,
                "Upload Successful",
                result.get(
                    "message",
                    "File uploaded successfully",
                ),
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Upload Failed",
                str(exc),
            )

        finally:
            self.upload_button.setEnabled(True)
            self.upload_button.setText("Upload File")

    def setup_ui(self):
        # =====================================================
        # Main Layout
        # =====================================================

        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # =====================================================
        # Sidebar
        # =====================================================

        sidebar = QFrame()
        sidebar.setFixedWidth(220)

        sidebar_layout = QVBoxLayout()
        sidebar_layout.setContentsMargins(20, 30, 20, 30)
        sidebar_layout.setSpacing(10)

        # Application title
        app_title = QLabel("Excel Management")
        app_title.setStyleSheet(
            """
            QLabel {
                font-size: 20px;
                font-weight: bold;
            }
            """
        )

        sidebar_layout.addWidget(app_title)
        sidebar_layout.addSpacing(30)

        # Navigation buttons
        dashboard_button = QPushButton("Dashboard")

        self.files_button = QPushButton("My Files")

        self.upload_button = QPushButton("Upload File")
        settings_button = QPushButton("Settings")

        for button in (
            dashboard_button,
            self.files_button,
            self.upload_button,
            settings_button,
        ):
            button.setMinimumHeight(42)
            sidebar_layout.addWidget(button)

        sidebar_layout.addStretch()

        # Logout button
        self.logout_button = QPushButton("Logout")
        self.logout_button.setMinimumHeight(42)

        sidebar_layout.addWidget(self.logout_button)

        sidebar.setLayout(sidebar_layout)

        # =====================================================
        # Connect Navigation
        # =====================================================

        self.files_button.clicked.connect(self.open_files_window)
        self.upload_button.clicked.connect(self.upload_file)

        # =====================================================
        # Content Area
        # =====================================================

        content = QFrame()

        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(40, 40, 40, 40)
        content_layout.setSpacing(20)

        # Welcome message
        welcome_label = QLabel(
            f"Welcome, {self.username}!"
        )

        welcome_label.setStyleSheet(
            """
            QLabel {
                font-size: 30px;
                font-weight: bold;
            }
            """
        )

        # Subtitle
        subtitle = QLabel(
            "Manage your Excel files easily."
        )

        subtitle.setStyleSheet(
            """
            QLabel {
                font-size: 16px;
            }
            """
        )

        # =====================================================
        # Dashboard Cards
        # =====================================================

        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(20)

        files_card = self.create_card(
            "My Files",
            "View and manage your Excel files.",
        )

        upload_card = self.create_card(
            "Upload File",
            "Upload a new Excel workbook.",
        )

        cards_layout.addWidget(files_card)
        cards_layout.addWidget(upload_card)

        content_layout.addWidget(welcome_label)
        content_layout.addWidget(subtitle)
        content_layout.addSpacing(20)
        content_layout.addLayout(cards_layout)
        content_layout.addStretch()

        content.setLayout(content_layout)

        # =====================================================
        # Add Sidebar + Content
        # =====================================================

        main_layout.addWidget(sidebar)
        main_layout.addWidget(content)

        self.setLayout(main_layout)

    def create_card(
        self,
        title: str,
        description: str,
    ) -> QFrame:
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)

        card_layout = QVBoxLayout()
        card_layout.setContentsMargins(25, 25, 25, 25)
        card_layout.setSpacing(10)

        title_label = QLabel(title)

        title_label.setStyleSheet(
            """
            QLabel {
                font-size: 20px;
                font-weight: bold;
            }
            """
        )

        description_label = QLabel(description)
        description_label.setWordWrap(True)

        card_layout.addWidget(title_label)
        card_layout.addWidget(description_label)
        card_layout.addStretch()

        card.setLayout(card_layout)

        return card

    def open_files_window(self):
        self.files_window = FilesWindow(
            self.api_client
        )

        self.files_window.show()