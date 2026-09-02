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

    # =========================================================
    # UI Setup
    # =========================================================

    def setup_ui(self):
        main_layout = QVBoxLayout()

        main_layout.setContentsMargins(
            60,
            40,
            60,
            40,
        )

        main_layout.setSpacing(20)

        # =====================================================
        # Top Header
        # =====================================================

        header_layout = QHBoxLayout()

        app_title = QLabel("Excel Management")

        app_title.setStyleSheet(
            """
            QLabel {
                font-size: 22px;
                font-weight: bold;
            }
            """
        )

        self.logout_button = QPushButton("Logout")

        self.logout_button.setMinimumHeight(40)

        self.logout_button.clicked.connect(
            self.handle_logout
        )

        header_layout.addWidget(app_title)
        header_layout.addStretch()
        header_layout.addWidget(self.logout_button)

        # =====================================================
        # Welcome Section
        # =====================================================

        welcome_label = QLabel(
            f"Welcome, {self.username}!"
        )

        welcome_label.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        welcome_label.setStyleSheet(
            """
            QLabel {
                font-size: 32px;
                font-weight: bold;
            }
            """
        )

        subtitle = QLabel(
            "Manage your Excel files easily."
        )

        subtitle.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        subtitle.setStyleSheet(
            """
            QLabel {
                font-size: 16px;
            }
            """
        )

        # =====================================================
        # Main Cards
        # =====================================================

        cards_layout = QHBoxLayout()

        cards_layout.setSpacing(25)

        # My Files Card
        files_card = self.create_card(
            title="My Files",
            description=(
                "View, download, delete and manage "
                "your Excel files."
            ),
            button_text="Open My Files",
            callback=self.open_files_window,
        )

        # Upload Card
        upload_card = self.create_card(
            title="Upload File",
            description=(
                "Upload a new .xlsx or .xls "
                "Excel workbook."
            ),
            button_text="Choose Excel File",
            callback=self.upload_file,
        )

        cards_layout.addWidget(files_card)
        cards_layout.addWidget(upload_card)

        # =====================================================
        # Layout
        # =====================================================

        main_layout.addLayout(header_layout)

        main_layout.addSpacing(50)

        main_layout.addWidget(
            welcome_label
        )

        main_layout.addWidget(
            subtitle
        )

        main_layout.addSpacing(30)

        main_layout.addLayout(
            cards_layout
        )

        main_layout.addStretch()

        self.setLayout(main_layout)

    # =========================================================
    # Create Card
    # =========================================================

    def create_card(
        self,
        title: str,
        description: str,
        button_text: str,
        callback,
    ) -> QFrame:

        card = QFrame()

        card.setFrameShape(
            QFrame.Shape.StyledPanel
        )

        card.setMinimumHeight(280)

        card_layout = QVBoxLayout()

        card_layout.setContentsMargins(
            30,
            30,
            30,
            30,
        )

        card_layout.setSpacing(15)

        # -----------------------------------------------------
        # Title
        # -----------------------------------------------------

        title_label = QLabel(title)

        title_label.setStyleSheet(
            """
            QLabel {
                font-size: 24px;
                font-weight: bold;
            }
            """
        )

        # -----------------------------------------------------
        # Description
        # -----------------------------------------------------

        description_label = QLabel(
            description
        )

        description_label.setWordWrap(True)

        description_label.setStyleSheet(
            """
            QLabel {
                font-size: 15px;
            }
            """
        )

        # -----------------------------------------------------
        # Button
        # -----------------------------------------------------

        action_button = QPushButton(
            button_text
        )

        action_button.setMinimumHeight(48)

        action_button.setStyleSheet(
            """
            QPushButton {
                font-size: 15px;
                font-weight: bold;
                padding: 8px;
            }

            QPushButton:hover {
                background-color: #e8e8e8;
            }

            QPushButton:pressed {
                background-color: #dcdcdc;
            }
            """
        )

        action_button.clicked.connect(
            callback
        )

        # -----------------------------------------------------
        # Add Widgets
        # -----------------------------------------------------

        card_layout.addWidget(
            title_label
        )

        card_layout.addWidget(
            description_label
        )

        card_layout.addStretch()

        card_layout.addWidget(
            action_button
        )

        card.setLayout(
            card_layout
        )

        return card

    # =========================================================
    # Upload File
    # =========================================================

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

            self.setEnabled(False)

            result = self.api_client.upload_file(
                file_path
            )

            QMessageBox.information(
                self,
                "Upload Successful",
                result.get(
                    "message",
                    "File uploaded successfully.",
                ),
            )

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Upload Failed",
                str(exc),
            )

        finally:

            self.setEnabled(True)

    # =========================================================
    # Open Files
    # =========================================================

    def open_files_window(self):

        try:

            self.files_window = FilesWindow(
                self.api_client
            )

            self.files_window.show()

        except Exception as exc:

            QMessageBox.critical(
                self,
                "Failed to Open Files",
                str(exc),
            )

    # =========================================================
    # Logout
    # =========================================================

    def handle_logout(self):

        reply = QMessageBox.question(
            self,
            "Logout",
            "Are you sure you want to logout?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        self.api_client.logout()

        self.close()