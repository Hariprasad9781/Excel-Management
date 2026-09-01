from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from desktop.services.api_client import ApiClient
from desktop.windows.dashboard_window import DashboardWindow


class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Management - Login")
        self.setFixedSize(500, 500)

        self.api_client = ApiClient()
        self.dashboard_window = None

        self.setup_ui()

        self.login_button.clicked.connect(self.handle_login)

    def setup_ui(self):
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(60, 50, 60, 50)
        main_layout.setSpacing(20)

        # Title
        title = QLabel("Excel Management")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title.setStyleSheet(
            """
            QLabel {
                font-size: 28px;
                font-weight: bold;
            }
            """
        )

        # Subtitle
        subtitle = QLabel("Sign in to your account")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)

        subtitle.setStyleSheet(
            """
            QLabel {
                font-size: 14px;
            }
            """
        )

        # Login card
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)

        card_layout = QVBoxLayout()
        card_layout.setContentsMargins(30, 30, 30, 30)
        card_layout.setSpacing(12)

        # Username label
        username_label = QLabel("Username")

        # Username input
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Enter your username")
        self.username_input.setMinimumHeight(40)

        # Password label
        password_label = QLabel("Password")

        # Password input
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Enter your password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setMinimumHeight(40)

        # Login button
        self.login_button = QPushButton("Login")
        self.login_button.setMinimumHeight(42)

        # Register section
        register_layout = QHBoxLayout()

        register_text = QLabel("Don't have an account?")

        self.register_button = QPushButton("Register")
        self.register_button.setFlat(True)

        register_layout.addStretch()
        register_layout.addWidget(register_text)
        register_layout.addWidget(self.register_button)
        register_layout.addStretch()

        # Add widgets to card
        card_layout.addWidget(username_label)
        card_layout.addWidget(self.username_input)

        card_layout.addSpacing(8)

        card_layout.addWidget(password_label)
        card_layout.addWidget(self.password_input)

        card_layout.addSpacing(15)

        card_layout.addWidget(self.login_button)

        card_layout.addSpacing(10)

        card_layout.addLayout(register_layout)

        card.setLayout(card_layout)

        # Add everything to main layout
        main_layout.addWidget(title)
        main_layout.addWidget(subtitle)
        main_layout.addSpacing(10)
        main_layout.addWidget(card)

        self.setLayout(main_layout)

    def handle_login(self):
        username = self.username_input.text().strip()
        password = self.password_input.text()

        if not username:
            QMessageBox.warning(
                self,
                "Login",
                "Please enter your username.",
            )
            return

        if not password:
            QMessageBox.warning(
                self,
                "Login",
                "Please enter your password.",
            )
            return

        self.login_button.setEnabled(False)
        self.login_button.setText("Logging in...")

        try:
            result = self.api_client.login(
                username=username,
                password=password,
            )

            print("Login response:", result)

            self.dashboard_window = DashboardWindow(
                username=username,
                api_client=self.api_client,
            )

            self.dashboard_window.show()

            self.close()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Login Failed",
                str(exc),
            )

        finally:
            self.login_button.setEnabled(True)
            self.login_button.setText("Login")