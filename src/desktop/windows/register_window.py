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


class RegisterWindow(QWidget):
    def __init__(
        self,
        api_client: ApiClient,
        login_window=None,
    ):
        super().__init__()

        self.setWindowTitle(
            "Excel Management - Register"
        )

        self.setFixedSize(500, 600)

        # API client
        self.api_client = api_client

        # Login window reference
        self.login_window = login_window

        # Setup UI
        self.setup_ui()

        # Connect buttons
        self.register_button.clicked.connect(
            self.handle_register
        )

        self.back_button.clicked.connect(
            self.handle_back
        )

        # Allow pressing Enter in confirm password
        self.confirm_password_input.returnPressed.connect(
            self.handle_register
        )

    # =========================================================
    # UI Setup
    # =========================================================

    def setup_ui(self):
        # Main layout
        main_layout = QVBoxLayout()

        main_layout.setContentsMargins(
            60,
            40,
            60,
            40,
        )

        main_layout.setSpacing(20)

        # =====================================================
        # Title
        # =====================================================

        title = QLabel(
            "Create Account"
        )

        title.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        title.setStyleSheet(
            """
            QLabel {
                font-size: 28px;
                font-weight: bold;
            }
            """
        )

        # =====================================================
        # Subtitle
        # =====================================================

        subtitle = QLabel(
            "Register for Excel Management"
        )

        subtitle.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        subtitle.setStyleSheet(
            """
            QLabel {
                font-size: 14px;
            }
            """
        )

        # =====================================================
        # Registration Card
        # =====================================================

        card = QFrame()

        card.setFrameShape(
            QFrame.Shape.StyledPanel
        )

        card_layout = QVBoxLayout()

        card_layout.setContentsMargins(
            30,
            30,
            30,
            30,
        )

        card_layout.setSpacing(10)

        # =====================================================
        # Username
        # =====================================================

        username_label = QLabel(
            "Username"
        )

        self.username_input = QLineEdit()

        self.username_input.setPlaceholderText(
            "Enter your username"
        )

        self.username_input.setMinimumHeight(
            40
        )

        # =====================================================
        # Email
        # =====================================================

        email_label = QLabel(
            "Email"
        )

        self.email_input = QLineEdit()

        self.email_input.setPlaceholderText(
            "Enter your email"
        )

        self.email_input.setMinimumHeight(
            40
        )

        # =====================================================
        # Password
        # =====================================================

        password_label = QLabel(
            "Password"
        )

        self.password_input = QLineEdit()

        self.password_input.setPlaceholderText(
            "Enter your password"
        )

        self.password_input.setEchoMode(
            QLineEdit.EchoMode.Password
        )

        self.password_input.setMinimumHeight(
            40
        )

        # =====================================================
        # Confirm Password
        # =====================================================

        confirm_password_label = QLabel(
            "Confirm Password"
        )

        self.confirm_password_input = QLineEdit()

        self.confirm_password_input.setPlaceholderText(
            "Re-enter your password"
        )

        self.confirm_password_input.setEchoMode(
            QLineEdit.EchoMode.Password
        )

        self.confirm_password_input.setMinimumHeight(
            40
        )

        # =====================================================
        # Register Button
        # =====================================================

        self.register_button = QPushButton(
            "Register"
        )

        self.register_button.setMinimumHeight(
            42
        )

        # =====================================================
        # Back Button
        # =====================================================

        self.back_button = QPushButton(
            "Back to Login"
        )

        self.back_button.setFlat(
            True
        )

        # =====================================================
        # Add Widgets
        # =====================================================

        card_layout.addWidget(
            username_label
        )

        card_layout.addWidget(
            self.username_input
        )

        card_layout.addSpacing(5)

        card_layout.addWidget(
            email_label
        )

        card_layout.addWidget(
            self.email_input
        )

        card_layout.addSpacing(5)

        card_layout.addWidget(
            password_label
        )

        card_layout.addWidget(
            self.password_input
        )

        card_layout.addSpacing(5)

        card_layout.addWidget(
            confirm_password_label
        )

        card_layout.addWidget(
            self.confirm_password_input
        )

        card_layout.addSpacing(15)

        card_layout.addWidget(
            self.register_button
        )

        card_layout.addSpacing(5)

        card_layout.addWidget(
            self.back_button
        )

        card.setLayout(
            card_layout
        )

        # =====================================================
        # Main Layout
        # =====================================================

        main_layout.addWidget(
            title
        )

        main_layout.addWidget(
            subtitle
        )

        main_layout.addSpacing(5)

        main_layout.addWidget(
            card
        )

        self.setLayout(
            main_layout
        )

        # Initial focus
        self.username_input.setFocus()

    # =========================================================
    # Registration
    # =========================================================

    def handle_register(self):
        username = (
            self.username_input
            .text()
            .strip()
        )

        email = (
            self.email_input
            .text()
            .strip()
        )

        password = (
            self.password_input
            .text()
        )

        confirm_password = (
            self.confirm_password_input
            .text()
        )

        # -----------------------------------------------------
        # Validate username
        # -----------------------------------------------------

        if not username:
            QMessageBox.warning(
                self,
                "Registration",
                "Please enter a username.",
            )

            self.username_input.setFocus()

            return

        # -----------------------------------------------------
        # Validate email
        # -----------------------------------------------------

        if not email:
            QMessageBox.warning(
                self,
                "Registration",
                "Please enter your email.",
            )

            self.email_input.setFocus()

            return

        # -----------------------------------------------------
        # Validate password
        # -----------------------------------------------------

        if not password:
            QMessageBox.warning(
                self,
                "Registration",
                "Please enter a password.",
            )

            self.password_input.setFocus()

            return

        # -----------------------------------------------------
        # Validate password length
        # -----------------------------------------------------

        if len(
            password.encode("utf-8")
        ) > 72:
            QMessageBox.warning(
                self,
                "Registration",
                "Password cannot be longer than "
                "72 bytes.",
            )

            self.password_input.setFocus()

            return

        # -----------------------------------------------------
        # Validate confirmation
        # -----------------------------------------------------

        if not confirm_password:
            QMessageBox.warning(
                self,
                "Registration",
                "Please confirm your password.",
            )

            self.confirm_password_input.setFocus()

            return

        # -----------------------------------------------------
        # Compare passwords
        # -----------------------------------------------------

        if password != confirm_password:
            QMessageBox.warning(
                self,
                "Registration",
                "Passwords do not match.",
            )

            self.confirm_password_input.clear()
            self.confirm_password_input.setFocus()

            return

        # -----------------------------------------------------
        # Disable buttons
        # -----------------------------------------------------

        self.register_button.setEnabled(
            False
        )

        self.back_button.setEnabled(
            False
        )

        self.register_button.setText(
            "Registering..."
        )

        try:
            # -------------------------------------------------
            # Call backend
            # -------------------------------------------------

            result = self.api_client.register(
                username=username,
                email=email,
                password=password,
            )

            print(
                "Registration response:",
                result,
            )

            # -------------------------------------------------
            # Success
            # -------------------------------------------------

            QMessageBox.information(
                self,
                "Registration Successful",
                "Your account has been created "
                "successfully.\n\n"
                "Please login with your new account.",
            )

            # Return to Login
            self.close()

            if self.login_window:
                self.login_window.show()

                self.login_window.username_input.setText(
                    username
                )

                self.login_window.password_input.clear()

                self.login_window.password_input.setFocus()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Registration Failed",
                str(exc),
            )

        finally:
            # -------------------------------------------------
            # Restore buttons
            # -------------------------------------------------

            self.register_button.setEnabled(
                True
            )

            self.back_button.setEnabled(
                True
            )

            self.register_button.setText(
                "Register"
            )

    # =========================================================
    # Back to Login
    # =========================================================

    def handle_back(self):
        self.close()

        if self.login_window:
            self.login_window.show()

            self.login_window.username_input.setFocus()