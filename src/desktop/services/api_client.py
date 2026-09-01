import os

import requests


class ApiClient:
    """
    HTTP client used by the desktop application
    to communicate with the FastAPI backend.
    """

    def __init__(
        self,
        base_url: str = "http://127.0.0.1:8000",
    ):
        self.base_url = base_url.rstrip("/")
        self.access_token: str | None = None

    # =========================================================
    # Authentication
    # =========================================================

    def login(
        self,
        username: str,
        password: str,
    ) -> dict:
        """
        Login and store the JWT access token.
        """

        try:
            response = requests.post(
                f"{self.base_url}/auth/login",
                json={
                    "username": username,
                    "password": password,
                },
                timeout=10,
            )
        except requests.RequestException as exc:
            raise Exception(
                f"Unable to connect to server: {exc}"
            ) from exc

        data = self._handle_response(
            response=response,
            expected_status=200,
            default_message="Login failed",
        )

        self.access_token = data["access_token"]

        return data

    def register(
        self,
        username: str,
        email: str,
        password: str,
    ) -> dict:
        """
        Register a new user.
        """

        try:
            response = requests.post(
                f"{self.base_url}/auth/register",
                json={
                    "username": username,
                    "email": email,
                    "password": password,
                },
                timeout=10,
            )
        except requests.RequestException as exc:
            raise Exception(
                f"Unable to connect to server: {exc}"
            ) from exc

        return self._handle_response(
            response=response,
            expected_status=201,
            default_message="Registration failed",
        )

    def logout(self) -> None:
        """
        Clear the locally stored access token.
        """

        self.access_token = None

    # =========================================================
    # Authentication Helpers
    # =========================================================

    def _require_login(self) -> None:
        """
        Make sure the user is authenticated.
        """

        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

    def _headers(self) -> dict:
        """
        Return authorization headers.
        """

        self._require_login()

        return {
            "Authorization": (
                f"Bearer {self.access_token}"
            ),
        }

    # =========================================================
    # Response Handling
    # =========================================================

    @staticmethod
    def _get_error_detail(
        response: requests.Response,
        default_message: str,
    ) -> str:
        """
        Extract FastAPI's 'detail' message when available.
        """

        try:
            data = response.json()

            if isinstance(data, dict):
                detail = data.get("detail")

                if detail:
                    return str(detail)

        except ValueError:
            pass

        return default_message

    def _handle_response(
        self,
        response: requests.Response,
        expected_status: int,
        default_message: str,
    ):
        """
        Validate an HTTP response and return JSON data.
        """

        if response.status_code != expected_status:
            detail = self._get_error_detail(
                response=response,
                default_message=default_message,
            )

            raise Exception(
                f"{detail} "
                f"(HTTP {response.status_code})"
            )

        try:
            return response.json()

        except ValueError as exc:
            raise Exception(
                "Server returned an invalid response."
            ) from exc

    # =========================================================
    # Get Current User
    # =========================================================

    def get_current_user(self) -> dict:
        """
        Get the currently authenticated user.
        """

        response = requests.get(
            f"{self.base_url}/auth/me",
            headers=self._headers(),
            timeout=10,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to fetch current user",
        )

    # =========================================================
    # File Operations
    # =========================================================

    def get_files(self) -> list[dict]:
        """
        Get files belonging to the logged-in user.
        """

        response = requests.get(
            f"{self.base_url}/files/",
            headers=self._headers(),
            timeout=10,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to fetch files",
        )

    def get_file_details(
        self,
        file_id: int,
    ) -> dict:
        """
        Get details of a specific Excel file.
        """

        response = requests.get(
            f"{self.base_url}/files/{file_id}",
            headers=self._headers(),
            timeout=10,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to fetch file details",
        )

    def upload_file(
        self,
        file_path: str,
    ) -> dict:
        """
        Upload an Excel file.
        """

        if not os.path.isfile(file_path):
            raise Exception(
                f"File not found: {file_path}"
            )

        try:
            with open(
                file_path,
                "rb",
            ) as file:

                response = requests.post(
                    f"{self.base_url}/files/upload",
                    headers=self._headers(),
                    files={
                        "file": file,
                    },
                    timeout=30,
                )

        except OSError as exc:
            raise Exception(
                f"Could not open file: {exc}"
            ) from exc

        except requests.RequestException as exc:
            raise Exception(
                f"File upload failed: {exc}"
            ) from exc

        return self._handle_response(
            response=response,
            expected_status=201,
            default_message="File upload failed",
        )

    def download_file(
        self,
        file_id: int,
        save_path: str,
    ) -> None:
        """
        Download an Excel file to the specified local path.
        """

        try:
            response = requests.get(
                f"{self.base_url}/files/{file_id}/download",
                headers=self._headers(),
                timeout=30,
            )

        except requests.RequestException as exc:
            raise Exception(
                f"File download failed: {exc}"
            ) from exc

        if response.status_code != 200:
            detail = self._get_error_detail(
                response=response,
                default_message="File download failed",
            )

            raise Exception(
                f"{detail} "
                f"(HTTP {response.status_code})"
            )

        try:
            parent_directory = os.path.dirname(
                os.path.abspath(save_path)
            )

            os.makedirs(
                parent_directory,
                exist_ok=True,
            )

            with open(
                save_path,
                "wb",
            ) as file:

                file.write(
                    response.content
                )

        except OSError as exc:
            raise Exception(
                f"Could not save downloaded file: {exc}"
            ) from exc

    def delete_file(
        self,
        file_id: int,
    ) -> dict:
        """
        Delete an Excel file.
        """

        response = requests.delete(
            f"{self.base_url}/files/{file_id}",
            headers=self._headers(),
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="File deletion failed",
        )

    # =========================================================
    # Sheet Operations
    # =========================================================

    def get_sheets(
        self,
        file_id: int,
    ) -> dict:
        """
        Get all worksheet names.
        """

        response = requests.get(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to fetch sheets",
        )

    def create_sheet(
        self,
        file_id: int,
        sheet_name: str,
    ) -> dict:
        """
        Create a new worksheet.
        """

        response = requests.post(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=201,
            default_message="Failed to create sheet",
        )

    def rename_sheet(
        self,
        file_id: int,
        sheet_name: str,
        new_sheet_name: str,
    ) -> dict:
        """
        Rename an existing worksheet.
        """

        response = requests.put(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "new_sheet_name": new_sheet_name,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to rename sheet",
        )

    def delete_sheet(
        self,
        file_id: int,
        sheet_name: str,
    ) -> dict:
        """
        Delete an existing worksheet.
        """

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to delete sheet",
        )

    # =========================================================
    # Sheet Preview
    # =========================================================

    def get_sheet_preview(
        self,
        file_id: int,
        sheet_name: str,
        rows: int = 10,
    ) -> dict:
        """
        Get a preview of worksheet data.
        """

        if rows < 1 or rows > 100:
            raise Exception(
                "Rows must be between 1 and 100."
            )

        response = requests.get(
            f"{self.base_url}/files/{file_id}/preview",
            headers=self._headers(),
            params={
                "sheet_name": sheet_name,
                "rows": rows,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to fetch sheet preview",
        )

    # =========================================================
    # Cell Operations
    # =========================================================

    def update_cell(
        self,
        file_id: int,
        sheet_name: str,
        cell: str,
        value,
    ) -> dict:
        """
        Update a single Excel cell.
        """

        response = requests.put(
            f"{self.base_url}/files/{file_id}/cell",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "cell": cell,
                "value": value,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to update cell",
        )

    # =========================================================
    # Row Operations
    # =========================================================

    def add_row(
        self,
        file_id: int,
        sheet_name: str,
        row_data: list,
    ) -> dict:
        """
        Add a row to a worksheet.
        """

        response = requests.post(
            f"{self.base_url}/files/{file_id}/rows",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "row_data": row_data,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=201,
            default_message="Failed to add row",
        )

    def update_row(
        self,
        file_id: int,
        sheet_name: str,
        row_number: int,
        row_data: list,
    ) -> dict:
        """
        Update an existing row.
        """

        response = requests.put(
            f"{self.base_url}/files/{file_id}/rows",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "row_number": row_number,
                "row_data": row_data,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to update row",
        )

    def delete_row(
        self,
        file_id: int,
        sheet_name: str,
        row_number: int,
    ) -> dict:
        """
        Delete a row.
        """

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/rows",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "row_number": row_number,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to delete row",
        )

    # =========================================================
    # Column Operations
    # =========================================================

    def add_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
    ) -> dict:
        """
        Add a new column.
        """

        response = requests.post(
            f"{self.base_url}/files/{file_id}/columns",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "column_number": column_number,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=201,
            default_message="Failed to add column",
        )

    def update_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
        column_name: str,
    ) -> dict:
        """
        Update the header/name of a column.
        """

        response = requests.put(
            f"{self.base_url}/files/{file_id}/columns",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "column_number": column_number,
                "column_name": column_name,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to update column",
        )

    def delete_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
    ) -> dict:
        """
        Delete a column.
        """

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/columns",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "column_number": column_number,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to delete column",
        )

    # =========================================================
    # Excel Search
    # =========================================================

    def search_excel(
        self,
        file_id: int,
        sheet_name: str,
        search_term: str,
    ) -> dict:
        """
        Search an Excel worksheet.

        IMPORTANT:
        The backend endpoint is GET /search,
        so sheet_name and search_term are sent
        as query parameters.
        """

        if not search_term.strip():
            raise Exception(
                "Search term cannot be empty."
            )

        response = requests.get(
            f"{self.base_url}/files/{file_id}/search",
            headers=self._headers(),
            params={
                "sheet_name": sheet_name,
                "search_term": search_term,
            },
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to search Excel",
        )

    # =========================================================
    # Excel Formatting
    # =========================================================

    def format_excel_range(
        self,
        file_id: int,
        sheet_name: str,
        cell_range: str,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: str | None = None,
        font_size: float | None = None,
        font_color: str | None = None,
        fill_color: str | None = None,
        horizontal_alignment: str | None = None,
        vertical_alignment: str | None = None,
        number_format: str | None = None,
    ) -> dict:
        """
        Apply formatting to a single cell or range.
        """

        payload = {
            "sheet_name": sheet_name,
            "cell_range": cell_range,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "font_size": font_size,
            "font_color": font_color,
            "fill_color": fill_color,
            "horizontal_alignment": horizontal_alignment,
            "vertical_alignment": vertical_alignment,
            "number_format": number_format,
        }

        response = requests.put(
            f"{self.base_url}/files/{file_id}/format",
            headers=self._headers(),
            json=payload,
            timeout=30,
        )

        return self._handle_response(
            response=response,
            expected_status=200,
            default_message="Failed to format Excel cells",
        )