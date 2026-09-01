import requests


class ApiClient:
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
        response = requests.post(
            f"{self.base_url}/auth/login",
            json={
                "username": username,
                "password": password,
            },
            timeout=10,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Login failed",
                )
            except ValueError:
                detail = "Login failed"

            raise Exception(detail)

        data = response.json()

        self.access_token = data["access_token"]

        return data

    # =========================================================
    # Headers
    # =========================================================

    def _headers(self) -> dict:
        if not self.access_token:
            return {}

        return {
            "Authorization": (
                f"Bearer {self.access_token}"
            ),
        }

    # =========================================================
    # Get Files
    # =========================================================

    def get_files(self) -> list[dict]:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.get(
            f"{self.base_url}/files/",
            headers=self._headers(),
            timeout=10,
        )

        print(
            "GET /files/ status:",
            response.status_code,
        )

        print(
            "GET /files/ response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to fetch files",
                )
            except ValueError:
                detail = "Failed to fetch files"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Upload File
    # =========================================================

    def upload_file(
        self,
        file_path: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
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

        if response.status_code != 201:
            try:
                detail = response.json().get(
                    "detail",
                    "File upload failed",
                )
            except ValueError:
                detail = "File upload failed"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Download File
    # =========================================================

    def download_file(
        self,
        file_id: int,
        save_path: str,
    ) -> None:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.get(
            f"{self.base_url}/files/{file_id}/download",
            headers=self._headers(),
            timeout=30,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "File download failed",
                )
            except ValueError:
                detail = "File download failed"

            raise Exception(detail)

        try:
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

    # =========================================================
    # Update Cell
    # =========================================================

    def update_cell(
        self,
        file_id: int,
        sheet_name: str,
        cell: str,
        value,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

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

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to update cell",
                )
            except ValueError:
                detail = "Failed to update cell"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Add Row
    # =========================================================

    def add_row(
        self,
        file_id: int,
        sheet_name: str,
        row_data: list,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.post(
            f"{self.base_url}/files/{file_id}/rows",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "row_data": row_data,
            },
            timeout=30,
        )

        print(
            "POST /rows status:",
            response.status_code,
        )

        print(
            "POST /rows response:",
            response.text,
        )

        if response.status_code != 201:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to add row",
                )
            except ValueError:
                detail = "Failed to add row"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Delete Row
    # =========================================================

    def delete_row(
        self,
        file_id: int,
        sheet_name: str,
        row_number: int,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/rows",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "row_number": row_number,
            },
            timeout=30,
        )

        print(
            "DELETE /rows status:",
            response.status_code,
        )

        print(
            "DELETE /rows response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to delete row",
                )
            except ValueError:
                detail = "Failed to delete row"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Update Row
    # =========================================================

    def update_row(
        self,
        file_id: int,
        sheet_name: str,
        row_number: int,
        row_data: list,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

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

        print(
            "PUT /rows status:",
            response.status_code,
        )

        print(
            "PUT /rows response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to update row",
                )
            except ValueError:
                detail = "Failed to update row"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Add Column
    # =========================================================

    def add_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.post(
            f"{self.base_url}/files/{file_id}/columns",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "column_number": column_number,
            },
            timeout=30,
        )

        print(
            "POST /columns status:",
            response.status_code,
        )

        print(
            "POST /columns response:",
            response.text,
        )

        if response.status_code != 201:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to add column",
                )
            except ValueError:
                detail = "Failed to add column"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Update Column / Rename Column
    # =========================================================

    def update_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
        column_name: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

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

        print(
            "PUT /columns status:",
            response.status_code,
        )

        print(
            "PUT /columns response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to update column",
                )
            except ValueError:
                detail = "Failed to update column"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Delete Column
    # =========================================================

    def delete_column(
        self,
        file_id: int,
        sheet_name: str,
        column_number: int,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/columns",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "column_number": column_number,
            },
            timeout=30,
        )

        print(
            "DELETE /columns status:",
            response.status_code,
        )

        print(
            "DELETE /columns response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to delete column",
                )
            except ValueError:
                detail = "Failed to delete column"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Get Sheet Names
    # =========================================================

    def get_sheets(
        self,
        file_id: int,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.get(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            timeout=30,
        )

        print(
            "GET /sheets status:",
            response.status_code,
        )

        print(
            "GET /sheets response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to fetch sheets",
                )
            except ValueError:
                detail = "Failed to fetch sheets"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Create Sheet
    # =========================================================

    def create_sheet(
        self,
        file_id: int,
        sheet_name: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.post(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
            },
            timeout=30,
        )

        print(
            "POST /sheets status:",
            response.status_code,
        )

        print(
            "POST /sheets response:",
            response.text,
        )

        if response.status_code != 201:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to create sheet",
                )
            except ValueError:
                detail = "Failed to create sheet"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Rename Sheet
    # =========================================================

    def rename_sheet(
        self,
        file_id: int,
        sheet_name: str,
        new_sheet_name: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.put(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
                "new_sheet_name": new_sheet_name,
            },
            timeout=30,
        )

        print(
            "PUT /sheets status:",
            response.status_code,
        )

        print(
            "PUT /sheets response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to rename sheet",
                )
            except ValueError:
                detail = "Failed to rename sheet"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Delete Sheet
    # =========================================================

    def delete_sheet(
        self,
        file_id: int,
        sheet_name: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        response = requests.delete(
            f"{self.base_url}/files/{file_id}/sheets",
            headers=self._headers(),
            json={
                "sheet_name": sheet_name,
            },
            timeout=30,
        )

        print(
            "DELETE /sheets status:",
            response.status_code,
        )

        print(
            "DELETE /sheets response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to delete sheet",
                )
            except ValueError:
                detail = "Failed to delete sheet"

            raise Exception(detail)

        return response.json()

    # =========================================================
    # Search Excel
    # =========================================================

    def search_excel(
        self,
        file_id: int,
        sheet_name: str,
        search_term: str,
    ) -> dict:
        if not self.access_token:
            raise Exception(
                "You are not logged in."
            )

        if not sheet_name.strip():
            raise Exception(
                "Sheet name cannot be empty."
            )

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

        print(
            "GET /search status:",
            response.status_code,
        )

        print(
            "GET /search response:",
            response.text,
        )

        if response.status_code != 200:
            try:
                detail = response.json().get(
                    "detail",
                    "Failed to search Excel",
                )
            except ValueError:
                detail = "Failed to search Excel"

            raise Exception(detail)

        return response.json()