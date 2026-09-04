from desktop.services.api_client import ApiClient


def main():
    client = ApiClient()

    # Use your existing login credentials
    client.login(
        username="hari_test2",
        password="Test@123",
    )

    result = client.restore_workbook_version(
        file_id=5,
        version_number=15,
    )

    print(result)


if __name__ == "__main__":
    main()