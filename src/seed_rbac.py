from database import SessionLocal
from models.permission import Permission
from models.role import Role
from models.role_permission import RolePermission


PERMISSIONS = [
    ("view_file", "View Excel files"),
    ("upload_file", "Upload Excel files"),
    ("edit_file", "Edit Excel files"),
    ("delete_file", "Delete Excel files"),
    ("download_file", "Download Excel files"),
    ("create_sheet", "Create worksheets"),
    ("edit_sheet", "Edit worksheets"),
    ("delete_sheet", "Delete worksheets"),
    ("restore_version", "Restore a previous workbook version"),
    ("view_version_history", "View workbook version history"),
]

ROLE_PERMISSIONS = {
    "Admin": [
        "view_file",
        "upload_file",
        "edit_file",
        "delete_file",
        "download_file",
        "create_sheet",
        "edit_sheet",
        "delete_sheet",
        "restore_version",
        "view_version_history",
    ],
    "Editor": [
        "view_file",
        "upload_file",
        "edit_file",
        "download_file",
        "create_sheet",
        "edit_sheet",
        "delete_sheet",
        "restore_version",
        "view_version_history",
    ],
    "Viewer": [
        "view_file",
        "download_file",
        "view_version_history",
    ],
}

ROLES = [
    ("Admin", "Full access to the application"),
    ("Editor", "Can view and modify Excel files"),
    ("Viewer", "Can view and download Excel files"),
]


def seed_rbac():
    db = SessionLocal()

    try:
        permission_map = {}

        # Create permissions
        for name, description in PERMISSIONS:
            permission = (
                db.query(Permission)
                .filter(Permission.name == name)
                .first()
            )

            if permission is None:
                permission = Permission(
                    name=name,
                    description=description,
                )
                db.add(permission)
                db.flush()

            permission_map[name] = permission

        # Create roles
        role_map = {}

        for name, description in ROLES:
            role = (
                db.query(Role)
                .filter(Role.name == name)
                .first()
            )

            if role is None:
                role = Role(
                    name=name,
                    description=description,
                )
                db.add(role)
                db.flush()

            role_map[name] = role

        # Create role-permission mappings
        for role_name, permission_names in ROLE_PERMISSIONS.items():
            role = role_map[role_name]

            for permission_name in permission_names:
                permission = permission_map[permission_name]

                mapping = (
                    db.query(RolePermission)
                    .filter(
                        RolePermission.role_id == role.id,
                        RolePermission.permission_id == permission.id,
                    )
                    .first()
                )

                if mapping is None:
                    db.add(
                        RolePermission(
                            role_id=role.id,
                            permission_id=permission.id,
                        )
                    )

        db.commit()

        print("RBAC seed completed successfully.")

        print("\nRoles:")
        for role in role_map.values():
            print(f"  {role.id}: {role.name}")

        print("\nPermissions:")
        for permission in permission_map.values():
            print(f"  {permission.id}: {permission.name}")

    except Exception:
        db.rollback()
        raise

    finally:
        db.close()


if __name__ == "__main__":
    seed_rbac()