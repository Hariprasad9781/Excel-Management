from sqlalchemy.orm import Session

from models.user import User
from utils.security import hash_password, verify_password


def create_user(
    db: Session,
    username: str,
    email: str,
    password: str,
) -> User:
    password_hash = hash_password(password)

    user = User(
        username=username,
        email=email,
        password_hash=password_hash,
    )

    db.add(user)
    db.commit()
    db.refresh(user)

    return user


def authenticate_user(
    db: Session,
    username: str,
    password: str,
) -> User | None:
    user = (
        db.query(User)
        .filter(User.username == username)
        .first()
    )

    if not user:
        return None

    if not verify_password(
        password,
        user.password_hash,
    ):
        return None

    return user