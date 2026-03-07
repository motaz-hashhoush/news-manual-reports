from fastapi import Request
from sqlalchemy.orm import Session

from app.auth import decode_session_token
from app.models import User


def get_current_user(request: Request, db: Session) -> User | None:
    token = request.cookies.get("session_token")
    if not token:
        return None
    data = decode_session_token(token)
    if not data:
        return None
    return db.query(User).filter(User.id == data["user_id"]).first()


def require_admin(request: Request, db: Session) -> User | None:
    user = get_current_user(request, db)
    if user and user.role == "admin":
        return user
    return None
