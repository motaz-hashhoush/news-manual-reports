import bcrypt
from itsdangerous import URLSafeSerializer

from app.config import SECRET_KEY

serializer = URLSafeSerializer(SECRET_KEY)


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(plain: str, hashed: str) -> bool:
    return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))


def create_session_token(user_id: int, role: str) -> str:
    return serializer.dumps({"user_id": user_id, "role": role})


def decode_session_token(token: str):
    try:
        return serializer.loads(token)
    except Exception:
        return None
