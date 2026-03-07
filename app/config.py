import os
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://postgres:postgres@localhost:5432/summary_report")

SECRET_KEY = os.getenv("SECRET_KEY", "change-me-in-production-super-secret-key-2026")

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

DISTRIBUTION_VALUES = [
    "ضيف", "تقرير", "مذيع", "فيلر", "وول", "عاجل", "تحليل", "مراسل", "مسؤول",
]

TYPE_VALUES = [
    "رياضة", "أخبار", "اقتصاد",
]
