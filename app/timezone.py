from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo

PALESTINE_TZ = ZoneInfo("Asia/Hebron")


def now_palestine() -> datetime:
    """Return current Palestine local time as a naive datetime
    so PostgreSQL stores it as-is without UTC conversion."""
    return datetime.now(PALESTINE_TZ).replace(tzinfo=None)
