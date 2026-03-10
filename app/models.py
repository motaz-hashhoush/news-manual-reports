from sqlalchemy import (
    Column, Integer, String, Text, DateTime, ForeignKey, Enum as SAEnum,
    LargeBinary,
)
from sqlalchemy.orm import relationship

from app.database import Base
from app.timezone import now_palestine


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(150), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    role = Column(String(20), nullable=False, default="user")  # "admin" or "user"
    created_at = Column(DateTime, default=now_palestine)

    entries = relationship("DataEntry", back_populates="user")
    sessions_created = relationship("ReportSession", back_populates="created_by_user")
    breaking_news_items = relationship("BreakingNews", back_populates="user")


class ReportSession(Base):
    __tablename__ = "report_sessions"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(255), nullable=False)
    description = Column(Text, nullable=True)
    created_by = Column(Integer, ForeignKey("users.id"), nullable=False)
    status = Column(String(20), default="active")  # active, closed
    duration_hours = Column(Integer, nullable=False, default=24)  # 12 or 24
    start_at = Column(DateTime, nullable=True)
    deadline_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=now_palestine)

    created_by_user = relationship("User", back_populates="sessions_created")
    entries = relationship("DataEntry", back_populates="session", order_by="DataEntry.monitoring_time")
    generated_reports = relationship("GeneratedReport", back_populates="session")
    breaking_news_items = relationship("BreakingNews", back_populates="session")


class DataEntry(Base):
    __tablename__ = "data_entries"

    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("report_sessions.id"), nullable=False)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)

    monitoring_time = Column(String(50), nullable=False)       # فترة الرصد
    title = Column(Text, nullable=False)                       # العنوان / NOW
    program = Column(String(255), nullable=True)               # البرنامج
    entry_type = Column(String(50), nullable=True)             # النوع
    distribution = Column(String(50), nullable=True)           # التوزيع
    guest_reporter_name = Column(String(255), nullable=True)   # اسم الضيف/المراسل
    publish_link = Column(Text, nullable=True)                 # رابط النشر
    importance = Column(String(100), nullable=True)            # الأهمية
    clip_duration = Column(String(50), nullable=True)          # المدة الزمنية للمقطع
    screenshot_path = Column(String(500), nullable=True)       # برنت سكرين
    screenshot_data = Column(LargeBinary, nullable=True)

    created_at = Column(DateTime, default=now_palestine)

    session = relationship("ReportSession", back_populates="entries")
    user = relationship("User", back_populates="entries")


class BreakingNews(Base):
    __tablename__ = "breaking_news"

    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("report_sessions.id"), nullable=False)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    description = Column(Text, nullable=True)
    screenshot_path = Column(String(500), nullable=True)
    screenshot_data = Column(LargeBinary, nullable=True)
    created_at = Column(DateTime, default=now_palestine)

    session = relationship("ReportSession", back_populates="breaking_news_items")
    user = relationship("User", back_populates="breaking_news_items")


class LookupValue(Base):
    __tablename__ = "lookup_values"

    id = Column(Integer, primary_key=True, index=True)
    category = Column(String(50), nullable=False, index=True)
    name = Column(String(300), nullable=False)
    created_at = Column(DateTime, default=now_palestine)


class GeneratedReport(Base):
    __tablename__ = "generated_reports"

    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("report_sessions.id"), nullable=False)
    file_path = Column(String(500), nullable=False)
    report_type = Column(String(20), nullable=False)  # "on_demand", "12h", "24h"
    generated_at = Column(DateTime, default=now_palestine)

    session = relationship("ReportSession", back_populates="generated_reports")