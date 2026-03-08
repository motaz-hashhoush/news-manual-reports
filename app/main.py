import os
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from app.timezone import now_palestine
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from apscheduler.schedulers.background import BackgroundScheduler

from app.database import engine, SessionLocal, Base
from app.models import User, ReportSession, DataEntry, GeneratedReport, BreakingNews
from app.auth import hash_password
from app.report_generator import generate_docx_report


def check_session_deadlines():
    """Check all active sessions and close those past their deadline after generating a report."""
    db = SessionLocal()
    try:
        now = now_palestine()
        expired_sessions = (
            db.query(ReportSession)
            .filter(
                ReportSession.status == "active",
                ReportSession.deadline_at.isnot(None),
                ReportSession.deadline_at <= now,
            )
            .all()
        )
        for session in expired_sessions:
            entries = (
                db.query(DataEntry)
                .filter(DataEntry.session_id == session.id)
                .order_by(DataEntry.monitoring_time)
                .all()
            )
            report_type = f"{session.duration_hours}h"
            bn_count = db.query(BreakingNews).filter(BreakingNews.session_id == session.id).count()
            if entries:
                file_path = generate_docx_report(session, entries, breaking_news_count=bn_count)
                report = GeneratedReport(
                    session_id=session.id,
                    file_path=file_path,
                    report_type=report_type,
                )
                db.add(report)
            session.status = "closed"
            print(f"Session '{session.name}' (id={session.id}) closed — deadline reached.")
        db.commit()
    except Exception as e:
        print(f"Deadline check error: {e}")
        db.rollback()
    finally:
        db.close()


scheduler = BackgroundScheduler()
scheduler.add_job(
    check_session_deadlines,
    "interval",
    minutes=1,
    id="deadline_checker",
)


def _ensure_screenshot_data_columns():
    """Add screenshot_data column to existing tables if missing."""
    from sqlalchemy import inspect as sa_inspect, text
    inspector = sa_inspect(engine)
    for table_name in ("data_entries", "breaking_news"):
        columns = [c["name"] for c in inspector.get_columns(table_name)]
        if "screenshot_data" not in columns:
            with engine.begin() as conn:
                conn.execute(text(
                    f'ALTER TABLE {table_name} ADD COLUMN screenshot_data BYTEA'
                ))
            print(f"Added screenshot_data column to {table_name}")


@asynccontextmanager
async def lifespan(app: FastAPI):
    Base.metadata.create_all(bind=engine)
    _ensure_screenshot_data_columns()

    db = SessionLocal()
    try:
        admin = db.query(User).filter(User.username == "admin").first()
        if not admin:
            admin = User(
                username="admin",
                password_hash=hash_password("admin123"),
                role="admin",
            )
            db.add(admin)
            db.commit()
            print("Default admin user created (admin / admin123)")
    finally:
        db.close()

    scheduler.start()
    yield
    scheduler.shutdown()


app = FastAPI(title="Summary Report System", lifespan=lifespan)

static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
app.mount("/static", StaticFiles(directory=static_dir), name="static")

templates_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
app.state.templates = Jinja2Templates(directory=templates_dir)

from app.routes.auth_routes import router as auth_router
from app.routes.dashboard import router as dashboard_router
from app.routes.admin_routes import router as admin_router
from app.routes.user_routes import router as user_router

app.include_router(auth_router)
app.include_router(dashboard_router)
app.include_router(admin_router)
app.include_router(user_router)


@app.get("/")
async def root():
    return RedirectResponse(url="/dashboard", status_code=303)
