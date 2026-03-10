import os
import uuid

from fastapi import APIRouter, Request, Depends, Form, UploadFile, File
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import ReportSession, DataEntry, BreakingNews
from app.deps import get_current_user
from app.config import UPLOAD_DIR, DISTRIBUTION_VALUES, TYPE_VALUES

router = APIRouter(prefix="/user")


@router.get("/sessions")
async def list_sessions(request: Request, db: Session = Depends(get_db)):
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    sessions = (
        db.query(ReportSession)
        .filter(ReportSession.status == "active")
        .order_by(ReportSession.created_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "user_sessions.html",
        {"request": request, "user": user, "sessions": sessions},
    )


@router.get("/entry/{session_id}")
async def entry_page(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(
        ReportSession.id == session_id,
        ReportSession.status == "active",
    ).first()
    if not report_session:
        return RedirectResponse(url="/user/sessions", status_code=303)

    return request.app.state.templates.TemplateResponse(
        "data_entry.html",
        _user_entry_ctx(request, user, report_session, db),
    )


def _user_entry_ctx(request, user, report_session, db, error=None, success=None):
    entries = (
        db.query(DataEntry)
        .filter(DataEntry.session_id == report_session.id)
        .order_by(DataEntry.monitoring_time)
        .all()
    )
    breaking_news_count = (
        db.query(BreakingNews)
        .filter(BreakingNews.session_id == report_session.id)
        .count()
    )
    return {
        "request": request,
        "user": user,
        "session": report_session,
        "distribution_values": DISTRIBUTION_VALUES,
        "type_values": TYPE_VALUES,
        "entries": entries,
        "breaking_news_count": breaking_news_count,
        "error": error,
        "success": success,
    }


@router.post("/entry/{session_id}")
async def submit_entry(
    session_id: int,
    request: Request,
    monitoring_time: str = Form(...),
    title: str = Form(...),
    program: str = Form(""),
    entry_type: str = Form(""),
    distribution: str = Form(""),
    guest_reporter_name: str = Form(""),
    publish_link: str = Form(""),
    importance: str = Form(""),
    clip_duration: str = Form(""),
    screenshot: UploadFile = File(None),
    db: Session = Depends(get_db),
):
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(
        ReportSession.id == session_id,
        ReportSession.status == "active",
    ).first()
    if not report_session:
        return RedirectResponse(url="/user/sessions", status_code=303)

    screenshot_path = None
    screenshot_data = None
    if screenshot and screenshot.filename:
        ext = os.path.splitext(screenshot.filename)[1]
        filename = f"{uuid.uuid4().hex}{ext}"
        save_path = os.path.join(UPLOAD_DIR, filename)
        content = await screenshot.read()
        with open(save_path, "wb") as f:
            f.write(content)
        screenshot_path = f"/static/uploads/{filename}"
        screenshot_data = content

    entry = DataEntry(
        session_id=session_id,
        user_id=user.id,
        monitoring_time=monitoring_time,
        title=title,
        program=program,
        entry_type=entry_type,
        distribution=distribution,
        guest_reporter_name=guest_reporter_name,
        publish_link=publish_link,
        importance=importance,
        clip_duration=clip_duration,
        screenshot_path=screenshot_path,
        screenshot_data=screenshot_data,
    )
    db.add(entry)
    db.commit()

    return request.app.state.templates.TemplateResponse(
        "data_entry.html",
        _user_entry_ctx(request, user, report_session, db, success="تم إضافة البيانات بنجاح"),
    )
