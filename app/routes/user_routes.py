import os
import uuid

from fastapi import APIRouter, Request, Depends, Form, UploadFile, File, Query
from fastapi.responses import RedirectResponse, FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import ReportSession, DataEntry, BreakingNews, GeneratedReport
from app.deps import get_current_user
from app.config import UPLOAD_DIR, DISTRIBUTION_VALUES, TYPE_VALUES
from datetime import datetime
from app.report_generator import generate_custom_docx_report
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


@router.get("/archive")
async def archive_page(
    request: Request,
    page: int = Query(1, ge=1),
    db: Session = Depends(get_db)
):
    """Display archive of all generated reports"""
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    
    # Pagination settings
    page_size = 20
    offset = (page - 1) * page_size
    
    # Get total count
    total_reports = db.query(GeneratedReport).count()
    
    # Query reports with pagination
    reports_query = db.query(GeneratedReport).order_by(
        GeneratedReport.generated_at.desc()
    ).offset(offset).limit(page_size).all()
    
    # Build report data with session info and entry counts
    reports_data = []
    for report in reports_query:
        session = db.query(ReportSession).filter(
            ReportSession.id == report.session_id
        ).first()
        
        if session:
            entry_count = db.query(DataEntry).filter(
                DataEntry.session_id == session.id
            ).count()
            
            reports_data.append({
                'id': report.id,
                'session_name': session.name,
                'session_date': session.start_at or session.created_at,
                'duration_hours': session.duration_hours,
                'generated_at': report.generated_at,
                'entries_count': entry_count,
                'file_path': report.file_path,
                'report_type': report.report_type
            })
    
    # Calculate pagination
    total_pages = (total_reports + page_size - 1) // page_size
    
    return request.app.state.templates.TemplateResponse(
        "archive.html",
        {
            "request": request,
            "user": user,
            "reports": reports_data,
            "current_page": page,
            "total_pages": total_pages,
            "total_reports": total_reports
        }
    )


@router.get("/download/{report_id}")
async def download_report(
    report_id: int,
    request: Request,
    db: Session = Depends(get_db)
):
    """Download a generated report file"""
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    
    # Get report from database
    report = db.query(GeneratedReport).filter(
        GeneratedReport.id == report_id
    ).first()
    
    if not report:
        return request.app.state.templates.TemplateResponse(
            "error.html",
            {"request": request, "user": user, "error": "التقرير غير موجود"}
        )
    
    # Convert relative path to absolute path if needed
    file_path = report.file_path
    if not os.path.isabs(file_path):
        # If it's a relative path, make it absolute from app root
        app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        file_path = os.path.join(app_root, file_path)
    
    # Check if file exists
    if not os.path.exists(file_path):
        return request.app.state.templates.TemplateResponse(
            "error.html",
            {"request": request, "user": user, "error": "الملف غير متوفر"}
        )
    
    # Return file for download
    filename = os.path.basename(file_path)
    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@router.get("/custom-report")
async def custom_report_page(request: Request, db: Session = Depends(get_db)):
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    return request.app.state.templates.TemplateResponse(
        "custom_report.html",
        {"request": request, "user": user, "error": None},
    )


@router.post("/custom-report")
async def generate_custom_report(
    request: Request,
    start_date: str = Form(...),
    start_time: str = Form(...),
    end_date: str = Form(...),
    end_time: str = Form(...),
    db: Session = Depends(get_db),
):
    user = get_current_user(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    try:
        start_dt = datetime.strptime(f"{start_date} {start_time}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{end_date} {end_time}", "%Y-%m-%d %H:%M")
    except ValueError:
        return request.app.state.templates.TemplateResponse(
            "custom_report.html",
            {"request": request, "user": user, "error": "صيغة التاريخ أو الوقت غير صحيحة"},
        )

    if start_dt >= end_dt:
        return request.app.state.templates.TemplateResponse(
            "custom_report.html",
            {"request": request, "user": user, "error": "وقت البداية يجب أن يكون قبل وقت النهاية"},
        )

    entries = (
        db.query(DataEntry)
        .filter(DataEntry.created_at >= start_dt, DataEntry.created_at <= end_dt)
        .order_by(DataEntry.created_at.asc())
        .all()
    )

    breaking_news_count = (
        db.query(BreakingNews)
        .filter(BreakingNews.created_at >= start_dt, BreakingNews.created_at <= end_dt)
        .count()
    )

    if not entries and breaking_news_count == 0:
        return request.app.state.templates.TemplateResponse(
            "custom_report.html",
            {"request": request, "user": user, "error": "لا توجد بيانات ضمن الفترة الزمنية المحددة"},
        )

    file_path = generate_custom_docx_report(
        start_dt=start_dt,
        end_dt=end_dt,
        entries=entries,
        breaking_news_count=breaking_news_count,
    )

    filename = os.path.basename(file_path)
    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
