import os
from datetime import datetime

from fastapi import APIRouter, Request, Depends, Form
from fastapi.responses import RedirectResponse, FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import ReportSession, DataEntry, BreakingNews, GeneratedReport
from app.report_generator import generate_custom_docx_report

router = APIRouter(prefix="/public")


@router.get("/{session_id}")
async def public_view(session_id: int, request: Request, db: Session = Depends(get_db)):
    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return request.app.state.templates.TemplateResponse(
            "public_view.html",
            {"request": request, "not_found": True},
        )

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
    reports = (
        db.query(GeneratedReport)
        .filter(GeneratedReport.session_id == report_session.id)
        .order_by(GeneratedReport.generated_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "public_view.html",
        {
            "request": request,
            "session": report_session,
            "entries": entries,
            "breaking_news_count": breaking_news_count,
            "reports": reports,
            "not_found": False,
        },
    )


@router.get("/{session_id}/download/{report_id}")
async def public_download_report(session_id: int, report_id: int, request: Request, db: Session = Depends(get_db)):
    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return RedirectResponse(url="/", status_code=303)

    report = (
        db.query(GeneratedReport)
        .filter(GeneratedReport.id == report_id, GeneratedReport.session_id == report_session.id)
        .first()
    )
    if not report or not os.path.exists(report.file_path):
        return RedirectResponse(url=f"/public/{session_id}", status_code=303)

    return FileResponse(
        report.file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(report.file_path),
    )


@router.get("/custom-report/{public_token}")
async def public_custom_report_page(public_token: str, request: Request):
    token = os.getenv("PUBLIC_CUSTOM_REPORT_TOKEN")
    if not token or public_token != token:
        return request.app.state.templates.TemplateResponse(
            "error.html",
            {"request": request, "error": "رابط غير صحيح أو منتهي الصلاحية"},
            status_code=404,
        )
    
    return request.app.state.templates.TemplateResponse(
        "public_custom_report.html",
        {"request": request},
    )


@router.post("/custom-report/{public_token}")
async def generate_public_custom_report(
    public_token: str,
    request: Request,
    start_at: str = Form(...),
    end_at: str = Form(...),
    db: Session = Depends(get_db),
):
    token = os.getenv("PUBLIC_CUSTOM_REPORT_TOKEN")
    if not token or public_token != token:
        return request.app.state.templates.TemplateResponse(
            "error.html",
            {"request": request, "error": "رابط غير صحيح أو منتهي الصلاحية"},
            status_code=404,
        )

    try:
        start_dt = datetime.fromisoformat(start_at)
        end_dt = datetime.fromisoformat(end_at)
    except ValueError:
        return request.app.state.templates.TemplateResponse(
            "public_custom_report.html",
            {"request": request, "error": "صيغة التاريخ أو الوقت غير صحيحة"},
        )

    if start_dt >= end_dt:
        return request.app.state.templates.TemplateResponse(
            "public_custom_report.html",
            {"request": request, "error": "وقت البداية يجب أن يكون قبل وقت النهاية"},
        )

    entries = (
        db.query(DataEntry)
        .filter(DataEntry.created_at >= start_dt, DataEntry.created_at <= end_dt)
        .order_by(DataEntry.monitoring_time)
        .all()
    )

    breaking_news_count = (
        db.query(BreakingNews)
        .filter(BreakingNews.created_at >= start_dt, BreakingNews.created_at <= end_dt)
        .count()
    )

    if not entries and breaking_news_count == 0:
        return request.app.state.templates.TemplateResponse(
            "public_custom_report.html",
            {"request": request, "error": "لا توجد بيانات ضمن الفترة الزمنية المحددة"},
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
