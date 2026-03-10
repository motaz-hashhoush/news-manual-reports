import os

from fastapi import APIRouter, Request, Depends
from fastapi.responses import RedirectResponse, FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import ReportSession, DataEntry, BreakingNews, GeneratedReport

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
