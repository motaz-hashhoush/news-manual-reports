from fastapi import APIRouter, Request, Depends
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import ReportSession
from app.deps import get_current_user

router = APIRouter()


@router.get("/dashboard")
async def dashboard(request: Request, db: Session = Depends(get_db)):
    user = get_current_user(request, db)
    if not user:
        from fastapi.responses import RedirectResponse
        return RedirectResponse(url="/login", status_code=303)

    sessions = (
        db.query(ReportSession)
        .order_by(ReportSession.created_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "dashboard.html",
        {"request": request, "user": user, "sessions": sessions},
    )
