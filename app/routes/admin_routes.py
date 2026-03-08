import io
import os
import uuid
from datetime import datetime, timedelta

import pandas as pd
from fastapi import APIRouter, Request, Depends, Form, UploadFile, File
from fastapi.responses import RedirectResponse, FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import User, ReportSession, DataEntry, GeneratedReport, BreakingNews
from app.deps import get_current_user
from app.auth import hash_password
from app.config import UPLOAD_DIR, DISTRIBUTION_VALUES, TYPE_VALUES
from app.timezone import now_palestine

COLUMN_MAP = {
    "فترة الرصد": "monitoring_time",
    "العنوان / NOW": "title",
    "العنوان": "title",
    "البرنامج": "program",
    "النوع": "entry_type",
    "النوع (أخبار - اقتصاد)": "entry_type",
    "التوزيع": "distribution",
    "اسم الضيف/المراسل": "guest_reporter_name",
    "اسم الضيف/المراسل - حال وجوده": "guest_reporter_name",
    "رابط النشر": "publish_link",
    "الأهمية": "importance",
}

router = APIRouter(prefix="/admin")


def _require_admin(request, db):
    user = get_current_user(request, db)
    if not user or user.role != "admin":
        return None
    return user


@router.get("/create-session")
async def create_session_page(request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    return request.app.state.templates.TemplateResponse(
        "create_session.html", {"request": request, "user": user, "error": None}
    )


@router.post("/create-session")
async def create_session(
    request: Request,
    name: str = Form(...),
    description: str = Form(""),
    duration_mode: str = Form("24"),
    custom_start: str = Form(""),
    custom_end: str = Form(""),
    db: Session = Depends(get_db),
):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    if duration_mode == "custom":
        try:
            start_at = datetime.strptime(custom_start, "%Y-%m-%dT%H:%M")
            deadline_at = datetime.strptime(custom_end, "%Y-%m-%dT%H:%M")
        except (ValueError, TypeError):
            return request.app.state.templates.TemplateResponse(
                "create_session.html",
                {"request": request, "user": user, "error": "يرجى تحديد تاريخ البداية والنهاية بشكل صحيح."},
            )
        if deadline_at <= start_at:
            return request.app.state.templates.TemplateResponse(
                "create_session.html",
                {"request": request, "user": user, "error": "يجب أن يكون وقت النهاية بعد وقت البداية."},
            )
        diff = deadline_at - start_at
        duration_hours = max(1, int(diff.total_seconds() / 3600))
    else:
        duration_hours = int(duration_mode) if duration_mode in ("12", "24") else 24
        now = now_palestine()
        start_at = now.replace(hour=18, minute=0, second=0, microsecond=0)
        deadline_at = start_at + timedelta(hours=duration_hours)

    session = ReportSession(
        name=name,
        description=description,
        created_by=user.id,
        duration_hours=duration_hours,
        start_at=start_at,
        deadline_at=deadline_at,
    )
    db.add(session)
    db.commit()
    return RedirectResponse(url="/dashboard", status_code=303)


@router.get("/entry/{session_id}")
async def data_entry_page(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session or report_session.status != "active":
        return RedirectResponse(url="/dashboard", status_code=303)

    return request.app.state.templates.TemplateResponse(
        "data_entry.html",
        {
            "request": request,
            "user": user,
            "session": report_session,
            "distribution_values": DISTRIBUTION_VALUES,
            "type_values": TYPE_VALUES,
            "error": None,
            "success": None,
        },
    )


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
    screenshot: UploadFile = File(None),
    db: Session = Depends(get_db),
):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session or report_session.status != "active":
        return RedirectResponse(url="/dashboard", status_code=303)

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
        screenshot_path=screenshot_path,
        screenshot_data=screenshot_data,
    )
    db.add(entry)
    db.commit()

    return request.app.state.templates.TemplateResponse(
        "data_entry.html",
        {
            "request": request,
            "user": user,
            "session": report_session,
            "distribution_values": DISTRIBUTION_VALUES,
            "type_values": TYPE_VALUES,
            "error": None,
            "success": "تم إضافة البيانات بنجاح",
        },
    )


@router.post("/import/{session_id}")
async def import_entries(
    session_id: int,
    request: Request,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session or report_session.status != "active":
        return RedirectResponse(url="/dashboard", status_code=303)

    def _render(error=None, success=None):
        return request.app.state.templates.TemplateResponse(
            "data_entry.html",
            {
                "request": request,
                "user": user,
                "session": report_session,
                "distribution_values": DISTRIBUTION_VALUES,
                "type_values": TYPE_VALUES,
                "error": error,
                "success": success,
            },
        )

    if not file or not file.filename:
        return _render(error="يرجى اختيار ملف CSV أو Excel")

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".csv", ".xlsx", ".xls"):
        return _render(error="صيغة الملف غير مدعومة — يرجى رفع ملف CSV أو Excel (.xlsx)")

    try:
        content = await file.read()
        if ext == ".csv":
            df = pd.read_csv(io.BytesIO(content), encoding="utf-8-sig")
        else:
            df = pd.read_excel(io.BytesIO(content))
    except Exception as e:
        return _render(error=f"خطأ في قراءة الملف: {e}")

    df.columns = df.columns.str.strip()
    rename = {}
    for col in df.columns:
        for ar_key, en_key in COLUMN_MAP.items():
            if ar_key in col:
                rename[col] = en_key
                break
    df.rename(columns=rename, inplace=True)

    required = {"monitoring_time", "title"}
    missing = required - set(df.columns)
    if missing:
        nice = {"monitoring_time": "فترة الرصد", "title": "العنوان"}
        labels = ", ".join(nice.get(m, m) for m in missing)
        return _render(error=f"الأعمدة المطلوبة غير موجودة في الملف: {labels}")

    df = df.fillna("")
    inserted = 0
    for _, row in df.iterrows():
        monitoring_time = str(row.get("monitoring_time", "")).strip()
        title = str(row.get("title", "")).strip()
        if not monitoring_time or not title:
            continue

        entry = DataEntry(
            session_id=session_id,
            user_id=user.id,
            monitoring_time=monitoring_time,
            title=title,
            program=str(row.get("program", "")).strip(),
            entry_type=str(row.get("entry_type", "")).strip(),
            distribution=str(row.get("distribution", "")).strip(),
            guest_reporter_name=str(row.get("guest_reporter_name", "")).strip(),
            publish_link=str(row.get("publish_link", "")).strip(),
            importance=str(row.get("importance", "")).strip(),
        )
        db.add(entry)
        inserted += 1

    db.commit()
    return _render(success=f"تم استيراد {inserted} سجل بنجاح من الملف")


@router.get("/entries/{session_id}")
async def view_entries(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return RedirectResponse(url="/dashboard", status_code=303)

    entries = (
        db.query(DataEntry)
        .filter(DataEntry.session_id == session_id)
        .order_by(DataEntry.monitoring_time)
        .all()
    )

    breaking_news_count = (
        db.query(BreakingNews)
        .filter(BreakingNews.session_id == session_id)
        .count()
    )

    return request.app.state.templates.TemplateResponse(
        "view_entries.html",
        {
            "request": request,
            "user": user,
            "session": report_session,
            "entries": entries,
            "breaking_news_count": breaking_news_count,
        },
    )


@router.post("/generate-report/{session_id}")
async def generate_report(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    from app.report_generator import generate_docx_report

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return RedirectResponse(url="/dashboard", status_code=303)

    entries = (
        db.query(DataEntry)
        .filter(DataEntry.session_id == session_id)
        .order_by(DataEntry.monitoring_time)
        .all()
    )

    bn_count = db.query(BreakingNews).filter(BreakingNews.session_id == session_id).count()

    file_path = generate_docx_report(report_session, entries, breaking_news_count=bn_count)

    report_record = GeneratedReport(
        session_id=session_id,
        file_path=file_path,
        report_type="on_demand",
    )
    db.add(report_record)
    db.commit()

    return RedirectResponse(
        url=f"/admin/reports/{session_id}", status_code=303
    )


@router.get("/reports/{session_id}")
async def list_reports(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return RedirectResponse(url="/dashboard", status_code=303)

    reports = (
        db.query(GeneratedReport)
        .filter(GeneratedReport.session_id == session_id)
        .order_by(GeneratedReport.generated_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "reports.html",
        {"request": request, "user": user, "session": report_session, "reports": reports},
    )


@router.get("/download-report/{report_id}")
async def download_report(report_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report = db.query(GeneratedReport).filter(GeneratedReport.id == report_id).first()
    if not report or not os.path.exists(report.file_path):
        return RedirectResponse(url="/dashboard", status_code=303)

    return FileResponse(
        report.file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=os.path.basename(report.file_path),
    )


@router.post("/delete-entry/{entry_id}")
async def delete_entry(entry_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    entry = db.query(DataEntry).filter(DataEntry.id == entry_id).first()
    if not entry:
        return RedirectResponse(url="/dashboard", status_code=303)

    session_id = entry.session_id

    if entry.screenshot_path:
        full_path = os.path.join(UPLOAD_DIR, os.path.basename(entry.screenshot_path))
        if os.path.exists(full_path):
            os.remove(full_path)

    db.delete(entry)
    db.commit()

    return RedirectResponse(url=f"/admin/entries/{session_id}", status_code=303)


# ── Breaking News ─────────────────────────────────────────────────────────────


@router.get("/breaking-news/{session_id}")
async def breaking_news_page(session_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session:
        return RedirectResponse(url="/dashboard", status_code=303)

    items = (
        db.query(BreakingNews)
        .filter(BreakingNews.session_id == session_id)
        .order_by(BreakingNews.created_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "breaking_news.html",
        {
            "request": request,
            "user": user,
            "session": report_session,
            "items": items,
            "error": None,
            "success": None,
        },
    )


@router.post("/breaking-news/{session_id}")
async def submit_breaking_news(
    session_id: int,
    request: Request,
    description: str = Form(""),
    screenshot: UploadFile = File(None),
    db: Session = Depends(get_db),
):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    report_session = db.query(ReportSession).filter(ReportSession.id == session_id).first()
    if not report_session or report_session.status != "active":
        return RedirectResponse(url="/dashboard", status_code=303)

    screenshot_path = None
    screenshot_data = None
    if screenshot and screenshot.filename:
        ext = os.path.splitext(screenshot.filename)[1]
        filename = f"bn_{uuid.uuid4().hex}{ext}"
        save_path = os.path.join(UPLOAD_DIR, filename)
        content = await screenshot.read()
        with open(save_path, "wb") as f:
            f.write(content)
        screenshot_path = f"/static/uploads/{filename}"
        screenshot_data = content

    item = BreakingNews(
        session_id=session_id,
        user_id=user.id,
        description=description.strip() or None,
        screenshot_path=screenshot_path,
        screenshot_data=screenshot_data,
    )
    db.add(item)
    db.commit()

    items = (
        db.query(BreakingNews)
        .filter(BreakingNews.session_id == session_id)
        .order_by(BreakingNews.created_at.desc())
        .all()
    )

    return request.app.state.templates.TemplateResponse(
        "breaking_news.html",
        {
            "request": request,
            "user": user,
            "session": report_session,
            "items": items,
            "error": None,
            "success": "تمت إضافة الخبر العاجل بنجاح",
        },
    )


@router.post("/breaking-news/delete/{item_id}")
async def delete_breaking_news(item_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    item = db.query(BreakingNews).filter(BreakingNews.id == item_id).first()
    if not item:
        return RedirectResponse(url="/dashboard", status_code=303)

    session_id = item.session_id

    if item.screenshot_path:
        full_path = os.path.join(UPLOAD_DIR, os.path.basename(item.screenshot_path))
        if os.path.exists(full_path):
            os.remove(full_path)

    db.delete(item)
    db.commit()

    return RedirectResponse(url=f"/admin/breaking-news/{session_id}", status_code=303)


# ── User Management ──────────────────────────────────────────────────────────


@router.get("/users")
async def list_users(request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    users = db.query(User).order_by(User.created_at.desc()).all()

    return request.app.state.templates.TemplateResponse(
        "manage_users.html",
        {"request": request, "user": user, "users": users, "error": None, "success": None},
    )


@router.post("/users/create")
async def create_user(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    role: str = Form("user"),
    db: Session = Depends(get_db),
):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    users = db.query(User).order_by(User.created_at.desc()).all()

    if not username.strip() or not password.strip():
        return request.app.state.templates.TemplateResponse(
            "manage_users.html",
            {"request": request, "user": user, "users": users,
             "error": "اسم المستخدم وكلمة المرور مطلوبان", "success": None},
        )

    existing = db.query(User).filter(User.username == username).first()
    if existing:
        return request.app.state.templates.TemplateResponse(
            "manage_users.html",
            {"request": request, "user": user, "users": users,
             "error": "اسم المستخدم مسجل مسبقاً", "success": None},
        )

    if role not in ("admin", "user"):
        role = "user"

    new_user = User(
        username=username,
        password_hash=hash_password(password),
        role=role,
    )
    db.add(new_user)
    db.commit()

    users = db.query(User).order_by(User.created_at.desc()).all()

    return request.app.state.templates.TemplateResponse(
        "manage_users.html",
        {"request": request, "user": user, "users": users,
         "error": None, "success": f"تم إنشاء المستخدم «{username}» بنجاح"},
    )


@router.post("/users/delete/{user_id}")
async def delete_user(user_id: int, request: Request, db: Session = Depends(get_db)):
    user = _require_admin(request, db)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    target = db.query(User).filter(User.id == user_id).first()
    if not target:
        return RedirectResponse(url="/admin/users", status_code=303)

    if target.id == user.id:
        return RedirectResponse(url="/admin/users", status_code=303)

    db.delete(target)
    db.commit()

    return RedirectResponse(url="/admin/users", status_code=303)
