from fastapi import APIRouter, Request, Depends, Form
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import User
from app.auth import hash_password, verify_password, create_session_token

router = APIRouter()


@router.get("/login")
async def login_page(request: Request):
    return request.app.state.templates.TemplateResponse(
        "login.html", {"request": request, "error": None}
    )


@router.post("/login")
async def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db),
):
    user = db.query(User).filter(User.username == username).first()
    if not user or not verify_password(password, user.password_hash):
        return request.app.state.templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "اسم المستخدم أو كلمة المرور غير صحيحة"},
        )
    token = create_session_token(user.id, user.role)
    response = RedirectResponse(url="/dashboard", status_code=303)
    response.set_cookie("session_token", token, httponly=True, max_age=86400)
    return response


@router.get("/register")
async def register_page(request: Request):
    return request.app.state.templates.TemplateResponse(
        "register.html", {"request": request, "error": None}
    )


@router.post("/register")
async def register(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    password_confirm: str = Form(...),
    db: Session = Depends(get_db),
):
    if password != password_confirm:
        return request.app.state.templates.TemplateResponse(
            "register.html",
            {"request": request, "error": "كلمتا المرور غير متطابقتين"},
        )

    existing = db.query(User).filter(User.username == username).first()
    if existing:
        return request.app.state.templates.TemplateResponse(
            "register.html",
            {"request": request, "error": "اسم المستخدم مسجل مسبقاً"},
        )

    user = User(
        username=username,
        password_hash=hash_password(password),
        role="user",
    )
    db.add(user)
    db.commit()

    token = create_session_token(user.id, user.role)
    response = RedirectResponse(url="/dashboard", status_code=303)
    response.set_cookie("session_token", token, httponly=True, max_age=86400)
    return response


@router.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie("session_token")
    return response
