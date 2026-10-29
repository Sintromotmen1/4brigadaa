"""Authenticated website API that reuses the running Telegram bot's files and functions."""
from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, Header, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

import bot
from site_shared import consume_site_login_code, create_session_token, verify_session_token

API_TITLE = "Военный Конспект API"
SITE_AUTH_SECRET = os.getenv("SITE_AUTH_SECRET", "").strip()

app = FastAPI(title=API_TITLE, docs_url=None, redoc_url=None, openapi_url=None)
_origins = [x.strip() for x in os.getenv("SITE_ALLOWED_ORIGINS", "https://city-4brigada.pages.dev").split(",") if x.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type"],
    expose_headers=["Content-Disposition", "X-Quality-Warnings"],
)

# Website-based creation intentionally supports general, non-operational educational disciplines only.
SAFE_WEB_DISCIPLINES = {
    "строевая подготовка",
    "военно-медицинская подготовка",
    "физическая подготовка",
    "общевоинские уставы",
    "государственно-правовая подготовка",
    "психологическая подготовка",
}


class LoginRequest(BaseModel):
    telegram_id: int = Field(gt=0)
    code: str = Field(min_length=8, max_length=8, pattern=r"^\d{8}$")


class ReportRequest(BaseModel):
    kind: str = Field(min_length=2, max_length=100)
    message: str = Field(min_length=5, max_length=3000)


class AccessRequest(BaseModel):
    telegram_id: int = Field(gt=0)


class CreateDocumentRequest(BaseModel):
    discipline: str = Field(min_length=2, max_length=100)
    topic_number: str = Field(min_length=1, max_length=5, pattern=r"^\d{1,4}$")
    lesson_number: str = Field(min_length=1, max_length=5, pattern=r"^\d{1,4}$")
    topic: str = Field(min_length=3, max_length=300)
    lesson_title: str = Field(min_length=3, max_length=300)
    lesson_type: str = Field(default="Практическое занятие", max_length=100)
    audience: str = Field(min_length=2, max_length=180)
    location: str = Field(min_length=2, max_length=180)
    duration: str = Field(min_length=1, max_length=40)
    goals: str = Field(default="", max_length=1000)
    manuals: str = Field(default="", max_length=500)
    materials: str = Field(default="", max_length=1000)
    rank: str = Field(min_length=2, max_length=100)
    full_name: str = Field(min_length=3, max_length=160)
    approval_position: str = Field(default="Командир артиллерийского дивизиона", max_length=180)


def _require_auth(authorization: Optional[str] = Header(default=None)) -> dict:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Нужно войти через Telegram.")
    token = authorization[7:].strip()
    payload = verify_session_token(token, SITE_AUTH_SECRET)
    if not payload:
        raise HTTPException(status_code=401, detail="Сессия истекла. Получите новый код через /sitecode.")
    user_id = int(payload["user_id"])
    if not bot._has_access(user_id):
        raise HTTPException(status_code=403, detail="Доступ к сервису отозван владельцем.")
    return {"user_id": user_id, "is_owner": bot._is_admin(user_id)}


def _require_owner(user: dict = Depends(_require_auth)) -> dict:
    if not user["is_owner"]:
        raise HTTPException(status_code=403, detail="Раздел доступен только владельцу.")
    return user


def _safe_error_message(exc: Exception) -> str:
    text = bot._safe_error_text(exc)
    return text[:500]


@app.get("/api/health")
def health():
    return {"ok": True, "service": API_TITLE, "auth_configured": len(SITE_AUTH_SECRET) >= 32}


@app.post("/api/login")
def login(payload: LoginRequest):
    user_id = int(payload.telegram_id)
    if not bot._has_access(user_id):
        raise HTTPException(status_code=403, detail="Доступ ещё не выдан. Сначала получите разрешение владельца.")
    if not SITE_AUTH_SECRET or len(SITE_AUTH_SECRET) < 32:
        raise HTTPException(status_code=503, detail="На сервере ещё не настроен SITE_AUTH_SECRET.")
    if not consume_site_login_code(user_id, payload.code):
        raise HTTPException(status_code=401, detail="Код неверный, истёк или уже использован. Запросите новый: /sitecode")
    token, ttl = create_session_token(user_id, SITE_AUTH_SECRET)
    return {"access_token": token, "token_type": "Bearer", "expires_in": ttl,
            "user": {"telegram_id": user_id, "role": "owner" if bot._is_admin(user_id) else "user"}}


@app.get("/api/me")
def me(user: dict = Depends(_require_auth)):
    return {"telegram_id": user["user_id"], "role": "owner" if user["is_owner"] else "user",
            "is_owner": bool(user["is_owner"])}


@app.get("/api/history")
def history(user: dict = Depends(_require_auth)):
    items = bot._history_entries_for(user["user_id"])
    return {"items": [{
        "id": item.get("id", ""),
        "created_at": item.get("created_at", ""),
        "title": item.get("title", "Конспект"),
        "discipline": item.get("discipline", ""),
        "topic": item.get("topic", ""),
        "filename": item.get("filename", "Конспект.docx"),
        "download_url": f"/api/history/{item.get('id', '')}/download",
    } for item in items]}


@app.get("/api/history/{entry_id}/download")
def history_download(entry_id: str, user: dict = Depends(_require_auth)):
    item, path = bot._history_item_by_id(user["user_id"], entry_id)
    if not item or not path:
        raise HTTPException(status_code=404, detail="Документ не найден в вашей истории.")
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        filename=item.get("filename") or path.name)


@app.get("/api/books")
def books_list(user: dict = Depends(_require_auth)):
    try:
        items = []
        seen = set()
        for path in bot.list_safe_book_paths():
            name = bot.safe_book_display_name(path)
            if name in seen:
                continue
            seen.add(name)
            categories = sorted(bot._infer_safe_book_categories(path))
            items.append({"name": name, "extension": path.suffix.lower(), "categories": categories})
        return {"items": items, "count": len(items)}
    except Exception as exc:
        bot._record_error(exc, user_id=user["user_id"], event="website_books_list")
        raise HTTPException(status_code=500, detail="Не удалось прочитать каталог учебников.") from exc


@app.get("/api/books/search")
def book_search(q: str = Query(min_length=3, max_length=160), user: dict = Depends(_require_auth)):
    if bot._restricted_book_query(q):
        raise HTTPException(status_code=400, detail="Поиск доступен для общих учебных тем без прикладных инструкций по оружию или тактике.")
    try:
        matches = bot._search_safe_book_chunks(q, max_results=8)
        return {"query": q, "items": [{"source": label, "snippet": snippet} for label, snippet in matches]}
    except Exception as exc:
        bot._record_error(exc, user_id=user["user_id"], event="website_book_search")
        raise HTTPException(status_code=500, detail="Не удалось выполнить поиск по учебникам.") from exc


@app.post("/api/reports")
def submit_report(payload: ReportRequest, user: dict = Depends(_require_auth)):
    message = payload.message.strip()
    lowered = message.casefold()
    if any(marker in lowered for marker in ("telegram_token=", "gemini_api_key=", "api_key=", "password=")):
        raise HTTPException(status_code=400, detail="Не отправляйте токены, пароли или секретные данные.")
    try:
        item = bot._save_user_report(user["user_id"], "website-user", f"[{payload.kind.strip()}] {message}")
        return {"ok": True, "time": item.get("time")}
    except Exception as exc:
        bot._record_error(exc, user_id=user["user_id"], event="website_report")
        raise HTTPException(status_code=500, detail="Не удалось записать сообщение. Попробуйте позже.") from exc


@app.get("/api/owner/overview")
def owner_overview(user: dict = Depends(_require_owner)):
    # BOT_STATS is updated by the Telegram handlers in this same process.
    return {
        "authorized_users": len(bot._load_authorized_users()),
        "created_docs": int(bot.BOT_STATS.get("created_docs", 0)),
        "errors": int(bot.BOT_STATS.get("errors", 0)),
        "last_created_at": bot.BOT_STATS.get("last_created_at"),
        "last_created_by": bot.BOT_STATS.get("last_created_by"),
    }


@app.get("/api/owner/users")
def owner_users(user: dict = Depends(_require_owner)):
    return {"items": [{"telegram_id": uid, "role": "user", "access": "granted"}
                      for uid in sorted(bot._load_authorized_users())]}


@app.post("/api/owner/users/grant")
def owner_grant(payload: AccessRequest, user: dict = Depends(_require_owner)):
    target = int(payload.telegram_id)
    if target in bot.OWNER_IDS:
        raise HTTPException(status_code=400, detail="Владелец уже имеет полный доступ.")
    bot.AUTHORIZED_USER_IDS.add(target)
    try:
        bot._save_authorized_users()
    except Exception as exc:
        bot.AUTHORIZED_USER_IDS.discard(target)
        bot._record_error(exc, user_id=user["user_id"], event="website_owner_grant")
        raise HTTPException(status_code=500, detail="Не удалось сохранить список пользователей.") from exc
    return {"ok": True, "telegram_id": target, "role": "user", "access": "granted"}


@app.post("/api/owner/users/revoke")
def owner_revoke(payload: AccessRequest, user: dict = Depends(_require_owner)):
    target = int(payload.telegram_id)
    if target in bot.OWNER_IDS:
        raise HTTPException(status_code=400, detail="Нельзя отозвать доступ у владельца.")
    was_present = target in bot.AUTHORIZED_USER_IDS
    bot.AUTHORIZED_USER_IDS.discard(target)
    try:
        bot._save_authorized_users()
    except Exception as exc:
        if was_present:
            bot.AUTHORIZED_USER_IDS.add(target)
        bot._record_error(exc, user_id=user["user_id"], event="website_owner_revoke")
        raise HTTPException(status_code=500, detail="Не удалось сохранить список пользователей.") from exc
    return {"ok": True, "telegram_id": target, "role": "user", "access": "revoked"}


@app.get("/api/owner/errors")
def owner_errors(user: dict = Depends(_require_owner)):
    return {"items": bot._read_recent_errors(20)}


@app.get("/api/owner/reports")
def owner_reports(user: dict = Depends(_require_owner)):
    return {"items": bot._read_user_reports(20)}


def _generate_doc(payload: CreateDocumentRequest, user_id: int) -> tuple[str, list[str], str]:
    discipline = payload.discipline.strip()
    if bot.normalize_text(discipline) not in SAFE_WEB_DISCIPLINES:
        raise HTTPException(status_code=400, detail="Через сайт пока доступны только общие учебные дисциплины.")
    combined_topic = f"{payload.topic} {payload.lesson_title}"
    if bot._restricted_book_query(combined_topic):
        raise HTTPException(status_code=400, detail="Эта тема не поддерживается веб-генератором. Выберите общую учебную тему без прикладных инструкций.")

    # Use an exact program match when it is unique. Manual title fields are retained if the program parser cannot match.
    match = None
    try:
        match = bot.find_program_lesson(discipline, payload.topic_number, payload.lesson_number)
    except Exception as exc:
        bot._record_error(exc, user_id=user_id, event="website_program_lookup")
    topic_title = payload.topic.strip()
    lesson_title = payload.lesson_title.strip()
    if isinstance(match, dict) and match.get("multiple"):
        candidates = match.get("multiple", [])
        exact = next((x for x in candidates if bot.normalize_text(x.get("lesson_title", "")) == bot.normalize_text(lesson_title)), None)
        if exact:
            topic_title = exact.get("topic", topic_title)
            lesson_title = exact.get("lesson_title", lesson_title)
    elif isinstance(match, dict) and match:
        topic_title = match.get("topic", topic_title) or topic_title
        lesson_title = match.get("lesson_title", lesson_title) or lesson_title

    output_path = None
    try:
        data = bot.generate_lesson_data(
            discipline, topic_title, lesson_title, payload.lesson_type.strip(), payload.audience.strip(),
            payload.duration.strip(), payload.goals.strip(), payload.manuals.strip(), payload.materials.strip(),
        )
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp:
            output_path = temp.name
        bot.create_from_template(
            payload.rank.strip(), payload.full_name.strip(), discipline,
            payload.topic_number.strip(), topic_title, payload.lesson_number.strip(), lesson_title,
            payload.lesson_type.strip(), payload.audience.strip(), payload.location.strip(),
            payload.duration.strip(), payload.goals.strip(), data, output_path,
            payload.approval_position.strip() or "Командир артиллерийского дивизиона",
        )
        warnings = bot.validate_generated_docx(output_path, expected_topic=topic_title, expected_title=lesson_title)
        bot._archive_document(user_id, output_path, lesson_title, discipline, topic_title)
        bot._record_generated_doc(user_id, "конспект через сайт")
        safe_name = re.sub(r"[^A-Za-zА-Яа-яЁё0-9_-]+", "_", lesson_title)[:60].strip("_") or "Konspekt"
        return output_path, warnings, f"Конспект_{safe_name}.docx"
    except HTTPException:
        if output_path and os.path.exists(output_path):
            os.remove(output_path)
        raise
    except Exception as exc:
        if output_path and os.path.exists(output_path):
            os.remove(output_path)
        bot._record_error(exc, user_id=user_id, event="website_doc_generation")
        raise HTTPException(status_code=500, detail="Не удалось создать Word-файл. Проверьте шаблон и настройки сервера, затем попробуйте ещё раз.") from exc


@app.post("/api/create")
async def create_document(payload: CreateDocumentRequest, user: dict = Depends(_require_auth)):
    output_path, warnings, filename = await run_in_threadpool(_generate_doc, payload, user["user_id"])
    warning_header = " | ".join(warnings)[:1000] if warnings else "none"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
        headers={"X-Quality-Warnings": warning_header},
        background=BackgroundTask(lambda: Path(output_path).unlink(missing_ok=True)),
    )
