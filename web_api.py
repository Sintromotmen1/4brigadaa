"""Authenticated website API that reuses the running Telegram bot's files and functions."""
from __future__ import annotations

import os
import re
import tempfile
import json
import time
import hmac
import ssl
import smtplib
import secrets
import hashlib
import threading
from email.message import EmailMessage
from pathlib import Path
from datetime import datetime, timezone
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

SITE_OWNER_EMAIL = os.getenv("SITE_OWNER_EMAIL", "").strip().lower()
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "").strip()
SMTP_PASSWORD = re.sub(r"\s+", "", os.getenv("SMTP_PASSWORD", ""))
SMTP_FROM = os.getenv("SMTP_FROM", SMTP_USERNAME).strip()

EMAIL_CODE_TTL_SECONDS = 10 * 60
EMAIL_CODE_RESEND_SECONDS = 60
EMAIL_SESSION_TTL_SECONDS = 6 * 60 * 60

ACCOUNT_FILE = Path(__file__).resolve().parent / "site_accounts.json"
EMAIL_CODES_FILE = Path(__file__).resolve().parent / "site_email_codes.json"

_ACCOUNT_LOCK = threading.RLock()
_EMAIL_CODE_LOCK = threading.RLock()
_PASSWORD_ITERATIONS = 310_000

_GMAIL_RE = re.compile(
    r"^[a-z0-9][a-z0-9._%+-]{0,63}@(?:gmail\.com|googlemail\.com)$",
    re.I
)


app = FastAPI(
    title=API_TITLE,
    docs_url=None,
    redoc_url=None,
    openapi_url=None
)

_origins = [
    x.strip()
    for x in os.getenv(
        "SITE_ALLOWED_ORIGINS",
        "https://city-4brigada.pages.dev"
    ).split(",")
    if x.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type"],
    expose_headers=["Content-Disposition", "X-Quality-Warnings"],
)


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
    code: str = Field(
        min_length=8,
        max_length=8,
        pattern=r"^\d{8}$"
    )


class EmailCredentials(BaseModel):
    email: str = Field(min_length=6, max_length=254)
    password: str = Field(min_length=8, max_length=128)
    display_name: str = Field(default="", max_length=80)


class EmailVerifyRequest(BaseModel):
    email: str = Field(min_length=6, max_length=254)
    code: str = Field(
        min_length=6,
        max_length=6,
        pattern=r"^\d{6}$"
    )


class EmailOnlyRequest(BaseModel):
    email: str = Field(min_length=6, max_length=254)


class ReportRequest(BaseModel):
    kind: str = Field(min_length=2, max_length=100)
    message: str = Field(min_length=5, max_length=3000)


class AccessRequest(BaseModel):
    telegram_id: int = Field(gt=0)


class CreateDocumentRequest(BaseModel):
    discipline: str = Field(min_length=2, max_length=100)
    topic_number: str = Field(
        min_length=1,
        max_length=5,
        pattern=r"^\d{1,4}$"
    )
    lesson_number: str = Field(
        min_length=1,
        max_length=5,
        pattern=r"^\d{1,4}$"
    )
    topic: str = Field(min_length=3, max_length=300)
    lesson_title: str = Field(min_length=3, max_length=300)
    lesson_type: str = Field(
        default="Практическое занятие",
        max_length=100
    )
    audience: str = Field(min_length=2, max_length=180)
    location: str = Field(min_length=2, max_length=180)
    duration: str = Field(min_length=1, max_length=40)
    goals: str = Field(default="", max_length=1000)
    manuals: str = Field(default="", max_length=500)
    materials: str = Field(default="", max_length=1000)
    rank: str = Field(min_length=2, max_length=100)
    full_name: str = Field(min_length=3, max_length=160)
    approval_position: str = Field(
        default="Командир артиллерийского дивизиона",
        max_length=180
    )


def datetime_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _read_json_dict(path: Path) -> dict:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, json.JSONDecodeError, TypeError):
        return {}


def _atomic_write_json(path: Path, value: dict) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    os.replace(temporary, path)


def _load_accounts() -> dict:
    with _ACCOUNT_LOCK:
        raw = _read_json_dict(ACCOUNT_FILE)
        return {
            str(key).lower(): item
            for key, item in raw.items()
            if isinstance(item, dict)
        }


def _save_accounts(accounts: dict) -> None:
    with _ACCOUNT_LOCK:
        _atomic_write_json(ACCOUNT_FILE, accounts)


def _normalise_gmail(value: str) -> str:
    email = str(value or "").strip().lower()

    if not _GMAIL_RE.fullmatch(email):
        raise HTTPException(
            status_code=400,
            detail="Укажите корректный адрес Gmail (@gmail.com)."
        )

    return email


def _hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)

    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        _PASSWORD_ITERATIONS
    )

    return (
        f"pbkdf2_sha256${_PASSWORD_ITERATIONS}$"
        f"{salt.hex()}${digest.hex()}"
    )


def _verify_password(password: str, encoded: str) -> bool:
    try:
        algorithm, iterations, salt_hex, digest_hex = encoded.split("$", 3)

        if algorithm != "pbkdf2_sha256":
            return False

        actual = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            bytes.fromhex(salt_hex),
            int(iterations)
        ).hex()

        return hmac.compare_digest(actual, digest_hex)

    except (ValueError, TypeError):
        return False


def _account_role(email: str, account: dict) -> str:
    return (
        "owner"
        if SITE_OWNER_EMAIL and email.lower() == SITE_OWNER_EMAIL
        else "user"
    )


def _new_account_user_id(accounts: dict) -> int:
    used = {
        int(value.get("user_id", 0) or 0)
        for value in accounts.values()
    }

    used.update(int(value) for value in bot.OWNER_IDS)
    used.update(int(value) for value in bot.AUTHORIZED_USER_IDS)

    while True:
        candidate = secrets.randbelow(8_000_000_000_000) + 1_000_000_000_000

        if candidate not in used:
            return candidate


def _email_session_token(
    account: dict,
    secret: str = SITE_AUTH_SECRET
) -> tuple[str, int]:

    if not secret or len(secret) < 32:
        raise HTTPException(
            status_code=503,
            detail="На сервере не настроен SITE_AUTH_SECRET (нужно не менее 32 символов)."
        )

    now = int(time.time())

    payload = {
        "kind": "email",
        "email": account["email"],
        "uid": int(account["user_id"]),
        "version": int(account.get("session_version", 1)),
        "iat": now,
        "exp": now + EMAIL_SESSION_TTL_SECONDS,
        "nonce": secrets.token_hex(12),
    }

    encoded = __import__("base64").urlsafe_b64encode(
        json.dumps(
            payload,
            separators=(",", ":"),
            sort_keys=True
        ).encode("utf-8")
    ).decode("ascii").rstrip("=")

    signature = hmac.new(
        secret.encode("utf-8"),
        encoded.encode("ascii"),
        hashlib.sha256
    ).digest()

    signature_text = __import__("base64").urlsafe_b64encode(
        signature
    ).decode("ascii").rstrip("=")

    return encoded + "." + signature_text, EMAIL_SESSION_TTL_SECONDS


def _verify_email_session(token: str) -> dict | None:
    if (
        not SITE_AUTH_SECRET
        or len(SITE_AUTH_SECRET) < 32
        or not token
        or "." not in token
    ):
        return None

    try:
        import base64

        encoded, supplied = token.split(".", 1)

        expected = base64.urlsafe_b64encode(
            hmac.new(
                SITE_AUTH_SECRET.encode("utf-8"),
                encoded.encode("ascii"),
                hashlib.sha256
            ).digest()
        ).decode("ascii").rstrip("=")

        if not hmac.compare_digest(expected, supplied):
            return None

        decoded = base64.urlsafe_b64decode(
            encoded + "=" * ((4 - len(encoded) % 4) % 4)
        )

        payload = json.loads(decoded.decode("utf-8"))

        if (
            payload.get("kind") != "email"
            or int(payload.get("exp", 0)) <= int(time.time())
        ):
            return None

        email = str(payload.get("email", "")).lower()
        account = _load_accounts().get(email)

        if (
            not account
            or not account.get("verified")
            or not account.get("active", True)
        ):
            return None

        if int(account.get("session_version", 1)) != int(
            payload.get("version", 0)
        ):
            return None

        return {
            "user_id": int(account["user_id"]),
            "is_owner": _account_role(email, account) == "owner",
            "email": email,
            "display_name": account.get("display_name", ""),
            "auth_type": "email",
        }

    except (
        ValueError,
        TypeError,
        KeyError,
        UnicodeDecodeError,
        json.JSONDecodeError
    ):
        return None


def _load_email_codes() -> dict:
    with _EMAIL_CODE_LOCK:
        return _read_json_dict(EMAIL_CODES_FILE)


def _save_email_codes(codes: dict) -> None:
    with _EMAIL_CODE_LOCK:
        _atomic_write_json(EMAIL_CODES_FILE, codes)


def _send_verification_email(email: str, code: str) -> None:
    if not SMTP_USERNAME or not SMTP_PASSWORD or not SMTP_FROM:
        print(
            "[SMTP] Не заданы SMTP_USERNAME / SMTP_PASSWORD / SMTP_FROM.",
            flush=True
        )

        raise HTTPException(
            status_code=503,
            detail="Отправка писем не настроена. Проверьте SMTP_USERNAME, SMTP_PASSWORD и SMTP_FROM в FadeHost."
        )

    message = EmailMessage()
    message["Subject"] = "Код подтверждения — Военный Конспект"
    message["From"] = SMTP_FROM
    message["To"] = email

    message.set_content(
        "Здравствуйте!\n\n"
        f"Код подтверждения Gmail: {code}\n\n"
        "Код действует 10 минут и может быть использован один раз. "
        "Если вы не регистрировались на сайте, просто проигнорируйте это письмо.\n"
    )

    def send_using_port(port: int) -> None:
        if port == 465:
            with smtplib.SMTP_SSL(
                SMTP_HOST,
                port,
                timeout=20,
                context=ssl.create_default_context()
            ) as server:
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
                server.send_message(message)

        else:
            with smtplib.SMTP(
                SMTP_HOST,
                port,
                timeout=20
            ) as server:
                server.ehlo()
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
                server.send_message(message)

    primary_port = SMTP_PORT
    fallback_port = 465 if primary_port != 465 else 587

    print(
        f"[SMTP] Начинаю отправку письма; "
        f"сервер={SMTP_HOST}; порт={primary_port}.",
        flush=True
    )

    try:
        send_using_port(primary_port)

        print(
            f"[SMTP] Письмо успешно отправлено; порт={primary_port}.",
            flush=True
        )
        return

    except smtplib.SMTPAuthenticationError as exc:
        bot._record_error(
            exc,
            event="website_email_verification"
        )

        print(
            f"[SMTP ERROR] Gmail отклонил авторизацию; "
            f"код={getattr(exc, 'smtp_code', 'unknown')}.",
            flush=True
        )

        raise HTTPException(
            status_code=502,
            detail="Gmail отклонил SMTP-авторизацию. Проверьте SMTP_USERNAME и пароль приложения Google."
        ) from exc

    except (
        OSError,
        smtplib.SMTPConnectError,
        smtplib.SMTPServerDisconnected
    ) as first_exc:

        bot._record_error(
            first_exc,
            event="website_email_verification"
        )

        print(
            f"[SMTP ERROR] Не удалось подключиться через порт "
            f"{primary_port}: {type(first_exc).__name__}. "
            f"Пробую порт {fallback_port}.",
            flush=True
        )

        try:
            send_using_port(fallback_port)

            print(
                f"[SMTP] Письмо успешно отправлено через "
                f"резервный порт {fallback_port}.",
                flush=True
            )
            return

        except smtplib.SMTPAuthenticationError as exc:
            bot._record_error(
                exc,
                event="website_email_verification"
            )

            print(
                f"[SMTP ERROR] Gmail отклонил авторизацию "
                f"на резервном порту; "
                f"код={getattr(exc, 'smtp_code', 'unknown')}.",
                flush=True
            )

            raise HTTPException(
                status_code=502,
                detail="Gmail отклонил SMTP-авторизацию. Проверьте SMTP_USERNAME и пароль приложения Google."
            ) from exc

        except (OSError, smtplib.SMTPException) as exc:
            bot._record_error(
                exc,
                event="website_email_verification"
            )

            print(
                f"[SMTP ERROR] Не сработали порты "
                f"{primary_port} и {fallback_port}; "
                f"ошибка={type(exc).__name__}; "
                f"код={getattr(exc, 'smtp_code', 'нет')}.",
                flush=True
            )

            raise HTTPException(
                status_code=502,
                detail="FadeHost не смог подключиться к Gmail через порты 587 и 465. Возможно, исходящие SMTP-подключения ограничены хостингом."
            ) from exc

    except smtplib.SMTPException as exc:
        bot._record_error(
            exc,
            event="website_email_verification"
        )

        print(
            f"[SMTP ERROR] Gmail отклонил письмо; "
            f"ошибка={type(exc).__name__}; "
            f"код={getattr(exc, 'smtp_code', 'нет')}.",
            flush=True
        )

        raise HTTPException(
            status_code=502,
            detail="Gmail отклонил отправку письма. Проверьте адрес отправителя SMTP_FROM и настройки Gmail."
        ) from exc


def _create_and_send_email_code(
    email: str,
    password: str,
    display_name: str,
    *,
    allow_resend: bool = False
) -> None:

    now = int(time.time())
    codes = _load_email_codes()
    previous = codes.get(email)

    if (
        previous
        and now - int(previous.get("created_at", 0))
        < EMAIL_CODE_RESEND_SECONDS
    ):
        wait = EMAIL_CODE_RESEND_SECONDS - (
            now - int(previous.get("created_at", 0))
        )

        raise HTTPException(
            status_code=429,
            detail=f"Повторно запросить код можно через {wait} сек."
        )

    code = f"{secrets.randbelow(1_000_000):06d}"

    codes[email] = {
        "code_hash": hashlib.sha256(
            code.encode("ascii")
        ).hexdigest(),
        "password_hash": _hash_password(password),
        "display_name": display_name.strip()[:80],
        "created_at": now,
        "expires_at": now + EMAIL_CODE_TTL_SECONDS,
        "attempts": 0,
    }

    _save_email_codes(codes)

    try:
        _send_verification_email(email, code)

    except HTTPException:
        codes.pop(email, None)
        _save_email_codes(codes)
        raise


def _require_auth(
    authorization: Optional[str] = Header(default=None)
) -> dict:

    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Сначала войдите на сайт."
        )

    token = authorization[7:].strip()

    email_user = _verify_email_session(token)

    if email_user:
        return email_user

    payload = verify_session_token(
        token,
        SITE_AUTH_SECRET
    )

    if not payload:
        raise HTTPException(
            status_code=401,
            detail="Сессия истекла. Войдите снова."
        )

    user_id = int(payload["user_id"])

    if any(
        int(item.get("user_id", 0) or 0) == user_id
        for item in _load_accounts().values()
    ):
        raise HTTPException(
            status_code=401,
            detail="Сессия отозвана или аккаунт заблокирован. Войдите снова."
        )

    if not bot._has_access(user_id):
        raise HTTPException(
            status_code=403,
            detail="Доступ к сервису отозван владельцем."
        )

    return {
        "user_id": user_id,
        "is_owner": bot._is_admin(user_id),
        "email": None,
        "display_name": "",
        "auth_type": "telegram",
    }


def _require_owner(
    user: dict = Depends(_require_auth)
) -> dict:

    if not user["is_owner"]:
        raise HTTPException(
            status_code=403,
            detail="Раздел доступен только владельцу."
        )

    return user


def _safe_error_message(exc: Exception) -> str:
    text = bot._safe_error_text(exc)
    return text[:500]


@app.get("/api/health")
def health():
    return {
        "ok": True,
        "service": API_TITLE,
        "auth_configured": len(SITE_AUTH_SECRET) >= 32
    }


@app.post("/api/auth/register")
def email_register(payload: EmailCredentials):
    email = _normalise_gmail(payload.email)

    if len(payload.password) < 8:
        raise HTTPException(
            status_code=400,
            detail="Пароль должен содержать не менее 8 символов."
        )

    if email in _load_accounts():
        raise HTTPException(
            status_code=409,
            detail="Этот Gmail уже зарегистрирован. Попробуйте войти."
        )

    _create_and_send_email_code(
        email,
        payload.password,
        payload.display_name
    )

    return {
        "ok": True,
        "verification_required": True,
        "message": "Код подтверждения отправлен на Gmail. Введите его на сайте в течение 10 минут."
    }


@app.post("/api/auth/resend")
def email_resend(payload: EmailCredentials):
    email = _normalise_gmail(payload.email)

    if email in _load_accounts():
        raise HTTPException(
            status_code=409,
            detail="Этот Gmail уже подтверждён. Попробуйте войти."
        )

    _create_and_send_email_code(
        email,
        payload.password,
        payload.display_name,
        allow_resend=True
    )

    return {
        "ok": True,
        "message": "Новый код отправлен на Gmail."
    }


@app.post("/api/auth/verify")
def email_verify(payload: EmailVerifyRequest):
    email = _normalise_gmail(payload.email)
    codes = _load_email_codes()
    item = codes.get(email)
    now = int(time.time())

    if not isinstance(item, dict):
        raise HTTPException(
            status_code=400,
            detail="Код не найден. Начните регистрацию заново."
        )

    if int(item.get("expires_at", 0)) <= now:
        codes.pop(email, None)
        _save_email_codes(codes)

        raise HTTPException(
            status_code=400,
            detail="Срок действия кода истёк. Запросите регистрацию ещё раз."
        )

    supplied = hashlib.sha256(
        payload.code.encode("ascii")
    ).hexdigest()

    if not hmac.compare_digest(
        str(item.get("code_hash", "")),
        supplied
    ):
        item["attempts"] = int(item.get("attempts", 0)) + 1

        if item["attempts"] >= 5:
            codes.pop(email, None)
        else:
            codes[email] = item

        _save_email_codes(codes)

        raise HTTPException(
            status_code=401,
            detail="Код неверный. Осталось ограниченное число попыток."
        )

    accounts = _load_accounts()

    if email in accounts:
        codes.pop(email, None)
        _save_email_codes(codes)

        raise HTTPException(
            status_code=409,
            detail="Этот Gmail уже зарегистрирован. Выполните вход."
        )

    account = {
        "email": email,
        "display_name": str(item.get("display_name", ""))[:80],
        "password_hash": str(item.get("password_hash", "")),
        "user_id": _new_account_user_id(accounts),
        "verified": True,
        "active": True,
        "session_version": 1,
        "created_at": datetime_now_iso(),
    }

    accounts[email] = account
    _save_accounts(accounts)

    codes.pop(email, None)
    _save_email_codes(codes)

    token, ttl = _email_session_token(account)
    role = _account_role(email, account)

    return {
        "access_token": token,
        "token_type": "Bearer",
        "expires_in": ttl,
        "user": {
            "email": email,
            "display_name": account["display_name"],
            "role": role,
            "is_owner": role == "owner"
        }
    }


@app.post("/api/auth/login")
def email_login(payload: EmailCredentials):
    email = _normalise_gmail(payload.email)
    account = _load_accounts().get(email)

    if (
        not account
        or not account.get("verified")
        or not _verify_password(
            payload.password,
            str(account.get("password_hash", ""))
        )
    ):
        raise HTTPException(
            status_code=401,
            detail="Неверный Gmail или пароль."
        )

    if not account.get("active", True):
        raise HTTPException(
            status_code=403,
            detail="Учётная запись заблокирована владельцем сайта."
        )

    token, ttl = _email_session_token(account)
    role = _account_role(email, account)

    return {
        "access_token": token,
        "token_type": "Bearer",
        "expires_in": ttl,
        "user": {
            "email": email,
            "display_name": account.get("display_name", ""),
            "role": role,
            "is_owner": role == "owner"
        }
    }


@app.post("/api/login")
def login(payload: LoginRequest):
    """Legacy Telegram login remains available for users of /sitecode."""

    user_id = int(payload.telegram_id)

    if not bot._has_access(user_id):
        raise HTTPException(
            status_code=403,
            detail="Доступ ещё не выдан. Сначала получите разрешение владельца."
        )

    if not SITE_AUTH_SECRET or len(SITE_AUTH_SECRET) < 32:
        raise HTTPException(
            status_code=503,
            detail="На сервере ещё не настроен SITE_AUTH_SECRET."
        )

    if not consume_site_login_code(user_id, payload.code):
        raise HTTPException(
            status_code=401,
            detail="Код неверный, истёк или уже использован. Запросите новый: /sitecode"
        )

    token, ttl = create_session_token(
        user_id,
        SITE_AUTH_SECRET
    )

    return {
        "access_token": token,
        "token_type": "Bearer",
        "expires_in": ttl,
        "user": {
            "telegram_id": user_id,
            "role": "owner" if bot._is_admin(user_id) else "user"
        }
    }


@app.get("/api/me")
def me(user: dict = Depends(_require_auth)):
    return {
        "telegram_id": user["user_id"] if user.get("auth_type") == "telegram" else None,
        "email": user.get("email"),
        "display_name": user.get("display_name", ""),
        "auth_type": user.get("auth_type", "telegram"),
        "role": "owner" if user["is_owner"] else "user",
        "is_owner": bool(user["is_owner"])
    }


@app.get("/api/history")
def history(user: dict = Depends(_require_auth)):
    items = bot._history_entries_for(user["user_id"])

    return {
        "items": [{
            "id": item.get("id", ""),
            "created_at": item.get("created_at", ""),
            "title": item.get("title", "Конспект"),
            "discipline": item.get("discipline", ""),
            "topic": item.get("topic", ""),
            "filename": item.get("filename", "Конспект.docx"),
            "download_url": f"/api/history/{item.get('id', '')}/download",
        } for item in items]
    }


@app.get("/api/history/{entry_id}/download")
def history_download(
    entry_id: str,
    user: dict = Depends(_require_auth)
):
    item, path = bot._history_item_by_id(
        user["user_id"],
        entry_id
    )

    if not item or not path:
        raise HTTPException(
            status_code=404,
            detail="Документ не найден в вашей истории."
        )

    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=item.get("filename") or path.name
    )


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

            items.append({
                "name": name,
                "extension": path.suffix.lower(),
                "categories": categories
            })

        return {
            "items": items,
            "count": len(items)
        }

    except Exception as exc:
        bot._record_error(
            exc,
            user_id=user["user_id"],
            event="website_books_list"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось прочитать каталог учебников."
        ) from exc


@app.get("/api/books/search")
def book_search(
    q: str = Query(min_length=3, max_length=160),
    user: dict = Depends(_require_auth)
):
    if bot._restricted_book_query(q):
        raise HTTPException(
            status_code=400,
            detail="Поиск доступен для общих учебных тем без прикладных инструкций по оружию или тактике."
        )

    try:
        matches = bot._search_safe_book_chunks(
            q,
            max_results=8
        )

        return {
            "query": q,
            "items": [
                {"source": label, "snippet": snippet}
                for label, snippet in matches
            ]
        }

    except Exception as exc:
        bot._record_error(
            exc,
            user_id=user["user_id"],
            event="website_book_search"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось выполнить поиск по учебникам."
        ) from exc


@app.post("/api/reports")
def submit_report(
    payload: ReportRequest,
    user: dict = Depends(_require_auth)
):
    message = payload.message.strip()
    lowered = message.casefold()

    if any(
        marker in lowered
        for marker in (
            "telegram_token=",
            "gemini_api_key=",
            "api_key=",
            "password="
        )
    ):
        raise HTTPException(
            status_code=400,
            detail="Не отправляйте токены, пароли или секретные данные."
        )

    try:
        item = bot._save_user_report(
            user["user_id"],
            "website-user",
            f"[{payload.kind.strip()}] {message}"
        )

        return {
            "ok": True,
            "time": item.get("time")
        }

    except Exception as exc:
        bot._record_error(
            exc,
            user_id=user["user_id"],
            event="website_report"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось записать сообщение. Попробуйте позже."
        ) from exc


@app.get("/api/owner/overview")
def owner_overview(user: dict = Depends(_require_owner)):
    accounts = _load_accounts()
    email_count = len(accounts)

    return {
        "authorized_users": len(bot._load_authorized_users()) + email_count,
        "created_docs": int(bot.BOT_STATS.get("created_docs", 0)),
        "errors": int(bot.BOT_STATS.get("errors", 0)),
        "last_created_at": bot.BOT_STATS.get("last_created_at"),
        "last_created_by": bot.BOT_STATS.get("last_created_by"),
    }


@app.get("/api/owner/users")
def owner_users(user: dict = Depends(_require_owner)):
    items = [
        {
            "type": "telegram",
            "telegram_id": uid,
            "role": "user",
            "access": "granted",
            "display_name": "Telegram-пользователь"
        }
        for uid in sorted(bot._load_authorized_users())
    ]

    accounts = _load_accounts()

    for email, account in sorted(accounts.items()):
        role = _account_role(email, account)

        items.append({
            "type": "email",
            "email": email,
            "display_name": account.get("display_name", ""),
            "role": role,
            "access": "active" if account.get("active", True) else "blocked",
            "created_at": account.get("created_at", "")
        })

    return {"items": items}


@app.post("/api/owner/email-users/block")
def owner_block_email_user(
    payload: EmailOnlyRequest,
    user: dict = Depends(_require_owner)
):
    email = _normalise_gmail(payload.email)
    accounts = _load_accounts()
    account = accounts.get(email)

    if not account:
        raise HTTPException(
            status_code=404,
            detail="Пользователь с таким Gmail не найден."
        )

    if _account_role(email, account) == "owner":
        raise HTTPException(
            status_code=400,
            detail="Нельзя заблокировать аккаунт владельца."
        )

    account["active"] = False
    account["session_version"] = int(
        account.get("session_version", 1)
    ) + 1

    accounts[email] = account
    _save_accounts(accounts)

    return {
        "ok": True,
        "email": email,
        "access": "blocked",
        "sessions_revoked": True
    }


@app.post("/api/owner/email-users/unblock")
def owner_unblock_email_user(
    payload: EmailOnlyRequest,
    user: dict = Depends(_require_owner)
):
    email = _normalise_gmail(payload.email)
    accounts = _load_accounts()
    account = accounts.get(email)

    if not account:
        raise HTTPException(
            status_code=404,
            detail="Пользователь с таким Gmail не найден."
        )

    account["active"] = True

    account["session_version"] = int(
        account.get("session_version", 1)
    ) + 1

    accounts[email] = account
    _save_accounts(accounts)

    return {
        "ok": True,
        "email": email,
        "access": "active"
    }


@app.post("/api/owner/users/grant")
def owner_grant(
    payload: AccessRequest,
    user: dict = Depends(_require_owner)
):
    target = int(payload.telegram_id)

    if target in bot.OWNER_IDS:
        raise HTTPException(
            status_code=400,
            detail="Владелец уже имеет полный доступ."
        )

    bot.AUTHORIZED_USER_IDS.add(target)

    try:
        bot._save_authorized_users()

    except Exception as exc:
        bot.AUTHORIZED_USER_IDS.discard(target)

        bot._record_error(
            exc,
            user_id=user["user_id"],
            event="website_owner_grant"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось сохранить список пользователей."
        ) from exc

    return {
        "ok": True,
        "telegram_id": target,
        "role": "user",
        "access": "granted"
    }


@app.post("/api/owner/users/revoke")
def owner_revoke(
    payload: AccessRequest,
    user: dict = Depends(_require_owner)
):
    target = int(payload.telegram_id)

    if target in bot.OWNER_IDS:
        raise HTTPException(
            status_code=400,
            detail="Нельзя отозвать доступ у владельца."
        )

    was_present = target in bot.AUTHORIZED_USER_IDS
    bot.AUTHORIZED_USER_IDS.discard(target)

    try:
        bot._save_authorized_users()

    except Exception as exc:
        if was_present:
            bot.AUTHORIZED_USER_IDS.add(target)

        bot._record_error(
            exc,
            user_id=user["user_id"],
            event="website_owner_revoke"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось сохранить список пользователей."
        ) from exc

    return {
        "ok": True,
        "telegram_id": target,
        "role": "user",
        "access": "revoked"
    }


@app.get("/api/owner/errors")
def owner_errors(user: dict = Depends(_require_owner)):
    return {"items": bot._read_recent_errors(20)}


@app.get("/api/owner/reports")
def owner_reports(user: dict = Depends(_require_owner)):
    return {"items": bot._read_user_reports(20)}


def _generate_doc(
    payload: CreateDocumentRequest,
    user_id: int
) -> tuple[str, list[str], str]:

    discipline = payload.discipline.strip()

    if bot.normalize_text(discipline) not in SAFE_WEB_DISCIPLINES:
        raise HTTPException(
            status_code=400,
            detail="Через сайт пока доступны только общие учебные дисциплины."
        )

    combined_topic = f"{payload.topic} {payload.lesson_title}"

    if bot._restricted_book_query(combined_topic):
        raise HTTPException(
            status_code=400,
            detail="Эта тема не поддерживается веб-генератором. Выберите общую учебную тему без прикладных инструкций."
        )

    match = None

    try:
        match = bot.find_program_lesson(
            discipline,
            payload.topic_number,
            payload.lesson_number
        )

    except Exception as exc:
        bot._record_error(
            exc,
            user_id=user_id,
            event="website_program_lookup"
        )

    topic_title = payload.topic.strip()
    lesson_title = payload.lesson_title.strip()

    if isinstance(match, dict) and match.get("multiple"):
        candidates = match.get("multiple", [])

        exact = next(
            (
                x for x in candidates
                if bot.normalize_text(
                    x.get("lesson_title", "")
                ) == bot.normalize_text(lesson_title)
            ),
            None
        )

        if exact:
            topic_title = exact.get("topic", topic_title)
            lesson_title = exact.get("lesson_title", lesson_title)

    elif isinstance(match, dict) and match:
        topic_title = match.get("topic", topic_title) or topic_title
        lesson_title = match.get("lesson_title", lesson_title) or lesson_title

    output_path = None

    try:
        data = bot.generate_lesson_data(
            discipline,
            topic_title,
            lesson_title,
            payload.lesson_type.strip(),
            payload.audience.strip(),
            payload.duration.strip(),
            payload.goals.strip(),
            payload.manuals.strip(),
            payload.materials.strip()
        )

        with tempfile.NamedTemporaryFile(
            suffix=".docx",
            delete=False
        ) as temp:
            output_path = temp.name

        bot.create_from_template(
            payload.rank.strip(),
            payload.full_name.strip(),
            discipline,
            payload.topic_number.strip(),
            topic_title,
            payload.lesson_number.strip(),
            lesson_title,
            payload.lesson_type.strip(),
            payload.audience.strip(),
            payload.location.strip(),
            payload.duration.strip(),
            payload.goals.strip(),
            data,
            output_path,
            payload.approval_position.strip()
            or "Командир артиллерийского дивизиона"
        )

        warnings = bot.validate_generated_docx(
            output_path,
            expected_topic=topic_title,
            expected_title=lesson_title
        )

        bot._archive_document(
            user_id,
            output_path,
            lesson_title,
            discipline,
            topic_title
        )

        bot._record_generated_doc(
            user_id,
            "конспект через сайт"
        )

        safe_name = re.sub(
            r"[^A-Za-zА-Яа-яЁё0-9_-]+",
            "_",
            lesson_title
        )[:60].strip("_") or "Konspekt"

        return (
            output_path,
            warnings,
            f"Конспект_{safe_name}.docx"
        )

    except HTTPException:
        if output_path and os.path.exists(output_path):
            os.remove(output_path)

        raise

    except Exception as exc:
        if output_path and os.path.exists(output_path):
            os.remove(output_path)

        bot._record_error(
            exc,
            user_id=user_id,
            event="website_doc_generation"
        )

        raise HTTPException(
            status_code=500,
            detail="Не удалось создать Word-файл. Проверьте шаблон и настройки сервера, затем попробуйте ещё раз."
        ) from exc


@app.post("/api/create")
async def create_document(
    payload: CreateDocumentRequest,
    user: dict = Depends(_require_auth)
):
    output_path, warnings, filename = await run_in_threadpool(
        _generate_doc,
        payload,
        user["user_id"]
    )

    warning_header = (
        " | ".join(warnings)[:1000]
        if warnings
        else "none"
    )

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
        headers={"X-Quality-Warnings": warning_header},
        background=BackgroundTask(
            lambda: Path(output_path).unlink(missing_ok=True)
        )
    )
