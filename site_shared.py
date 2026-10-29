"""Shared one-time login codes and signed sessions for the website API."""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
LOGIN_CODES_FILE = BASE_DIR / "site_login_codes.json"
_LOCK = threading.Lock()
CODE_TTL_SECONDS = 600
SESSION_TTL_SECONDS = 6 * 60 * 60


def _load_codes() -> dict[str, Any]:
    try:
        payload = json.loads(LOGIN_CODES_FILE.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else {}
    except (OSError, json.JSONDecodeError, TypeError):
        return {}


def _save_codes(payload: dict[str, Any]) -> None:
    temporary = LOGIN_CODES_FILE.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, LOGIN_CODES_FILE)


def issue_site_login_code(user_id: int) -> tuple[str, int]:
    """Create an 8-digit, single-use code whose hash is stored on disk."""
    uid = str(int(user_id))
    code = f"{secrets.randbelow(100_000_000):08d}"
    now = int(time.time())
    entry = {
        "code_hash": hashlib.sha256(code.encode("ascii")).hexdigest(),
        "expires_at": now + CODE_TTL_SECONDS,
        "attempts": 0,
    }
    with _LOCK:
        items = _load_codes()
        # Drop expired records on every issuance.
        items = {key: value for key, value in items.items()
                 if isinstance(value, dict) and int(value.get("expires_at", 0)) > now}
        items[uid] = entry
        _save_codes(items)
    return code, CODE_TTL_SECONDS


def consume_site_login_code(user_id: int, code: str) -> bool:
    """Validate and consume a one-time code; limit guesses to five attempts."""
    uid = str(int(user_id))
    now = int(time.time())
    supplied_hash = hashlib.sha256(str(code or "").strip().encode("ascii", errors="ignore")).hexdigest()
    with _LOCK:
        items = _load_codes()
        entry = items.get(uid)
        if not isinstance(entry, dict):
            return False
        if int(entry.get("expires_at", 0)) <= now:
            items.pop(uid, None)
            _save_codes(items)
            return False
        expected = str(entry.get("code_hash", ""))
        if not hmac.compare_digest(expected, supplied_hash):
            entry["attempts"] = int(entry.get("attempts", 0)) + 1
            if entry["attempts"] >= 5:
                items.pop(uid, None)
            else:
                items[uid] = entry
            _save_codes(items)
            return False
        items.pop(uid, None)
        _save_codes(items)
        return True


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _unb64url(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * ((4 - len(value) % 4) % 4))


def create_session_token(user_id: int, secret: str, ttl_seconds: int = SESSION_TTL_SECONDS) -> tuple[str, int]:
    if not secret or len(secret) < 32:
        raise ValueError("SITE_AUTH_SECRET должен содержать не менее 32 символов.")
    now = int(time.time())
    payload = {"uid": int(user_id), "iat": now, "exp": now + int(ttl_seconds), "nonce": secrets.token_hex(8)}
    encoded = _b64url(json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8"))
    signature = hmac.new(secret.encode("utf-8"), encoded.encode("ascii"), hashlib.sha256).digest()
    return encoded + "." + _b64url(signature), int(ttl_seconds)


def verify_session_token(token: str, secret: str) -> dict[str, Any] | None:
    if not secret or len(secret) < 32 or not token or "." not in token:
        return None
    encoded, supplied_signature = token.split(".", 1)
    expected_signature = _b64url(hmac.new(secret.encode("utf-8"), encoded.encode("ascii"), hashlib.sha256).digest())
    if not hmac.compare_digest(expected_signature, supplied_signature):
        return None
    try:
        payload = json.loads(_unb64url(encoded).decode("utf-8"))
        if not isinstance(payload, dict) or int(payload.get("exp", 0)) <= int(time.time()):
            return None
        uid = int(payload.get("uid", 0))
        if uid <= 0:
            return None
        return {"user_id": uid, "expires_at": int(payload["exp"])}
    except (ValueError, TypeError, json.JSONDecodeError, UnicodeDecodeError):
        return None
