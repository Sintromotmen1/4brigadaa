import os
import json
import tempfile
import time
import random
import traceback
import re
import io
import hashlib
import html
import urllib.parse
import shutil
import subprocess
from functools import lru_cache
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv
from docx import Document
from pypdf import PdfReader
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw, ImageFont
from google import genai
from site_shared import issue_site_login_code

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

if not TELEGRAM_TOKEN:
    raise ValueError("Не найден в файле .env")
if not GEMINI_API_KEY:
    raise ValueError("Не найден  в файле .env")
client = genai.Client(api_key=GEMINI_API_KEY)
TEMPLATE_PATH = Path(__file__).parent / "template.docx"
# Положите программу боевой подготовки рядом с bot.py. Поддерживается исходное имя файла.
PROGRAM_PATHS = [
    Path(__file__).parent / "8. Программа БП адн рус 2026(1).docx",
    Path(__file__).parent / "8. Программа БП адн рус 2026.docx",
    Path(__file__).parent / "program.docx",
]
PROGRAM_PATH = next((path for path in PROGRAM_PATHS if path.exists()), PROGRAM_PATHS[0])


# Управление доступом к боту.
# OWNER_ID — единственный владелец. Все остальные получают только обычный доступ.
# Только OWNER_ID/BOT_OWNER_ID определяет владельца. Обычные пользователи — через authorized_users.json.
ACCESS_FILE = Path(__file__).parent / "authorized_users.json"

STATS_FILE = Path(__file__).parent / "owner_stats.json"
ERROR_LOG_FILE = Path(__file__).parent / "bot_errors.log"


def _default_stats():
    return {
        "created_docs": 0,
        "created_by_user": {},
        "created_types": {},
        "errors": 0,
        "last_created_at": None,
        "last_created_by": None,
        "last_created_type": None,
        "started_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }


def _load_stats():
    defaults = _default_stats()
    try:
        if not STATS_FILE.exists():
            return defaults
        loaded = json.loads(STATS_FILE.read_text(encoding="utf-8"))
        if not isinstance(loaded, dict):
            return defaults
        defaults.update(loaded)
        if not isinstance(defaults.get("created_by_user"), dict):
            defaults["created_by_user"] = {}
        if not isinstance(defaults.get("created_types"), dict):
            defaults["created_types"] = {}
        return defaults
    except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
        print(f"Не удалось загрузить статистику владельца: {exc}")
        return defaults


BOT_STATS = _load_stats()


def _save_stats():
    temporary = STATS_FILE.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(BOT_STATS, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, STATS_FILE)


def _record_generated_doc(user_id, doc_type="конспект"):
    key = str(int(user_id))
    BOT_STATS["created_docs"] = int(BOT_STATS.get("created_docs", 0)) + 1
    by_user = BOT_STATS.setdefault("created_by_user", {})
    by_user[key] = int(by_user.get(key, 0)) + 1
    by_type = BOT_STATS.setdefault("created_types", {})
    by_type[doc_type] = int(by_type.get(doc_type, 0)) + 1
    BOT_STATS["last_created_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    BOT_STATS["last_created_by"] = key
    BOT_STATS["last_created_type"] = doc_type
    try:
        _save_stats()
    except OSError as exc:
        print(f"Не удалось сохранить статистику документов: {exc}")


def _safe_error_text(error):
    value = f"{type(error).__name__}: {error}"
    # Never write live credentials into a persistent error log.
    for secret_name in ("TELEGRAM_TOKEN", "GEMINI_API_KEY"):
        secret = os.getenv(secret_name, "")
        if secret:
            value = value.replace(secret, "[СКРЫТО]")
    value = re.sub(r"\b\d{8,}:[A-Za-z0-9_-]{20,}\b", "[СКРЫТЫЙ_ТОКЕН]", value)
    return value.replace("\n", " ")[:700]


def _record_error(error, user_id=None, event="bot"):
    BOT_STATS["errors"] = int(BOT_STATS.get("errors", 0)) + 1
    try:
        _save_stats()
    except OSError as exc:
        print(f"Не удалось сохранить счётчик ошибок: {exc}")
    entry = {
        "time": datetime.now().astimezone().isoformat(timespec="seconds"),
        "event": str(event)[:80],
        "user_id": str(user_id) if user_id is not None else None,
        "error": _safe_error_text(error),
    }
    try:
        with ERROR_LOG_FILE.open("a", encoding="utf-8") as log_file:
            log_file.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except OSError as exc:
        print(f"Не удалось записать журнал ошибок: {exc}")


def _read_recent_errors(limit=10):
    if not ERROR_LOG_FILE.exists():
        return []
    try:
        lines = ERROR_LOG_FILE.read_text(encoding="utf-8").splitlines()[-limit:]
        entries = []
        for line in reversed(lines):
            try:
                item = json.loads(line)
                entries.append(item)
            except json.JSONDecodeError:
                entries.append({"time": "время неизвестно", "event": "log", "error": line[:300]})
        return entries
    except OSError:
        return []



HISTORY_ROOT = Path(__file__).parent / "document_history"
HISTORY_INDEX_FILE = Path(__file__).parent / "document_history.json"
USER_REPORTS_FILE = Path(__file__).parent / "user_reports.jsonl"
HISTORY_LIMIT_PER_USER = 15


def _load_document_history():
    try:
        data = json.loads(HISTORY_INDEX_FILE.read_text(encoding="utf-8")) if HISTORY_INDEX_FILE.exists() else []
        return data if isinstance(data, list) else []
    except (OSError, json.JSONDecodeError, TypeError):
        return []


def _save_document_history(items):
    temp = HISTORY_INDEX_FILE.with_suffix(".json.tmp")
    temp.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, HISTORY_INDEX_FILE)


def _archive_document(user_id, source_path, title, discipline="", topic=""):
    """Keep a personal Word copy so the user can download it again later."""
    from datetime import timezone
    user_dir = HISTORY_ROOT / str(int(user_id))
    user_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().astimezone()
    safe_title = re.sub(r"[^A-Za-zА-Яа-яЁё0-9_-]+", "_", str(title or "Конспект"))[:60].strip("_") or "Конспект"
    entry_id = timestamp.strftime("%Y%m%d%H%M%S") + "_" + str(random.randint(100, 999))
    dest = user_dir / f"{entry_id}_{safe_title}.docx"
    shutil.copy2(source_path, dest)
    items = _load_document_history()
    items.append({
        "id": entry_id,
        "user_id": int(user_id),
        "created_at": timestamp.isoformat(timespec="minutes"),
        "title": str(title or "Конспект")[:180],
        "discipline": str(discipline or "")[:100],
        "topic": str(topic or "")[:180],
        "path": dest.relative_to(Path(__file__).parent).as_posix(),
        "filename": dest.name,
    })
    user_entries = sorted((x for x in items if int(x.get("user_id", 0) or 0) == int(user_id)), key=lambda x: x.get("created_at", ""), reverse=True)
    for old in user_entries[HISTORY_LIMIT_PER_USER:]:
        try:
            old_path = Path(__file__).parent / str(old.get("path", ""))
            if old_path.resolve().is_relative_to(HISTORY_ROOT.resolve()) and old_path.exists():
                old_path.unlink()
        except (OSError, ValueError, RuntimeError):
            pass
        items = [x for x in items if x.get("id") != old.get("id") or int(x.get("user_id", 0) or 0) != int(user_id)]
    _save_document_history(items)
    return entry_id


def _history_entries_for(user_id, limit=HISTORY_LIMIT_PER_USER):
    items = [x for x in _load_document_history() if int(x.get("user_id", 0) or 0) == int(user_id)]
    items.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    base = Path(__file__).parent
    valid = []
    for item in items:
        path = (base / str(item.get("path", ""))).resolve()
        try:
            if path.is_relative_to(HISTORY_ROOT.resolve()) and path.is_file():
                valid.append(item)
        except (OSError, ValueError, RuntimeError):
            continue
    return valid[:limit]


def _save_user_report(user_id, username, report_text):
    entry = {
        "time": datetime.now().astimezone().isoformat(timespec="seconds"),
        "user_id": int(user_id),
        "username": str(username or "")[:100],
        "message": str(report_text or "").strip()[:3000],
    }
    with USER_REPORTS_FILE.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return entry


def _read_user_reports(limit=10):
    if not USER_REPORTS_FILE.exists():
        return []
    try:
        raw_lines = USER_REPORTS_FILE.read_text(encoding="utf-8").splitlines()[-limit:]
        result = []
        for line in reversed(raw_lines):
            try:
                item = json.loads(line)
                if isinstance(item, dict):
                    result.append(item)
            except json.JSONDecodeError:
                continue
        return result
    except OSError:
        return []


def _parse_owner_ids():
    """Return at most one owner ID. Legacy ADMIN_IDS is intentionally ignored."""
    raw_owner = (
        os.getenv("OWNER_ID", "").strip()
        or os.getenv("BOT_OWNER_ID", "").strip()
    )
    if raw_owner.isdigit():
        return {int(raw_owner)}
    return set()


def _load_authorized_users():
    try:
        if not ACCESS_FILE.exists():
            return set()
        raw = json.loads(ACCESS_FILE.read_text(encoding="utf-8"))
        # Поддерживаются как простой список ID, так и объект {"users": [...] }.
        if isinstance(raw, dict):
            raw = raw.get("users", [])
        if not isinstance(raw, list):
            return set()
        return {int(value) for value in raw if str(value).strip().isdigit()}
    except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
        print(f"Не удалось прочитать список пользователей {ACCESS_FILE.name}: {exc}")
        return set()


def _save_authorized_users():
    """Атомарно сохраняет список разрешённых Telegram ID."""
    temporary_path = ACCESS_FILE.with_suffix(".json.tmp")
    temporary_path.write_text(
        json.dumps({"users": sorted(AUTHORIZED_USER_IDS)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temporary_path, ACCESS_FILE)


OWNER_IDS = _parse_owner_ids()
AUTHORIZED_USER_IDS = _load_authorized_users()


def _is_admin(user_id):
    return user_id is not None and int(user_id) in OWNER_IDS


def _has_access(user_id):
    return user_id is not None and (int(user_id) in OWNER_IDS or int(user_id) in AUTHORIZED_USER_IDS)


def _access_denied_text(user):
    user_id = user.id if user else "не определён"
    return (
        "🔒 ДОСТУП НЕ ОФОРМЛЕН\n"
        "━━━━━━━━━━━━━━━━━━\n\n"
        "Для использования бота нужен доступ от владельца.\n\n"
        f"🆔 Ваш Telegram ID: {user_id}\n\n"
        "Скопируйте этот номер и отправьте владельцу бота. После подтверждения вы сможете пользоваться функциями «Военного Конспекта».\n\n"
        "Если вы владелец, проверьте OWNER_ID в файле .env."
    )

BUTTON_CONFIRM_YES = "✅ Да, верно"
BUTTON_CONFIRM_NO = "❌ Нет, изменить"

BUTTON_NEW = "📝 Создать новый конспект"
BUTTON_OWNER_DASHBOARD = "🛡 Кабинет владельца"
BUTTON_OWNER_GRANT = "➕ Выдать доступ"
BUTTON_OWNER_REVOKE = "➖ Отозвать доступ"
BUTTON_OWNER_USERS = "👥 Список пользователей"
BUTTON_OWNER_ERRORS = "📋 Журнал ошибок"
BUTTON_OWNER_STATS = "📊 Обновить статистику"
BUTTON_OWNER_BACK = "🏠 В главное меню"
BUTTON_SPECIAL = "🎯 По специальности"
BUTTON_TECHNICAL = "🛠 Техническая подготовка"
BUTTON_PICK_PREFIX = "Выбрать вариант "
BUTTON_BOOKS = "📚 Учебники"
BUTTON_HISTORY = "🗂 Мои конспекты"
BUTTON_SEARCH_BOOKS = "🔎 Поиск в учебниках"
BUTTON_REPORT_ERROR = "🆘 Сообщить об ошибке"
BUTTON_OWNER_REPORTS = "📨 Сообщения пользователей"
BUTTON_MEDICAL_BOOKS = "📘 Военно-медицинская подготовка"
BUTTON_LAW_BOOKS = "📕 Государственно-правовая подготовка"
BUTTON_REGULATION_BOOKS = "📗 Общевоинские уставы"
BUTTON_DRILL_BOOKS = "📙 Строевая подготовка"
BUTTON_PHYSICAL_BOOKS = "🏃 Физическая подготовка"
BUTTON_PSYCHOLOGY_BOOKS = "🧠 Психологическая подготовка"
BUTTON_MATERIAL_BOOKS = "📦 Материальное обеспечение"
BUTTON_BACK_TO_MENU = "↩️ Главное меню"
BUTTON_HELP = "ℹ️ Помощь"
BUTTON_CANCEL = "❌ Отмена"
BUTTON_RETRY = "🔄 Повторить создание"
BUTTON_CUSTOM = "✍️ Ввести свой вариант"
BUTTON_SKIP = "⏭️ Не указывать"

# Дисциплины взяты из заголовков разделов программы боевой подготовки 2026.
DISCIPLINES = [
    "Базовая общевойсковая подготовка",
    "Подготовка по специальности",
    "Подготовка расчетов (отделений)",
    "Подготовка взводов",
    "Подготовка батарей",
    "Подготовка дивизионов",
    "Тактическая подготовка",
    "Огневая подготовка",
    "Военно-медицинская подготовка",
    "Подготовка по связи",
    "Инженерная подготовка",
    "Военная топография",
    "Государственно-правовая подготовка",
    "Разведывательная подготовка",
    "Общевоинские уставы",
    "Строевая подготовка",
    "Физическая подготовка",
    "Техническая подготовка",
    "Специальная подготовка",
    "Вождение для механиков – водителей и водителей",
    "Радиационная, химическая и биологическая защита",
    "Занятия по боевой готовности",
]

# Форматы занятий соответствуют формулировкам, встречающимся в программе.
LESSON_TYPES = [
    "Рассказ",
    "Групповое занятие",
    "Практическое занятие",
    "Контрольное занятие",
    "Тренировка",
    "Тактико-строевое занятие",
    "Самостоятельная подготовка",
    "Учение",
]

MANUAL_OPTIONS = [
    "📘 Программа боевой подготовки",
    "📗 Руководство",
    "📙 Учебное пособие",
]
MATERIAL_OPTIONS = [
    "📋 Плакаты и схемы",
    "🖥️ Проектор и экран",
    "🧰 Учебное оборудование",
    "📚 Учебные материалы",
]


def normalize_text(value):
    value = (value or "").casefold().replace("ё", "е")
    return re.sub(r"\s+", " ", value).strip()


def remove_program_lesson_duration_phrase(value):
    """Удаляет из текста служебную пометку вида «(практическое занятие–2часа)»."""
    value = str(value or "")
    value = re.sub(
        r"\s*\(\s*практическое\s+занятие\s*[–—-]\s*\d+\s*час(?:а|ов)?\s*\)\s*\.?",
        "",
        value,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s{2,}", " ", value).strip()


def find_program_lesson(discipline, topic_number, lesson_number):
    """Находит тему/занятие в программе и возвращает только заголовочные сведения."""
    if not PROGRAM_PATH.exists():
        raise FileNotFoundError(
            "Не найдена программа боевой подготовки. Положите файл программы рядом с bot.py."
        )
    doc = Document(str(PROGRAM_PATH))
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    # В документе названия дисциплин обычно оформлены отдельными строками заглавными буквами.
    sections = []
    current_section = ""
    topic_pattern = re.compile(r"^тема\s*№?\s*(\d+)\s*[.:]?\s*(.*)$", re.I)
    lesson_pattern = re.compile(r"^заняти(?:е|я)\s*№?\s*(\d+)\b\s*(.*)$", re.I | re.S)
    target_section = normalize_text(discipline)
    candidates = []
    current_topic = None
    current_section = ""
    for index, line in enumerate(paragraphs):
        if not line:
            continue
        normalized = normalize_text(line)
        # Заголовки разделов обычно короткие и набраны заглавными буквами.
        is_topic = topic_pattern.match(normalized)
        is_lesson = lesson_pattern.match(normalized)
        if not is_topic and not is_lesson and len(line) < 100 and line == line.upper() and any(ch.isalpha() for ch in line):
            current_section = normalized
            current_topic = None
            continue
        if is_topic:
            match = topic_pattern.match(normalized)
            topic_title = match.group(2).strip(" .:—-")
            # Убираем служебные пометки в скобках, например «(планируется летом)».
            topic_title = re.sub(r"\s*\([^)]*\)", "", topic_title).strip(" .:—-")
            current_topic = {"number": match.group(1), "title": topic_title, "section": current_section}
            continue
        if is_lesson and current_topic:
            match = lesson_pattern.match(normalized)
            if match.group(1) != str(lesson_number):
                continue
            section_match = target_section in current_topic["section"] or current_topic["section"] in target_section
            # Keep only section/topic/lesson headings; do not pass the full operational description to the confirmation prompt.
            if current_topic["number"] == str(topic_number) and section_match:
                lesson_tail = match.group(2).strip()
                lesson_kind = re.split(r"[–—]\s*\d|\s+\d+\s*час", lesson_tail, maxsplit=1, flags=re.I)[0].strip(" .()")
                lesson_kind = lesson_kind or "Вид занятия уточнить по программе"
                # Название занятия в программе идёт после скобок с видом и длительностью.
                lesson_title = re.sub(r"^\s*\([^)]*\)\s*\.?\s*", "", lesson_tail).strip()
                lesson_title = remove_program_lesson_duration_phrase(lesson_title)
                lesson_title = re.sub(r"\s+", " ", lesson_title).strip(" .")
                if lesson_title:
                    lesson_title = re.split(r"(?<=[.!?])\s+", lesson_title, maxsplit=1)[0].strip()
                    if len(lesson_title) > 180:
                        lesson_title = lesson_title[:177].rsplit(" ", 1)[0] + "..."
                else:
                    lesson_title = "Название занятия не указано отдельно в программе"
                candidates.append({
                    "discipline": current_topic["section"],
                    "topic_number": current_topic["number"],
                    "topic": current_topic["title"] or "Название темы не указано в строке заголовка",
                    "lesson_number": match.group(1),
                    "lesson_type": lesson_kind,
                    "lesson_title": lesson_title,
                })
    if not candidates:
        return None
    # Remove duplicate paragraph matches while preserving distinct subsections.
    unique = []
    seen = set()
    for item in candidates:
        key = (item["discipline"], item["topic_number"], item["topic"], item["lesson_number"])
        if key not in seen:
            seen.add(key)
            unique.append(item)
    # Avoid silently selecting among duplicate numbering in different subsections.
    return unique[0] if len(unique) == 1 else {"multiple": unique}


def find_special_program_lessons(topic_number, lesson_number, category="special"):
    """Ищет все совпадения темы/занятия во всех подразделах выбранной подготовки."""
    if not PROGRAM_PATH.exists():
        raise FileNotFoundError(
            "Не найдена программа боевой подготовки. Положите файл программы рядом с bot.py."
        )

    doc = Document(str(PROGRAM_PATH))

    # Собираем текст документа в порядке следования: сначала обычные абзацы,
    # а также текст таблиц, если программа хранит часть сведений в таблицах.
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    topic_pattern = re.compile(r"^тема\s*№?\s*(\d+)\s*[.:]?\s*(.*)$", re.I)
    lesson_pattern = re.compile(r"^заняти(?:е|я)\s*№?\s*(\d+)\b\s*(.*)$", re.I | re.S)
    target = "technical" if category == "technical" else "special"

    active_category = None
    current_section = ""
    current_topic = None
    candidates = []

    for raw_line in paragraphs:
        if not raw_line:
            continue

        # Некоторые заголовки и подписи в Word содержат переносы строк.
        # Для составных абзацев берём первую строку: следующая часто содержит
        # служебную подпись «Перечень тем и расчёт часов».
        line = raw_line.splitlines()[0].strip()
        line = re.sub(r"\s+", " ", line).strip()
        normalized = normalize_text(line)
        is_topic = topic_pattern.match(line)
        is_lesson = lesson_pattern.match(line)

        upper_heading = (
            not is_topic and not is_lesson and len(line) < 140
            and line == line.upper() and any(ch.isalpha() for ch in line)
        )

        if upper_heading:
            heading = normalize_text(line)
            if "специальная подготовка" in heading:
                active_category = "special"
                current_section = line
                current_topic = None
                continue
            if "техническая подготовка" in heading:
                active_category = "technical"
                current_section = line
                current_topic = None
                continue
            # Любой следующий крупный раздел подготовки завершает поиск в текущем.
            if any(word in heading for word in ("подготовка", "вождение", "уставы")):
                active_category = None
                current_topic = None
                current_section = line
                continue

        if active_category != target:
            continue

        # Подразделы «Для ...» — это категории личного состава.
        # Меняем категорию, но НЕ сбрасываем тему: в программе название темы
        # иногда стоит перед заголовком «Для ...», а занятия идут после него.
        if (
            not is_topic and not is_lesson
            and re.match(r"^для\s+", normalized)
            and len(line) < 120
            and ":" not in line
        ):
            current_section = line.rstrip(" .:")
            continue

        # Служебные строки вроде «Перечень тем и расчёт часов» не являются
        # новыми категориями и не должны сбрасывать найденную тему.
        if is_topic:
            match = topic_pattern.match(line)
            current_topic = {
                "number": match.group(1),
                "title": match.group(2).strip(" .:—-"),
                "section": current_section or (
                    "СПЕЦИАЛЬНАЯ ПОДГОТОВКА" if target == "special"
                    else "ТЕХНИЧЕСКАЯ ПОДГОТОВКА"
                ),
            }
            continue

        if is_lesson and current_topic:
            match = lesson_pattern.match(line)
            if current_topic["number"] != str(topic_number) or match.group(1) != str(lesson_number):
                continue

            lesson_tail = match.group(2).strip()
            lesson_kind = re.split(
                r"[–—-]\s*\d|\s+\d+\s*час", lesson_tail,
                maxsplit=1, flags=re.I
            )[0].strip(" .()")
            lesson_kind = lesson_kind or "Вид занятия уточнить по программе"

            lesson_title = re.sub(r"^\s*\([^)]*\)\s*\.?\s*", "", lesson_tail).strip()
            lesson_title = remove_program_lesson_duration_phrase(lesson_title)
            lesson_title = re.sub(r"\s+", " ", lesson_title).strip(" .")
            if lesson_title:
                lesson_title = re.split(r"(?<=[.!?])\s+", lesson_title, maxsplit=1)[0].strip()
                if len(lesson_title) > 180:
                    lesson_title = lesson_title[:177].rsplit(" ", 1)[0] + "..."
            else:
                lesson_title = "Название занятия не указано отдельно в программе"

            candidates.append({
                "specialty": current_section,
                "topic_number": current_topic["number"],
                "topic": current_topic["title"] or "Название темы не указано в строке заголовка",
                "lesson_number": match.group(1),
                "lesson_type": lesson_kind,
                "lesson_title": lesson_title,
            })

    unique = []
    seen = set()
    for item in candidates:
        key = (
            normalize_text(item["specialty"]), item["topic_number"],
            normalize_text(item["topic"]), item["lesson_number"],
            normalize_text(item["lesson_title"]),
        )
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def program_match_label(index):
    return f"{BUTTON_PICK_PREFIX}{index}"


def program_match_keyboard(count):
    rows = [[program_match_label(i)] for i in range(1, count + 1)]
    rows.extend([[BUTTON_CANCEL], [BUTTON_BACK_TO_MENU]])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=False)


def parse_lookup_request(text):
    """Accepts e.g. 'Тактическая подготовка, тема 2, занятие 1'."""
    topic_match = re.search(r"тема\s*№?\s*(\d+)", text, re.I)
    lesson_match = re.search(r"занятие\s*№?\s*(\d+)", text, re.I)
    if not topic_match or not lesson_match:
        return None
    discipline = re.sub(r"[,;|]+", " ", text[:topic_match.start()]).strip(" .:-")
    discipline = re.sub(r"\b(по|подготовка по)\s*$", "", discipline, flags=re.I).strip()
    if not discipline:
        return None
    return discipline, topic_match.group(1), lesson_match.group(1)


def choice_keyboard(options, add_custom=False, add_skip=False):
    rows = [[item] for item in options]
    if add_custom:
        rows.append([BUTTON_CUSTOM])
    if add_skip:
        rows.append([BUTTON_SKIP])
    rows.append([BUTTON_CANCEL])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True, one_time_keyboard=False)

KEYBOARD = ReplyKeyboardMarkup(
    [
        [BUTTON_NEW],
        [BUTTON_BOOKS],
        [BUTTON_HELP],
    ],
    resize_keyboard=True,
)


def main_keyboard_for(user_id=None):
    rows = [[BUTTON_NEW], [BUTTON_HISTORY], [BUTTON_BOOKS], [BUTTON_REPORT_ERROR], [BUTTON_HELP]]
    if _is_admin(user_id):
        rows.append([BUTTON_OWNER_DASHBOARD])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


OWNER_PANEL_KEYBOARD = ReplyKeyboardMarkup(
    [
        [BUTTON_OWNER_GRANT, BUTTON_OWNER_REVOKE],
        [BUTTON_OWNER_USERS, BUTTON_OWNER_ERRORS],
        [BUTTON_OWNER_REPORTS],
        [BUTTON_OWNER_STATS],
        [BUTTON_OWNER_BACK],
    ],
    resize_keyboard=True,
)

# Все разделы из программы доступны через кнопки; специальные режимы оставлены отдельно.
_PREPARATION_CHOICES = [
    name for name in DISCIPLINES
    if name not in ("Подготовка по специальности", "Техническая подготовка", "Специальная подготовка")
]
_PREPARATION_ROWS = [[BUTTON_SPECIAL], [BUTTON_TECHNICAL]]
_PREPARATION_ROWS.extend([_PREPARATION_CHOICES[i:i + 2] for i in range(0, len(_PREPARATION_CHOICES), 2)])
_PREPARATION_ROWS.extend([[BUTTON_BACK_TO_MENU], [BUTTON_CANCEL]])
PREPARATION_KEYBOARD = ReplyKeyboardMarkup(_PREPARATION_ROWS, resize_keyboard=True)


BOOKS_KEYBOARD = ReplyKeyboardMarkup(
    [
        [BUTTON_MEDICAL_BOOKS],
        [BUTTON_LAW_BOOKS],
        [BUTTON_REGULATION_BOOKS],
        [BUTTON_DRILL_BOOKS],
        [BUTTON_PHYSICAL_BOOKS],
        [BUTTON_PSYCHOLOGY_BOOKS],
        [BUTTON_MATERIAL_BOOKS],
        [BUTTON_BACK_TO_MENU],
    ],
    resize_keyboard=True,
)

# Категории — логические метки, а не обязательные названия папок.
# Бот рекурсивно просматривает books/ и классифицирует книгу по имени/пути.
SAFE_BOOK_BUTTON_FOLDERS = {
    BUTTON_MEDICAL_BOOKS: "medical",
    BUTTON_LAW_BOOKS: "legal",
    BUTTON_REGULATION_BOOKS: "regulations",
    BUTTON_DRILL_BOOKS: "drill",
    BUTTON_PHYSICAL_BOOKS: "physical",
    BUTTON_PSYCHOLOGY_BOOKS: "psychological",
    BUTTON_MATERIAL_BOOKS: "material_support",
}
BOOKS_ROOT = Path(__file__).parent / "books"
GENERAL_BOOKS_DIR = BOOKS_ROOT / "general"  # обратная совместимость для старой структуры


def book_list_keyboard(book_names):
    rows = [[name] for name in book_names]
    rows.append([BUTTON_BACK_TO_MENU])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)



def history_keyboard(entries):
    rows = []
    mapping = {}
    for item in entries:
        label = f"📄 {item.get('created_at', '')[:10]} | #{item.get('id', '')[-5:]} | {item.get('discipline') or 'Конспект'} | {item.get('topic') or item.get('title') or 'Тема'}"
        label = label[:60]
        rows.append([label])
        mapping[label] = item.get("id")
    rows.extend([[BUTTON_BACK_TO_MENU], [BUTTON_CANCEL]])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True), mapping


def _history_item_by_id(user_id, entry_id):
    for item in _history_entries_for(user_id, limit=HISTORY_LIMIT_PER_USER):
        if item.get("id") == entry_id:
            path = (Path(__file__).parent / str(item.get("path", ""))).resolve()
            if path.is_relative_to(HISTORY_ROOT.resolve()) and path.is_file():
                return item, path
    return None, None


def _restricted_book_query(query):
    blocked = (
        "оруж", "стрельб", "стрелков", "тактическ", "артиллер", "наведен", "поражен", "боеприпас",
        "мина", "взрыв", "разведк", "огнев", "боевое примен", "прицел", "targeting", "weapon",
        "shooting", "firearm", "ammunition", "explosive", "tactical", "battle drill",
    )
    normalized = normalize_text(query)
    return any(normalize_text(term) in normalized for term in blocked)


def _search_safe_book_chunks(query, max_results=5):
    query = (query or "").strip()
    if len(query) < 3 or _restricted_book_query(query):
        return []
    terms = [w for w in re.findall(r"[а-яa-z0-9-]+", normalize_text(query)) if len(w) >= 3]
    if not terms:
        return []
    scored = []
    for path in list_safe_book_paths():
        try:
            st = path.stat()
            chunks = _read_book_chunks_cached(str(path), st.st_mtime_ns, st.st_size)
        except OSError:
            continue
        for label, chunk in chunks:
            norm_chunk = normalize_text(chunk)
            # Skip snippets that appear to include weapon, tactical, explosive, or targeting instructions.
            if _restricted_book_query(chunk):
                continue
            count = sum(1 for term in terms if term in norm_chunk)
            if count:
                pos = min((norm_chunk.find(term) for term in terms if term in norm_chunk), default=0)
                left = max(0, pos - 350)
                snippet = chunk[left:left + 1200].strip()
                scored.append((count, label, snippet))
    scored.sort(key=lambda item: (-item[0], item[1]))
    out, seen = [], set()
    for score, label, snippet in scored:
        key = (label, snippet[:160])
        if key in seen:
            continue
        seen.add(key)
        out.append((label, snippet))
        if len(out) >= max_results:
            break
    return out

def list_safe_book_paths(folder_name=None):
    """Ищет разрешённые PDF/DOCX/DJVU по всей books/, включая корень и любые подпапки."""
    if not BOOKS_ROOT.is_dir():
        return []
    allowed_extensions = {".pdf", ".docx", ".djvu"}
    paths = []
    for path in sorted(BOOKS_ROOT.rglob("*")):
        if not path.is_file() or path.suffix.casefold() not in allowed_extensions:
            continue
        if not _book_name_is_allowed(path):
            continue
        categories = _infer_safe_book_categories(path)
        # Показываем только документы, которые можно отнести к разрешённой категории.
        if categories and (folder_name is None or folder_name in categories):
            paths.append(path)
    return paths


def safe_book_display_name(path):
    try:
        relative = path.relative_to(BOOKS_ROOT).as_posix()
    except ValueError:
        relative = path.name
    # Кнопка Telegram должна оставаться читаемой.
    return relative[:180]
CANCEL_KEYBOARD = ReplyKeyboardMarkup(
    [[BUTTON_CANCEL]],
    resize_keyboard=True,
)


def set_paragraph_text(paragraph, new_text):
    """Меняет текст абзаца, стараясь сохранить его первое оформление."""
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def replace_in_paragraph(paragraph, replacements):
    """Заменяет текст, в том числе если Word разбил поле на фрагменты."""
    original = paragraph.text
    updated = original
    for old, new in replacements.items():
        updated = updated.replace(old, new)
    if updated != original:
        set_paragraph_text(paragraph, updated)


def process_paragraphs(paragraphs, replacements):
    for paragraph in paragraphs:
        replace_in_paragraph(paragraph, replacements)


def replace_approval_fields(doc, rank, full_name):
    replacements = {
        "{{ЗВАНИЕ_УТВЕРЖДАЮЩЕГО}}": rank,
        "{{ФИО_УТВЕРЖДАЮЩЕГО}}": full_name,
        "ЗВАНИЕ": rank,
        "ФИО": full_name,
        "майор": rank,
        "К.Серикпаев": full_name,
    }
    process_paragraphs(doc.paragraphs, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs, replacements)
    for section in doc.sections:
        for area in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
        ):
            process_paragraphs(area.paragraphs, replacements)


def fill_template(rank, full_name, output_path):
    """Сохраняет исходный шаблон и меняет только поля утверждения."""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            "Пошло что то не так."
        )
    doc = Document(str(TEMPLATE_PATH))
    replace_approval_fields(doc, rank, full_name)
    doc.save(output_path)


def extract_json(text):
    """Извлекает JSON, даже если модель добавила обрамляющий текст."""
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].strip()
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end < start:
        raise ValueError("Ошибочка.")
    return json.loads(text[start:end + 1])


def get_regulation_reference_text(discipline, topic, lesson_title, max_pages=6):
    """Извлекает из сборника уставов наиболее подходящие страницы для темы занятия."""
    normalized_discipline = normalize_text(discipline)
    if "общевоинские уставы" not in normalized_discipline and "строевая подготовка" not in normalized_discipline:
        return ""

    pdf_path = Path(__file__).parent / "books" / "Obshchevoinskie_ustavy_RK.pdf"
    if not pdf_path.exists():
        return "Источник не подключён: файл books/Obshchevoinskie_ustavy_RK.pdf не найден. Не утверждай, что использовал его текст."

    try:
        reader = PdfReader(str(pdf_path))
        page_texts = [(i + 1, (page.extract_text() or "").strip()) for i, page in enumerate(reader.pages)]
    except Exception as exc:
        print(f"Не удалось прочитать PDF с уставами: {exc}")
        return "Не удалось извлечь текст из PDF. Не утверждай, что использовал его текст."

    query = normalize_text(f"{topic} {lesson_title}")
    words = [w for w in re.findall(r"[а-яa-z0-9-]+", query) if len(w) >= 4]
    stop_words = {"тема", "занятие", "порядок", "основные", "общие", "вопросы", "изучение", "проведение", "подготовка", "устав", "устава", "уставе", "воинской", "общевоинские"}
    words = [w for w in words if w not in stop_words]

    if "строевая подготовка" in normalized_discipline:
        preferred = ["строевой устав", "строевая стойка", "строи", "строевые приемы", "строевое положение", "движение", "команды"]
    else:
        preferred = ["устав внутренней службы", "внутренней службы", "обязанности военнослужащих", "распорядок дня", "воинская дисциплина"]

    scored = []
    for page_num, raw in page_texts:
        text = normalize_text(raw)
        if not text:
            continue
        score = sum(2 for term in preferred if normalize_text(term) in text)
        score += sum(1 for word in words if word in text)
        if score:
            scored.append((score, page_num, raw))

    # Если точных совпадений нет, не подменяем источник случайными страницами.
    if not scored:
        return "В PDF не найдены страницы с терминами, соответствующими этой теме. Не выдумывай цитаты и уточни источник вручную."

    scored.sort(key=lambda item: (-item[0], item[1]))
    chosen = sorted(scored[:max_pages], key=lambda item: item[1])
    excerpts = []
    for _, page_num, raw in chosen:
        cleaned = re.sub(r"\n{3,}", "\n\n", raw)
        # Ограничиваем длину отдельной страницы, чтобы не перегружать запрос.
        excerpts.append(f"Страница PDF {page_num}:\n{cleaned[:2200]}")
    return "\n\n---\n\n".join(excerpts)



# Поиск источников включён только для общих, неприкладных учебных дисциплин.
# Поиск выполняется рекурсивно по books/, поэтому книги можно оставлять в корне
# или в существующих подпапках — вручную прописывать имена книг не требуется.
SAFE_LIBRARY_FOLDERS = {
    "военно-медицинская подготовка": "medical",
    "государственно-правовая подготовка": "legal",
    "общевоинские уставы": "regulations",
    "строевая подготовка": "drill",
    "физическая подготовка": "physical",
    "психологическая подготовка": "psychological",
}

# Не передаём в запрос Gemini источники, явно посвящённые применению оружия,
# стрельбе, боевому применению, разведке, тактическим действиям, минированию и т.п.
_BLOCKED_BOOK_NAME_TERMS = (
    "артиллер", "стрельб", "огнев", "оруж", "боевое применение",
    "боевой работе", "управление огнем", "управление огнём",
    "расчет стрельбы", "расчёт стрельбы", "снаряд", "ракета",
    "наведение", "поражение целей", "курс стрельб", "стрелков",
    "тактическ", "разведывательн", "минирован", "разминирован",
    "взрывчат", "противотанк", "противопехот", "мина", "минное поле",
    "огневая позиция", "боевой порядок", "топогеодез", "корректура стрельбы",
    "инженерная подготовка", "маскировка огневых", "радиационн", "химическ",
)

# Привязка по понятным фрагментам в имени файла ИЛИ имени родительской папки.
# Неизвестные документы автоматически не используются — это предотвращает
# случайное подключение книг неподходящей тематики из общей папки.
_SAFE_CATEGORY_MARKERS = {
    "medical": (
        "военно-медицин", "медицин", "medical_training", "medical",
        "первая помощь", "первая_помощь", "санитар",
    ),
    "legal": (
        "государственно-прав", "государственно_прав", "state_legal_training",
        "правовая подготовка", "правов", "закон", "законодательств", "правонаруш",
    ),
    "regulations": (
        "общевоинские_уставы", "общевоинские уставы", "obshchevoinskie_ustavy",
        "ustavy", "устав", "внутренней службы", "дисциплинарный устав",
        "гарнизонной службы",
    ),
    "drill": (
        "строевая", "строев", "drill_training", "строевой устав",
    ),
    "physical": (
        "физическая", "физичес", "physical_training", "physical",
        "нормативы физической", "физической подготовке", "спорт",
    ),
    "psychological": (
        "психолог", "psychological_preparation", "psychological",
        "морально-психолог", "работа с личным составом", "воспитательн",
    ),
    "material_support": (
        "material_support", "материальное обеспечение", "материально-техническ",
        "палатк", "эксплуатации и хранению", "хранение имущества",
        "вещевое имущество", "снабжение",
    ),
}


def _book_name_is_allowed(path: Path) -> bool:
    """Консервативный фильтр по полному относительному пути и имени файла."""
    try:
        name = normalize_text(path.relative_to(BOOKS_ROOT).as_posix())
    except ValueError:
        name = normalize_text(path.as_posix())
    return not any(normalize_text(term) in name for term in _BLOCKED_BOOK_NAME_TERMS)


def _infer_safe_book_categories(path: Path):
    """Возвращает безопасные категории по пути/имени; неизвестные книги пропускает."""
    try:
        searchable = normalize_text(path.relative_to(BOOKS_ROOT).as_posix())
    except ValueError:
        searchable = normalize_text(path.as_posix())
    return {
        category
        for category, markers in _SAFE_CATEGORY_MARKERS.items()
        if any(normalize_text(marker) in searchable for marker in markers)
    }


@lru_cache(maxsize=256)
def _read_book_chunks_cached(path_string: str, mtime_ns: int, file_size: int):
    """Читает разрешённый PDF/DOCX/DJVU и кэширует текст до изменения файла."""
    path = Path(path_string)
    suffix = path.suffix.casefold()
    chunks = []
    try:
        if suffix == ".pdf":
            reader = PdfReader(str(path))
            for page_number, page in enumerate(reader.pages, start=1):
                text = (page.extract_text() or "").strip()
                if text:
                    chunks.append((f"{path.name}, PDF стр. {page_number}", text[:9000]))
        elif suffix == ".docx":
            doc = Document(str(path))
            paragraph_number = 0
            for paragraph in doc.paragraphs:
                value = paragraph.text.strip()
                if value:
                    paragraph_number += 1
                    for part_number, start in enumerate(range(0, len(value), 1800), start=1):
                        label = f"{path.name}, абзац {paragraph_number}"
                        if len(value) > 1800:
                            label += f", фрагмент {part_number}"
                        chunks.append((label, value[start:start + 1800]))
            for table_number, table in enumerate(doc.tables, start=1):
                for row_number, row in enumerate(table.rows, start=1):
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        chunks.append((f"{path.name}, таблица {table_number}, строка {row_number}", row_text[:1800]))
        elif suffix == ".djvu":
            # DJVU читается, если на хостинге установлена утилита DjVuLibre.
            # Если её нет, конвертируйте книгу в PDF и загрузите PDF рядом с оригиналом.
            djvutxt = shutil.which("djvutxt")
            if not djvutxt:
                print(f"Пропуск DJVU {path.name}: на хостинге нет djvutxt; конвертируйте файл в PDF.")
                return tuple()
            result = subprocess.run(
                [djvutxt, str(path)], capture_output=True, text=True,
                encoding="utf-8", errors="replace", timeout=90, check=False,
            )
            if result.returncode != 0 or not result.stdout.strip():
                print(f"Не удалось извлечь текст из DJVU {path.name}: {result.stderr[:300]}")
                return tuple()
            pages = result.stdout.split("\f")
            for page_number, page_text in enumerate(pages, start=1):
                page_text = page_text.strip()
                if page_text:
                    chunks.append((f"{path.name}, DJVU стр. {page_number}", page_text[:9000]))
    except Exception as exc:
        print(f"Не удалось прочитать учебный источник {path.name}: {exc}")
    return tuple(chunks)


def get_safe_library_reference_text(discipline, topic, lesson_title, max_excerpts=5):
    """Ищет фрагменты по всей books/, но только в безопасной категории темы."""
    category = SAFE_LIBRARY_FOLDERS.get(normalize_text(discipline))
    if not category or not BOOKS_ROOT.is_dir():
        return ""

    query = normalize_text(f"{topic} {lesson_title}")
    query_words = [
        word for word in re.findall(r"[а-яa-z0-9-]+", query)
        if len(word) >= 4 and word not in {
            "тема", "занятие", "порядок", "основные", "общие", "вопросы",
            "изучение", "проведение", "подготовка", "основных", "воинской",
        }
    ]
    if not query_words:
        return ""

    scored = []
    allowed_extensions = {".pdf", ".docx", ".djvu"}
    for path in sorted(BOOKS_ROOT.rglob("*")):
        if not path.is_file() or path.suffix.casefold() not in allowed_extensions:
            continue
        if not _book_name_is_allowed(path):
            continue
        if category not in _infer_safe_book_categories(path):
            continue
        try:
            stat = path.stat()
            chunks = _read_book_chunks_cached(str(path), stat.st_mtime_ns, stat.st_size)
        except OSError:
            continue
        for label, chunk_text in chunks:
            normalized_chunk = normalize_text(chunk_text)
            score = sum(2 for word in query_words if word in normalized_chunk)
            if score:
                scored.append((score, label, chunk_text))

    if not scored:
        return ""

    scored.sort(key=lambda item: (-item[0], item[1]))
    excerpts = []
    seen = set()
    total_chars = 0
    for _, label, chunk_text in scored:
        key = (label, chunk_text[:300])
        if key in seen:
            continue
        seen.add(key)
        excerpt = f"Источник: {label}\n{chunk_text.strip()}"
        if total_chars + len(excerpt) > 9500:
            break
        excerpts.append(excerpt)
        total_chars += len(excerpt)
        if len(excerpts) >= max_excerpts:
            break
    return "\n\n---\n\n".join(excerpts)

def generate_lesson_data(discipline, topic, lesson_title, lesson_type, audience, duration, goals, manuals, materials):
    """Сначала формирует 2–3 учебных вопроса, затем готовит раздел по каждому из них."""
    regulation_reference = get_regulation_reference_text(discipline, topic, lesson_title)
    library_reference = get_safe_library_reference_text(discipline, topic, lesson_title)
    reference_parts = [part for part in (regulation_reference, library_reference) if part]
    combined_reference = "\n\n=== ФРАГМЕНТЫ ИСТОЧНИКОВ ===\n\n".join(reference_parts)
    prompt = f"""
Подготовь на русском языке учебный план-конспект для занятия.
Нужно заполнить имеющийся шаблон Word, а не создавать новую структуру документа.
Пиши профессионально, ясно, без Markdown-разделителей и без выдуманных ссылок на приказы.
Не выдумывай названия конкретных нормативных документов. Если данных не хватает,
укажи, что их следует уточнить по утверждённой программе.

Содержание должно быть учебным и общим. Не включай реальные координаты, планы операций,
точные инструкции по применению оружия, наведению/поражению целей, обходу защиты или
иные подробности, которые могут непосредственно помочь причинить вред.
Для технических и тактических тем делай акцент на терминах, назначении, безопасности,
организации учебного процесса и проверке знаний.

Данные занятия:
Дисциплина: {discipline}
Тема: {topic}
Название занятия из программы боевой подготовки (используй дословно, не придумывай и не заменяй): {lesson_title}
Категория обучаемых: {audience}
Продолжительность: {duration}
Руководства и пособия (предоставлены пользователем): {manuals}
Материальное обеспечение (предоставлено пользователем): {materials}
Дополнительные пожелания: {goals or "Не указаны"}

Текст из подключённых учебных источников (если подходящие фрагменты найдены):
{combined_reference or "Для этой дисциплины подходящие фрагменты из разрешённой библиотеки не найдены."}
Используй только переданные фрагменты как подтверждённую основу. Не выдумывай пункты, формулировки, страницы, названия документов или номера статей. Если фрагментов недостаточно, прямо отметь, какие сведения необходимо уточнить по утверждённому источнику. Если в фрагменте есть метка «Источник», сохрани название документа и указанную страницу/место в разделе источников конспекта.

Верни только корректный JSON со следующими ключами:
{{
  "objectives": ["цель 1", "цель 2", "цель 3"],
  "questions": ["Учебный вопрос 1", "Учебный вопрос 2", "Учебный вопрос 3"],
  "question_sections": [
    {{"question": "точный текст учебного вопроса 1 из массива questions", "content": "развёрнутый материал только по учебному вопросу 1"}},
    {{"question": "точный текст учебного вопроса 2 из массива questions", "content": "развёрнутый материал только по учебному вопросу 2"}},
    {{"question": "точный текст учебного вопроса 3 из массива questions", "content": "развёрнутый материал только по учебному вопросу 3"}}
  ],
  "materials": ["материальное обеспечение 1", "материальное обеспечение 2"],
  "manuals": "руководства и учебные пособия; не выдумывай конкретные названия",
  "intro_minutes": 10,
  "main_minutes": 30,
  "conclusion_minutes": 5,
  "intro_leader": "действия руководителя во вводной части",
  "intro_learners": "действия обучаемых во вводной части",
  "main_learners": "действия обучаемых в основной части по каждому учебному вопросу",
  "conclusion_leader": "подведение итогов и задание на самостоятельную подготовку",
  "conclusion_learners": "действия обучаемых в заключительной части",
  "review_questions": ["вопрос для закрепления 1", "вопрос 2", "вопрос 3"],
  "self_study": ["задание 1", "задание 2"],
  "diagrams": [
    {{"question_index": 1, "title": "Краткое название схемы", "style": "flow", "nodes": ["Понятие", "Основные элементы", "Проверка понимания"]}}
  ],
  "illustrations": [
    {{"question_index": 1, "caption": "Что показано на фотографии", "search_query": "neutral educational photo query"}}
  ]
}}

Порядок построения конспекта — обязателен:
1. Сначала внимательно проанализируй точное название занятия и переданные фрагменты учебных источников.
2. На этой основе сформулируй 2 или 3 логически последовательных учебных вопроса. Используй 2 вопроса, если тема узкая, и 3, если название занятия действительно охватывает несколько самостоятельных аспектов. Вопросы должны раскрывать тему занятия по частям, а не повторять название другими словами. Иди от основных понятий к следующим связанным аспектам. Не выходи за рамки названия занятия и источников.
3. Затем напиши отдельный раздел question_sections для КАЖДОГО вопроса в том же порядке. Поле question в каждом элементе должно дословно совпадать с соответствующим элементом массива questions. Поле content должно раскрывать именно этот вопрос, а не соседние вопросы. Не создавай дополнительный вопрос и не меняй формулировки после формирования массива questions.
4. Код бота вставит вопрос как заголовок соответствующего раздела конспекта и разместит под ним именно content из соответствующего элемента question_sections. Поэтому число элементов questions и question_sections должно совпадать.

Требования:
- objectives: ровно 3 короткие учебные цели.
- questions: ровно 2 или 3 учебных вопроса.
- question_sections: столько же элементов, сколько вопросов; каждый раздел строго соответствует вопросу с тем же индексом.
- Время вводной, основной и заключительной частей в сумме должно примерно соответствовать
  указанной продолжительности. Если длительность указана в часах, переведи её в минуты.
- Каждый content в question_sections — это основная часть графы «Краткое содержание (действия руководителя)» для соответствующего учебного вопроса. Не ограничивайся кратким пересказом: подробно раскрой вопрос.
  Для каждого вопроса дай несколько содержательных абзацев:
  определения и основные понятия; назначение и значение; порядок организации занятия;
  важные признаки/этапы/требования; типичные ошибки и меры безопасности; что руководитель
  объясняет, показывает, уточняет и проверяет у обучаемых. Добавляй примеры учебных ситуаций,
  контрольные вопросы и логичные переходы между вопросами, если это подходит теме.
  Ориентир — не менее 150–250 слов на каждый учебный вопрос, а для сложной темы больше.
  Общий объём основной части обычно 800–1400 слов для занятия продолжительностью 90–180 минут;
  не заполняй объём повторениями и не выдумывай нормативы или названия документов.
- intro_leader: подробный связный текст примерно 100–160 слов с последовательностью действий
  руководителя во вводной части.
- main_learners: конкретные действия обучаемых по каждому учебному вопросу, а не одна общая фраза.
- conclusion_leader: развёрнутое подведение итогов, проверка достижения целей, разбор типичных
  ошибок, вопросы для закрепления и самостоятельная подготовка.
- diagrams: создай 1–3 разнообразные учебные схемы только там, где они помогают объяснить материал. Для каждой укажи question_index, title, style и nodes (3–5 коротких подписей, до 7 слов каждая). style выбирай из flow (последовательность), branches (центральное понятие и ветви), compare (сравнение двух категорий), cycle (цикл/повторяющийся процесс). Не повторяй один стиль во всех схемах одного документа; выбирай его по смыслу темы. Схемы должны быть визуально понятными и не дублировать один и тот же шаблон. Не включай координаты, расчёты наведения, процедуры применения оружия или иные оперативные инструкции. Если схема неуместна, верни пустой список.
- illustrations: предложи 1–3 чёрно-белые учебные рисунки в стиле линейных иллюстраций для безопасных неприкладных тем, только когда они помогают понять учебный вопрос. Для каждой укажи question_index (номер из questions), короткую подпись caption и краткий search_query на английском или русском для поиска изображения с открытой лицензией. Для строевой подготовки допустимы общие иллюстрации строевой стойки, поворотов и строевого шага; для первой помощи — не графичные учебные иллюстрации. Не запрашивай и не используй изображения, показывающие применение оружия, стрельбу, наведение, поражение целей, боевые действия или тактические процедуры. Если подходящих безопасных изображений нет, верни пустой список. Код бота может не найти изображение, и тогда конспект всё равно должен быть создан.
- Не используй смайлики, эмодзи или декоративные символы в любом тексте конспекта.
- Не добавляй выдуманные источники, страницы, нормативы или номера статей. Код бота сам добавит в конец основной части только те метки источников, которые реально есть во входных фрагментах.
- Не включай фразу «Дежурный докладывает о готовности группы» в действия обучаемых.
- Не добавляй раздел или поле «Метод».
- Верни только JSON, без тройных кавычек и пояснений.
"""
    # Используем только Google Gemini. Временные ошибки сервера (503/500/502/504)
    # повторяем с увеличивающейся паузой, не повторяя ошибки ключа/модели.
    response = None
    last_error = None
    for attempt, delay in enumerate((0, 3, 7), start=1):
        if delay:
            print(f"Gemini временно недоступен; повторная попытка {attempt}/3 через {delay} сек.")
            time.sleep(delay)
        try:
            response = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=prompt,
            )
            last_error = None
            break
        except Exception as error:
            last_error = error
            error_text = str(error).upper()
            if "429" in error_text or "RESOURCE_EXHAUSTED" in error_text:
                raise RuntimeError(
                    "Подождите, Либо напишите в Поддержку "
                ) from error
            temporary = any(marker in error_text for marker in ("503", "UNAVAILABLE", "500", "502", "504", "INTERNAL SERVER ERROR"))
            if not temporary or attempt == 3:
                if temporary:
                    raise RuntimeError(
                        "Нажмите «Повторить создание» немного позже."
                    ) from error
                raise

    if response is None:
        raise RuntimeError("Gemini не вернул ответ после повторных попыток.") from last_error
    response_text = response.text
    if not response_text:
        raise RuntimeError("Gemini вернул пустой ответ. Попробуйте повторить создание.")
    provider_used = "Gemini"

    data = extract_json(response_text)
    print(f"Конспект сгенерирован через {provider_used}.")

    # Gemini может иногда опускать одно из вспомогательных полей JSON,
    # хотя основной материал (учебные вопросы и их разделы) присутствует.
    # Заполняем безопасные шаблонные поля по умолчанию вместо срыва всего конспекта.
    # Ключевые содержательные поля questions/question_sections по-прежнему проверяются ниже.
    fallback_fields = {
        "intro_minutes": 10,
        "main_minutes": 30,
        "conclusion_minutes": 5,
        "intro_leader": (
            "Объявляю тему и цели занятия, довожу порядок проведения, "
            "проверяю готовность обучаемых и наличие учебных материалов."
        ),
        "intro_learners": (
            "Слушают руководителя, уясняют тему и цели занятия, "
            "готовятся к изучению учебного материала."
        ),
        "main_learners": (
            "Воспринимают учебный материал, записывают основные положения, "
            "отвечают на вопросы и выполняют предусмотренные учебные задания."
        ),
        "conclusion_leader": (
            "Подвожу итоги занятия, проверяю усвоение материала контрольными "
            "вопросами и довожу задание на самостоятельную подготовку."
        ),
        "conclusion_learners": (
            "Отвечают на контрольные вопросы, уточняют непонятные положения "
            "и записывают задание на самостоятельную подготовку."
        ),
        "review_questions": [],
        "self_study": [],
        "diagrams": [],
        "illustrations": [],
        "materials": [],
        "manuals": manuals or "Не указано",
    }
    for key, fallback_value in fallback_fields.items():
        if key not in data or data[key] is None or data[key] == "":
            data[key] = fallback_value

    required = ["objectives", "questions", "question_sections"]
    missing = [key for key in required if key not in data]
    if missing:
        raise ValueError(
            f"Модель {provider_used} не вернула обязательные поля: " + ", ".join(missing)
        )

    # Закрепляем зависимость «учебный вопрос -> его материал»: заголовок каждого
    # раздела берётся из questions, а соответствующее содержание — из question_sections
    # с тем же индексом. Это не даёт модели слегка изменить формулировку заголовка.
    raw_questions = data.get("questions")
    raw_sections = data.get("question_sections")
    if not isinstance(raw_questions, list) or not isinstance(raw_sections, list):
        raise ValueError("Gemini вернул неправильный формат учебных вопросов. Нажмите «Повторить создание».")
    questions = [strip_emojis(str(q)).strip() for q in raw_questions if str(q).strip()]
    questions = questions[:3]
    if len(questions) < 2:
        raise ValueError("Gemini должен сформулировать 2–3 учебных вопроса. Нажмите «Повторить создание».")
    if len(raw_sections) < len(questions):
        raise ValueError("Gemini не подготовил материал для каждого учебного вопроса. Нажмите «Повторить создание».")

    lesson_sections = []
    for index, question in enumerate(questions):
        section = raw_sections[index]
        if not isinstance(section, dict):
            raise ValueError("Неправильный формат раздела учебного вопроса. Нажмите «Повторить создание».")
        content = strip_emojis(str(section.get("content", ""))).strip()
        if not content:
            raise ValueError(f"Нет материала по учебному вопросу {index + 1}. Нажмите «Повторить создание».")
        lesson_sections.append(f"Учебный вопрос {index + 1}. {question}\n\n{content}")

    # Формируем основную часть сами, используя вопросы в каноническом порядке.
    main_leader = "\n\n".join(lesson_sections)
    # Перечисляем только источники, метки которых реально передавались модели.
    source_labels = []
    for label in re.findall(r"Источник:\s*([^\n]+)", combined_reference):
        label = label.strip()
        if label and label not in source_labels:
            source_labels.append(label)
    if source_labels:
        main_leader += "\n\nИспользованные источники:\n" + "\n".join(f"- {label}" for label in source_labels)
    data["questions"] = questions
    data["main_leader"] = main_leader
    data["manuals"] = manuals or "Не указано"
    data["materials"] = [item.strip() for item in materials.splitlines() if item.strip()] or ["Не указано"]
    if not isinstance(data.get("diagrams"), list):
        data["diagrams"] = []
    # Проверяем структуру схем до передачи в генератор изображений.
    safe_diagrams = []
    for diagram in data["diagrams"][:3]:
        if not isinstance(diagram, dict):
            continue
        nodes = diagram.get("nodes", [])
        if not isinstance(nodes, list):
            continue
        nodes = [strip_emojis(str(node)).strip()[:90] for node in nodes if str(node).strip()][:5]
        if len(nodes) < 3:
            continue
        safe_diagrams.append({
            "question_index": max(1, min(3, int(diagram.get("question_index", 1) or 1))),
            "title": strip_emojis(str(diagram.get("title", "Учебная схема"))).strip()[:120] or "Учебная схема",
            "style": str(diagram.get("style", "flow")).lower() if str(diagram.get("style", "flow")).lower() in {"flow", "branches", "compare", "cycle"} else "flow",
            "nodes": nodes,
        })
    data["diagrams"] = safe_diagrams

    # Иллюстрации — необязательное поле: их отсутствие не должно ломать конспект.
    safe_illustrations = []
    raw_illustrations = data.get("illustrations", [])
    if isinstance(raw_illustrations, list):
        for illustration in raw_illustrations[:3]:
            if not isinstance(illustration, dict):
                continue
            try:
                question_index = int(illustration.get("question_index", 1) or 1)
            except (TypeError, ValueError):
                question_index = 1
            caption = strip_emojis(str(illustration.get("caption", "Иллюстрация к учебному вопросу"))).strip()[:160]
            search_query = strip_emojis(str(illustration.get("search_query", ""))).strip()[:180]
            if search_query:
                safe_illustrations.append({
                    "question_index": max(1, min(len(questions), question_index)),
                    "caption": caption or "Иллюстрация к учебному вопросу",
                    "search_query": search_query,
                })
    data["illustrations"] = safe_illustrations
    for key in ("intro_learners", "main_learners", "conclusion_learners"):
        data[key] = str(data.get(key, "")).replace("Дежурный докладывает о готовности группы.", "").replace("Дежурный докладывает о готовности группы", "").strip()
    return data


def find_paragraph(doc, starts_with):
    for paragraph in doc.paragraphs:
        if paragraph.text.strip().startswith(starts_with):
            return paragraph
    return None


def set_if_found(doc, starts_with, new_text):
    paragraph = find_paragraph(doc, starts_with)
    if paragraph is not None:
        set_paragraph_text(paragraph, new_text)


def set_cell(cell, text):
    """Заполняет существующую ячейку таблицы, сохраняя таблицу шаблона."""
    cell.text = str(text or "")


def format_discipline(discipline):
    """Преобразует распространённые названия дисциплин в форму для строки «по ...»."""
    value = (discipline or "").strip()
    forms = {
        "тактическая подготовка": "тактической подготовке",
        "техническая подготовка": "технической подготовке",
        "строевая подготовка": "строевой подготовке",
        "огневая подготовка": "огневой подготовке",
        "разведывательная подготовка": "разведывательной подготовке",
        "физическая подготовка": "физической подготовке",
        "инженерная подготовка": "инженерной подготовке",
        "радиационная, химическая и биологическая защита": "радиационной, химической и биологической защите",
    }
    return forms.get(value.casefold(), value)


def strip_emojis(text):
    """Удаляет emoji и декоративные пиктограммы из текста конспекта."""
    import re
    value = str(text or "")
    emoji_pattern = re.compile(
        "[\U0001F1E6-\U0001F1FF\U0001F300-\U0001FAFF"
        "\U00002600-\U000027BF\U0000FE0F\U0000200D]"
    )
    value = emoji_pattern.sub("", value)
    # Убираем декоративные маркеры, оставляя обычную пунктуацию и нумерацию.
    value = value.replace("📗", "").replace("📘", "").replace("📙", "")
    value = value.replace("📋", "").replace("🖥", "").replace("🧰", "").replace("📚", "")
    return value


def set_cell_text(cell, text):
    """Записывает многострочный текст в ячейку таблицы без emoji."""
    cell.text = strip_emojis(text)


def _diagram_font(size=28, bold=False):
    """Подбирает доступный шрифт для схемы; работает на Windows и Linux."""
    candidates = []
    if os.name == "nt":
        candidates.append(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf")
    candidates.extend([
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ])
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size=size)
        except OSError:
            pass
    return ImageFont.load_default()


def create_diagram_image(diagram, image_path):
    """Рисует аккуратные учебные схемы разных типов."""
    nodes = [str(x).strip()[:90] for x in diagram.get("nodes", []) if str(x).strip()][:5]
    if len(nodes) < 3:
        return False
    style = diagram.get("style", "flow")
    width, margin = 1400, 70
    title = str(diagram.get("title", "Учебная схема"))[:100]
    title_font = _diagram_font(38, True)
    node_font = _diagram_font(28, True)
    small_font = _diagram_font(24)
    box_w, box_h, gap = 1120, 118, 38
    if style == "branches":
        width, box_w, box_h, gap = 1400, 510, 118, 42
        height = 300 + ((len(nodes) + 1) // 2) * (box_h + gap)
    elif style == "compare":
        width, box_w, box_h, gap = 1400, 560, 142, 45
        height = 320 + ((len(nodes) + 1) // 2) * (box_h + gap)
    elif style == "cycle":
        width, box_w, box_h, gap = 1400, 700, 118, 38
        height = 300 + len(nodes) * (box_h + gap)
    else:
        height = 170 + len(nodes) * box_h + (len(nodes) - 1) * gap
    image = Image.new("RGB", (width, height), "#f7f9fc")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((22, 18, width-22, height-18), radius=28, fill="#ffffff", outline="#d6dee9", width=3)
    draw.text((width//2, 52), title, fill="#17324d", font=title_font, anchor="ma")

    def draw_box(cx, cy, text, number=None, fill="#eaf2fb"):
        x0, y0 = cx-box_w//2, cy-box_h//2
        draw.rounded_rectangle((x0, y0, x0+box_w, y0+box_h), radius=24, fill=fill, outline="#47739b", width=4)
        label = (f"{number}. " if number is not None else "") + text
        words, lines, line = label.split(), [], ""
        maxw = box_w - 70
        for word in words:
            candidate = (line + " " + word).strip()
            if draw.textbbox((0,0), candidate, font=node_font)[2] > maxw and line:
                lines.append(line); line = word
            else: line = candidate
        if line: lines.append(line)
        lines = lines[:3]
        heights = [draw.textbbox((0,0), t, font=node_font)[3] for t in lines]
        y = cy - (sum(heights)+12*(len(lines)-1))//2
        for line_text, line_h in zip(lines, heights):
            draw.text((cx, y), line_text, fill="#183047", font=node_font, anchor="ma")
            y += line_h + 12
        return (x0, y0, x0+box_w, y0+box_h)

    if style == "branches":
        center = (width//2, 215)
        root = draw_box(*center, nodes[0], 1, "#dceaf8")
        child_count = len(nodes)-1
        cols = 2
        positions=[]
        for i, node in enumerate(nodes[1:]):
            row, col = divmod(i, cols)
            cx = width//4 if col == 0 else 3*width//4
            cy = 360 + row*(box_h+gap)
            positions.append((cx,cy,node,i+2))
            draw.line((center[0], root[3], cx, cy-box_h//2), fill="#6c8ba8", width=5)
        for cx,cy,node,num in positions: draw_box(cx,cy,node,num,"#edf5ee")
    elif style == "compare":
        leftx, rightx = width//4, 3*width//4
        draw.text((leftx, 145), "Вариант А", fill="#315d83", font=small_font, anchor="ma")
        draw.text((rightx, 145), "Вариант Б", fill="#315d83", font=small_font, anchor="ma")
        for i,node in enumerate(nodes):
            cx = leftx if i%2==0 else rightx
            cy = 220 + (i//2)*(box_h+gap)
            draw_box(cx,cy,node,i+1,"#eaf2fb" if i%2==0 else "#edf5ee")
        draw.line((width//2, 170, width//2, height-55), fill="#d4dce5", width=3)
    elif style == "cycle":
        centers=[]
        cy=210
        for i,node in enumerate(nodes):
            centers.append((width//2,cy+ i*(box_h+gap)))
        for i,(cx,cy) in enumerate(centers):
            draw_box(cx,cy,nodes[i],i+1,"#eaf2fb" if i%2==0 else "#edf5ee")
            if i < len(centers)-1:
                draw.line((cx,cy+box_h//2+4,cx,centers[i+1][1]-box_h//2-10), fill="#6c8ba8", width=5)
                draw.polygon([(cx-12,centers[i+1][1]-box_h//2-20),(cx+12,centers[i+1][1]-box_h//2-20),(cx,centers[i+1][1]-box_h//2-4)],fill="#6c8ba8")
        # Return arrow visually marks the cycle.
        draw.arc((width//2+box_w//2-20, 190, width-65, height-55), 270, 90, fill="#bd7c32", width=5)
        draw.polygon([(width-75,height//2-8),(width-48,height//2),(width-75,height//2+8)],fill="#bd7c32")
    else:
        centers=[]
        for i,node in enumerate(nodes):
            cy=145 + box_h//2 + i*(box_h+gap)
            centers.append((width//2,cy))
            draw_box(width//2,cy,node,i+1,"#eaf2fb" if i%2==0 else "#edf5ee")
        for i in range(len(centers)-1):
            cx=width//2; y1=centers[i][1]+box_h//2+3; y2=centers[i+1][1]-box_h//2-8
            draw.line((cx,y1,cx,y2),fill="#6c8ba8",width=5)
            draw.polygon([(cx-12,y2-14),(cx+12,y2-14),(cx,y2+2)],fill="#6c8ba8")
    image.save(image_path, format="PNG", optimize=True)
    return True

def append_diagrams_to_cell(cell, diagrams, image_dir):
    """Добавляет схемы прямо в нужную ячейку таблицы Word, а не после таблицы."""
    if not diagrams:
        return
    from docx.shared import Inches
    for index, diagram in enumerate(diagrams, start=1):
        image_path = Path(image_dir) / f"lesson_diagram_{index}.png"
        if not create_diagram_image(diagram, image_path):
            continue
        caption = cell.add_paragraph()
        caption.paragraph_format.space_before = Inches(0.08)
        caption.paragraph_format.space_after = Inches(0.04)
        caption.add_run(f"Схема к учебному вопросу {diagram.get('question_index', index)}: {diagram.get('title', 'Учебная схема')}").bold = True
        picture_paragraph = cell.add_paragraph()
        picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture_paragraph.add_run().add_picture(str(image_path), width=Inches(4.5))


# Разрешённые дисциплины для автоматического поиска иллюстраций. Для тем с оружием,
# стрельбой, артиллерийскими расчётами или боевыми действиями изображения не ищутся.
PHOTO_SAFE_DISCIPLINES = (
    "строевая подготовка",
    "военно-медицинская подготовка",
    "физическая подготовка",
    "государственно-правовая подготовка",
    "психологическая подготовка",
)
PHOTO_BLOCK_TERMS = (
    "артиллер", "тактическ", "огнев", "стрел", "стрельб", "оруж", "боеприпас",
    "боев", "пораж", "навед", "прицел", "мина", "взрыв", "разведк",
    "наступлен", "оборон", "огневой позиции", "расчет установок", "цель",
    "artillery", "weapon", "rifle", "firearm", "shooting", "targeting", "aiming",
    "combat", "battle", "tactical", "attack", "defense", "munition", "ammunition",
    "explosive", "landmine", "gunfire", "sniper", "bombing", "shelling",
)
PHOTO_IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
IMAGE_LIBRARY_DIR = Path(__file__).parent / "images"
IMAGE_CACHE_DIR = IMAGE_LIBRARY_DIR / ".commons_cache_bw_public_domain"


def _photo_topic_allowed(discipline, topic, lesson_title, query=""):
    """Allow image lookup only for explicitly non-operational lesson categories."""
    combined = normalize_text(" ".join((discipline or "", topic or "", lesson_title or "", query or "")))
    if not any(normalize_text(name) in normalize_text(discipline) for name in PHOTO_SAFE_DISCIPLINES):
        return False
    return not any(term in combined for term in PHOTO_BLOCK_TERMS)


def _clean_html_text(value):
    value = html.unescape(re.sub(r"<[^>]+>", " ", str(value or "")))
    return re.sub(r"\s+", " ", value).strip()


def _find_local_lesson_images(discipline, topic, lesson_title, query, limit=2):
    """Find user-managed pictures by matching filename/folder words in images/."""
    if not IMAGE_LIBRARY_DIR.exists():
        return []
    combined = normalize_text(" ".join((discipline, topic, lesson_title, query)))
    words = set(re.findall(r"[a-zа-яё0-9]{4,}", combined))
    stop = {"подготовка", "занятие", "учебный", "вопрос", "порядок", "основные", "требования", "изучение", "военно", "военная"}
    words -= stop
    # Частые русские названия также встречаются в именах файлов латиницей.
    aliases = {
        "строевая": {"drill", "stroevaya"},
        "стойк": {"stoyka", "attention"},
        "поворот": {"povorot", "turn"},
        "шаг": {"shag", "march", "marching"},
        "медицин": {"medical", "firstaid", "first_aid", "bandage"},
        "физическ": {"fitness", "exercise", "physical"},
        "психолог": {"psychology", "teamwork"},
    }
    for russian_prefix, equivalents in aliases.items():
        if russian_prefix in combined:
            words.update(equivalents)
    scored = []
    for path in IMAGE_LIBRARY_DIR.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in PHOTO_IMAGE_SUFFIXES or ".commons_cache" in path.parts:
            continue
        candidate = normalize_text(" ".join(path.relative_to(IMAGE_LIBRARY_DIR).parts))
        if any(term in candidate for term in PHOTO_BLOCK_TERMS):
            continue
        score = sum(1 for word in words if word in candidate)
        if score:
            scored.append((score, str(path).lower(), path))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [{"path": path, "caption": path.stem.replace("_", " ").replace("-", " "), "credit": "Локальная иллюстрация из папки images/"} for _, _, path in scored[:limit]]


def _fetch_commons_image(query):
    """Fetch one freely licensed Commons thumbnail; failures are non-fatal."""
    if not query:
        return None
    try:
        IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_key = hashlib.sha256(query.encode("utf-8")).hexdigest()[:20]
        cached_img = IMAGE_CACHE_DIR / f"{cache_key}.jpg"
        cached_meta = IMAGE_CACHE_DIR / f"{cache_key}.txt"
        if cached_img.exists():
            credit = cached_meta.read_text(encoding="utf-8") if cached_meta.exists() else "Источник: Wikimedia Commons"
            return {"path": cached_img, "caption": query, "credit": credit}

        params = {
            "action": "query",
            "generator": "search",
            "gsrsearch": f'filetype:bitmap "line drawing" black white {query}',
            "gsrnamespace": 6,
            "gsrlimit": 8,
            "prop": "imageinfo",
            "iiprop": "url|extmetadata|mime",
            "iiurlwidth": 1000,
            "format": "json",
        }
        api_url = "https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(api_url, headers={"User-Agent": "VoennyKonspektEducationalBot/1.0 (image attribution included)"})
        with urllib.request.urlopen(req, timeout=7) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
        pages = (payload.get("query") or {}).get("pages") or {}
        candidates = list(pages.values())
        for page in candidates:
            imageinfo = (page.get("imageinfo") or [{}])[0]
            ext = imageinfo.get("extmetadata") or {}
            license_name = _clean_html_text((ext.get("LicenseShortName") or {}).get("value", ""))
            license_lower = license_name.casefold()
            if not any(key in license_lower for key in ("cc0", "public domain", "public-domain")) or "cc by" in license_lower:
                continue
            image_url = imageinfo.get("thumburl") or imageinfo.get("url")
            if not image_url or not image_url.startswith("https://"):
                continue
            image_req = urllib.request.Request(image_url, headers={"User-Agent": "VoennyKonspektEducationalBot/1.0"})
            with urllib.request.urlopen(image_req, timeout=8) as image_response:
                raw = image_response.read(12_000_000)
            # Normalize format to JPEG so python-docx can embed reliably.
            with Image.open(io.BytesIO(raw)) as image:
                from PIL import ImageOps
                image.thumbnail((1400, 1200))
                image = ImageOps.grayscale(image)
                image = ImageOps.autocontrast(image, cutoff=2)
                # High-contrast black/white raster for a manual-like educational appearance.
                image = image.point(lambda value: 255 if value >= 190 else 0, mode="1").convert("L")
                image.save(cached_img, format="JPEG", quality=92, optimize=True)
            source_url = imageinfo.get("descriptionurl", "https://commons.wikimedia.org/")
            author = _clean_html_text((ext.get("Artist") or {}).get("value", "")) or "автор не указан"
            credit = ""  # CC0/public-domain only; no source line inside the lesson document.
            cached_meta.write_text(credit, encoding="utf-8")
            return {"path": cached_img, "caption": _clean_html_text(page.get("title", "" )).replace("File:", ""), "credit": credit}
    except Exception as exc:
        print(f"Не удалось получить иллюстрацию из Wikimedia Commons ({query!r}): {exc}")
    return None


def _fallback_photo_queries(discipline, topic, lesson_title, questions):
    """Conservative search terms used only for allowlisted non-operational subjects."""
    combined = normalize_text(f"{topic} {lesson_title}")
    if "строевая подготовка" in normalize_text(discipline):
        if any(word in combined for word in ("стойк", "смирно", "положение")):
            return [(1, "Строевая стойка", "black white line drawing soldier standing at attention front side view")]
        if "поворот" in combined:
            return [(1, "Повороты на месте", "black white line drawing military drill turning in place soldier")]
        if any(word in combined for word in ("шаг", "движение", "марш")):
            return [(1, "Движение строевым шагом", "black white line drawing soldier marching drill step")]
        return [(1, "Строевая подготовка", "black white line drawing military drill soldier standing at attention")]
    if "военно-медицинская подготовка" in normalize_text(discipline):
        return [(1, "Учебная иллюстрация по первой помощи", "first aid training demonstration non graphic")]
    if "физическая подготовка" in normalize_text(discipline):
        return [(1, "Учебная иллюстрация по физической подготовке", "physical fitness training exercise demonstration")]
    if "психологическая подготовка" in normalize_text(discipline):
        return [(1, "Психологическая подготовка", "group communication training teamwork workshop")]
    if "государственно-правовая подготовка" in normalize_text(discipline):
        return [(1, "Учебная иллюстрация", "classroom legal education students discussion")]
    return []


def retrieve_lesson_images(discipline, topic, lesson_title, illustrations, questions, max_images=3):
    """Return safe topical photos; use local images first, then freely licensed Commons images."""
    if not _photo_topic_allowed(discipline, topic, lesson_title):
        return []
    candidates = []
    if isinstance(illustrations, list):
        for item in illustrations[:3]:
            if not isinstance(item, dict):
                continue
            query = str(item.get("search_query", "")).strip()
            caption = str(item.get("caption", "Иллюстрация к теме")).strip()
            index = int(item.get("question_index", 1) or 1)
            if query and _photo_topic_allowed(discipline, topic, lesson_title, query):
                candidates.append((index, caption, query))
    if not candidates:
        candidates = _fallback_photo_queries(discipline, topic, lesson_title, questions)
    found = []
    seen_paths = set()
    for index, caption, query in candidates:
        if len(found) >= max_images:
            break
        local = _find_local_lesson_images(discipline, topic, lesson_title, query, limit=1)
        item = local[0] if local else _fetch_commons_image(query)
        if not item:
            continue
        path = str(item["path"])
        if path in seen_paths:
            continue
        seen_paths.add(path)
        found.append({
            "path": path,
            "question_index": max(1, min(3, int(index or 1))),
            "caption": caption or item.get("caption") or "Иллюстрация к учебному вопросу",
            "credit": item.get("credit", ""),
        })
    return found


def append_lesson_images_to_cell(cell, photos, image_dir):
    """Inserts monochrome educational illustrations after the matching lesson question."""
    if not photos:
        return
    from docx.shared import Inches
    original = cell.text
    if not original.strip():
        return
    # main_leader carries canonical headings generated as "Учебный вопрос N.".
    parts = [p.strip() for p in re.split(r"(?=Учебный вопрос\s+\d+\s*\.)", original, flags=re.I) if p.strip()]
    if not parts:
        parts = [original.strip()]
    by_question = {}
    for photo in photos:
        try:
            q_index = int(photo.get("question_index", 1) or 1)
        except (TypeError, ValueError):
            q_index = 1
        by_question.setdefault(q_index, []).append(photo)
    cell.text = ""
    last_question = 0
    for part in parts:
        match = re.match(r"Учебный вопрос\s+(\d+)\s*\.", part, flags=re.I)
        if match:
            last_question = int(match.group(1))
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Inches(0.04)
        p.add_run(part)
        if match and last_question in by_question:
            for img in by_question[last_question]:
                image_path = Path(img.get("path", ""))
                if not image_path.is_file():
                    continue
                caption = cell.add_paragraph()
                caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption.paragraph_format.space_before = Inches(0.03)
                caption.paragraph_format.space_after = Inches(0.02)
                caption.add_run(f"Рисунок к учебному вопросу {last_question}: {img.get('caption', 'Учебный рисунок')}").bold = True
                pic = cell.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pic.paragraph_format.space_after = Inches(0.08)
                pic.add_run().add_picture(str(image_path), width=Inches(4.2))

def create_from_template(
    rank, full_name, discipline, topic_number, topic, lesson_number, lesson_title, lesson_type, audience, location,
    duration, goals, data, output_path, approval_position="Командир артиллерийского дивизиона"
):
    """Заполняет титульные поля и ВСЕ части конспекта в исходном Word-шаблоне."""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            "Не найден template.docx. Положи его рядом с bot.py."
        )

    doc = Document(str(TEMPLATE_PATH))

    # Блок «УТВЕРЖДАЮ» по образцу пользователя:
    # заголовок по центру, должность и звание слева, ФИО справа,
    # дата отдельной строкой ниже.
    replace_approval_fields(doc, rank, full_name)
    original_approval_position = "Командир артиллерийского дивизиона"
    if approval_position and approval_position != original_approval_position:
        set_if_found(doc, original_approval_position, approval_position)
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.shared import Inches

    approval_label = find_paragraph(doc, "УТВЕРЖДАЮ")
    approval_title = find_paragraph(doc, "Командир артиллерийского дивизиона")
    approval_line = find_paragraph(doc, rank)
    approval_date = find_paragraph(doc, "«")

    if approval_label is not None:
        approval_label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if approval_title is not None:
        approval_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if approval_date is not None:
        approval_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Звание и ФИО остаются в правом блоке утверждения на одной строке.
    if approval_line is not None and approval_line is not approval_title:
        set_paragraph_text(approval_line, f"{rank}\t{full_name}")
        approval_line.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        approval_line.paragraph_format.tab_stops.add_tab_stop(
            Inches(2.7), WD_TAB_ALIGNMENT.RIGHT
        )
    elif len(doc.paragraphs) > 2:
        approval_line = doc.paragraphs[2]
        set_paragraph_text(approval_line, f"{rank}\t{full_name}")
        approval_line.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        approval_line.paragraph_format.tab_stops.add_tab_stop(
            Inches(2.7), WD_TAB_ALIGNMENT.RIGHT
        )

    # Титульная страница и реквизиты.
    set_if_found(
        doc, "проведения занятия по",
        f"проведения занятия по {format_discipline(discipline).lower()}"
    )
    set_if_found(doc, " с  топогеодезистами", f"с личным составом {audience}")
    set_if_found(doc, "с  топогеодезистами", f"с личным составом {audience}")
    set_if_found(doc, "с топогеодезистами", f"с личным составом {audience}")
    set_if_found(doc, "Тема№1:", f'Тема № {topic_number}: «{topic}».')
    set_if_found(doc, "Занятие№1:", f"Занятие № {lesson_number}: {lesson_title}.")
    set_if_found(doc, "Время:", f"Время: {duration}." if duration.strip() else "Время:")
    set_if_found(doc, "Место:", f"Место: {location}")
    set_if_found(doc, "Руководства и пособия:", f"Руководства и пособия: {data.get('manuals') or 'Не указано'}")

    objectives = list(data.get("objectives", []))[:4]
    while len(objectives) < 4:
        objectives.append("")
    objective_prefixes = (
        "1. Ознакомить", "2. Обучить", "3. Объяснить", "4. Объяснить"
    )
    for idx, prefix in enumerate(objective_prefixes):
        set_if_found(doc, prefix, f"{idx + 1}. {objectives[idx]}" if objectives[idx] else "")

    questions = data.get("questions", [])
    questions_text = "\n".join(
        f"{i + 1}. {question}" for i, question in enumerate(questions)
    )
    set_if_found(doc, "1. Деление угломера", questions_text)

    # Материальное обеспечение — одна строка, без нумерации, через запятую.
    material_items = data.get("materials", [])
    if isinstance(material_items, str):
        material_items = re.split(r"[,;\n]+", material_items)
    cleaned_materials = []
    for item in material_items:
        item = re.sub(r"^\s*\d+[.)]\s*", "", str(item or "")).strip()
        item = item.lstrip("•-–— ").strip()
        if item and item not in cleaned_materials:
            cleaned_materials.append(item)
    if not cleaned_materials or cleaned_materials == ["Не указано"]:
        cleaned_materials = ["Учебное оборудование", "По теме занятия"]
    elif len(cleaned_materials) == 1 and normalize_text(cleaned_materials[0]) != normalize_text("По теме занятия"):
        cleaned_materials.append("По теме занятия")
    materials_line = ", ".join(cleaned_materials)
    set_if_found(doc, "1. Приборы разведки.", materials_line)
    set_if_found(doc, "2.УМБ по одиночной подготовке.", "")

    if not doc.tables:
        raise ValueError("В template.docx не найдена таблица «ХОД ЗАНЯТИЯ».")
    table = doc.tables[0]
    if len(table.rows) < 4 or len(table.columns) < 4:
        raise ValueError(
            "В таблице «ХОД ЗАНЯТИЯ» должно быть 4 строки и 4 столбца."
        )

    # Вводная часть.
    set_cell_text(table.cell(1, 0), "I. Вводная часть")
    set_cell_text(table.cell(1, 1), f"{data.get('intro_minutes', 10)} минут")
    set_cell_text(table.cell(1, 2), data.get("intro_leader", "Организую занятие, объявляю тему и цели."))
    set_cell_text(table.cell(1, 3), data.get("intro_learners", "Слушают руководителя, уясняют тему и цели занятия."))

    # Основная часть: учебные вопросы и развёрнутый материал теперь реально попадают в шаблон.
    main_questions = "\n".join(
        f"{i + 1}. {question}" for i, question in enumerate(questions)
    )
    main_leader = data.get("main_leader", "Изложить учебный материал по теме занятия.")
    main_learners = data.get("main_learners", "Слушают, конспектируют и отвечают на вопросы.")
    main_content = (
        "II. Основная часть\n\n"
        "Учебные вопросы:\n" + (main_questions or "Уточнить по программе")
    )
    set_cell_text(table.cell(2, 0), main_content)
    set_cell_text(table.cell(2, 1), f"{data.get('main_minutes', 30)} минут")
    set_cell_text(table.cell(2, 2), main_leader)
    set_cell_text(table.cell(2, 3), main_learners)
    # Рисунки сначала раскладываются после текста соответствующего учебного вопроса.
    photo_dir = Path(tempfile.mkdtemp(prefix="lesson_images_"))
    photos = retrieve_lesson_images(
        discipline=discipline,
        topic=topic,
        lesson_title=lesson_title,
        illustrations=data.get("illustrations", []),
        questions=data.get("questions", []),
        max_images=2,
    )
    append_lesson_images_to_cell(table.cell(2, 2), photos, photo_dir)

    # Схемы добавляются после основной части в ту же ячейку.
    diagrams = data.get("diagrams", [])
    if diagrams:
        diagram_dir = Path(tempfile.mkdtemp(prefix="lesson_diagrams_"))
        append_diagrams_to_cell(table.cell(2, 2), diagrams, diagram_dir)

    # Заключительная часть.
    set_cell_text(table.cell(3, 0), "III. Заключительная часть")
    set_cell_text(table.cell(3, 1), f"{data.get('conclusion_minutes', 5)} минут")
    conclusion_leader = data.get("conclusion_leader", "Подвожу итоги занятия, проверяю усвоение материала и даю задание.")
    conclusion_learners = data.get("conclusion_learners", "Отвечают на контрольные вопросы, записывают задание.")
    review_questions = data.get("review_questions", [])
    self_study = data.get("self_study", [])
    if review_questions:
        conclusion_leader += "\n\nВопросы для закрепления:\n" + "\n".join(
            f"{i + 1}. {q}" for i, q in enumerate(review_questions)
        )
    if self_study:
        conclusion_leader += "\n\nСамостоятельная подготовка:\n" + "\n".join(
            f"{i + 1}. {task}" for i, task in enumerate(self_study)
        )
    set_cell_text(table.cell(3, 2), conclusion_leader)
    set_cell_text(table.cell(3, 3), conclusion_learners)

    # Удаляем нежелательные фразы, не трогая остальное содержание.
    for table_item in doc.tables:
        for row in table_item.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    cleaned = (
                        paragraph.text
                        .replace("Дежурный докладывает о готовности группы.", "")
                        .replace("Дежурный докладывает о готовности группы", "")
                        .strip()
                    )
                    if cleaned.startswith("Метод:"):
                        cleaned = ""
                    if cleaned != paragraph.text:
                        set_paragraph_text(paragraph, cleaned)
    for paragraph in list(doc.paragraphs):
        if paragraph.text.strip().startswith("Метод:"):
            paragraph._element.getparent().remove(paragraph._element)

    # Финальная очистка: удаляем emoji и служебную пометку о виде/длительности занятия.
    for paragraph in doc.paragraphs:
        cleaned = remove_program_lesson_duration_phrase(strip_emojis(paragraph.text))
        if cleaned != paragraph.text:
            set_paragraph_text(paragraph, cleaned)
    for table_item in doc.tables:
        for row in table_item.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    cleaned = remove_program_lesson_duration_phrase(strip_emojis(paragraph.text))
                    if cleaned != paragraph.text:
                        set_paragraph_text(paragraph, cleaned)
    for section in doc.sections:
        for area in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            for paragraph in area.paragraphs:
                cleaned = remove_program_lesson_duration_phrase(strip_emojis(paragraph.text))
                if cleaned != paragraph.text:
                    set_paragraph_text(paragraph, cleaned)

    doc.save(output_path)


def validate_generated_docx(path, expected_topic="", expected_title=""):
    """Validate the Word file structurally; return warnings rather than false assurances."""
    candidate = Path(path)
    if not candidate.is_file() or candidate.stat().st_size < 1000:
        raise ValueError("Файл Word не создан или получился слишком маленьким.")
    doc = Document(str(candidate))
    if not doc.tables and not doc.paragraphs:
        raise ValueError("Документ Word не содержит текста и таблиц.")
    texts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    for section in doc.sections:
        for area in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            texts.extend(p.text for p in area.paragraphs)
    full = "\n".join(texts)
    warnings = []
    if re.search(r"\{\{[^}]+\}\}", full):
        warnings.append("В документе остались незаполненные поля шаблона.")
    if expected_topic and normalize_text(expected_topic) not in normalize_text(full):
        warnings.append("Не удалось найти название темы в тексте документа — проверьте титульную страницу.")
    if expected_title and normalize_text(expected_title) not in normalize_text(full):
        warnings.append("Не удалось найти название занятия в тексте документа — проверьте титульную страницу.")
    if not doc.tables:
        warnings.append("В документе нет таблиц; проверьте, соответствует ли он вашему шаблону.")
    if len(full.strip()) < 600:
        warnings.append("Документ выглядит коротким — проверьте содержание перед использованием.")
    return warnings


async def show_my_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает Telegram ID; команда намеренно доступна до выдачи доступа."""
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    status = "владелец" if _is_admin(user.id) else (
        "доступ выдан" if _has_access(user.id) else "доступ пока не выдан"
    )
    await message.reply_text(
        f"Ваш Telegram ID: {user.id}\nСтатус: {status}.\n\n"
        "Владелец может выдать доступ командой /grant ID или отозвать командой /revoke ID."
    )


async def grant_access(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Эта команда доступна только владельцу бота.")
        return
    if len(context.args) != 1 or not context.args[0].isdigit():
        await message.reply_text(
            "Использование: /grant TELEGRAM_ID\n"
            "Пользователь должен отправить боту команду /id и сообщить вам свой номер.\n"
            "Команда выдаёт обычный доступ, не права владельца."
        )
        return
    target_id = int(context.args[0])
    if target_id in OWNER_IDS:
        await message.reply_text(f"ID {target_id} — это владелец бота; добавлять его как пользователя не требуется.")
        return
    if target_id in AUTHORIZED_USER_IDS:
        await message.reply_text(f"У пользователя {target_id} уже есть доступ.")
        return
    AUTHORIZED_USER_IDS.add(target_id)
    try:
        _save_authorized_users()
    except OSError as exc:
        AUTHORIZED_USER_IDS.discard(target_id)
        print(f"Не удалось сохранить доступ: {exc}")
        await message.reply_text("Не удалось сохранить список доступа. Проверьте права на запись в папку бота.")
        return
    await message.reply_text(
        f"Обычный пользователь {target_id} добавлен в список доступа.\n"
        "Он может пользоваться функциями бота, но не управляет доступом."
    )


async def revoke_access(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Эта команда доступна только владельцу бота.")
        return
    if len(context.args) != 1 or not context.args[0].isdigit():
        await message.reply_text("Использование: /revoke TELEGRAM_ID")
        return
    target_id = int(context.args[0])
    if target_id in OWNER_IDS:
        await message.reply_text("Нельзя отозвать доступ владельца через эту команду.")
        return
    if target_id not in AUTHORIZED_USER_IDS:
        await message.reply_text(f"У пользователя {target_id} нет выданного доступа.")
        return
    AUTHORIZED_USER_IDS.remove(target_id)
    try:
        _save_authorized_users()
    except OSError as exc:
        AUTHORIZED_USER_IDS.add(target_id)
        print(f"Не удалось сохранить список доступа: {exc}")
        await message.reply_text("Не удалось сохранить изменения. Проверьте права на запись в папку бота.")
        return
    await message.reply_text(f"Доступ пользователя {target_id} отозван.")


async def list_authorized_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Эта команда доступна только владельцу бота.")
        return
    if not AUTHORIZED_USER_IDS:
        await message.reply_text(
            "Список пользователей с выданным доступом пуст.\n"
            "Выдача доступа: /grant TELEGRAM_ID"
        )
        return
    lines = "\n".join(f"• {item_id}" for item_id in sorted(AUTHORIZED_USER_IDS))
    await message.reply_text("Пользователи с доступом:\n" + lines)


async def access_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text(_access_denied_text(user))
        return
    await message.reply_text(
        "Управление доступом владельца:\n"
        "/id — узнать свой Telegram ID\n"
        "/grant TELEGRAM_ID — добавить обычного пользователя\n"
        "/revoke TELEGRAM_ID — отозвать доступ\n"
        "/users — список пользователей с доступом\n\n"
        "Чтобы добавить обычного пользователя: попросите его отправить боту /id, затем используйте /grant с его номером.\n"
        "Пользователь не получает команды /grant, /revoke и /users."
    )



async def show_owner_dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Кабинет владельца доступен только владельцу бота.")
        return
    users_count = len(AUTHORIZED_USER_IDS)
    created_count = int(BOT_STATS.get("created_docs", 0))
    errors_count = int(BOT_STATS.get("errors", 0))
    last_created = BOT_STATS.get("last_created_at") or "ещё не создавались"
    last_type = BOT_STATS.get("last_created_type") or "—"
    top_users = sorted(
        BOT_STATS.get("created_by_user", {}).items(),
        key=lambda pair: int(pair[1]), reverse=True
    )[:5]
    top_text = "\n".join(f"• ID {uid}: {count} документ(ов)" for uid, count in top_users) or "Пока нет данных."
    text = (
        "🛡 КАБИНЕТ ВЛАДЕЛЬЦА\n"
        "━━━━━━━━━━━━━━━━━━\n\n"
        "📊 СТАТИСТИКА\n"
        f"👥 Пользователей с доступом: {users_count}\n"
        f"📄 Создано документов: {created_count}\n"
        f"⚠️ Зафиксировано ошибок: {errors_count}\n"
        f"🕒 Последний документ: {last_created}\n"
        f"📌 Тип документа: {last_type}\n\n"
        "🏅 АКТИВНЫЕ ПОЛЬЗОВАТЕЛИ\n" + top_text + "\n\n"
        "⚙️ УПРАВЛЕНИЕ\n"
        "/grant TELEGRAM_ID — выдать обычный доступ\n"
        "/revoke TELEGRAM_ID — отозвать доступ\n"
        "/users — список пользователей\n"
        "/errors — журнал ошибок\n"
        "/reports — сообщения пользователей\n"
        "/stats — обновить статистику\n\n"
        "🔐 Кабинет доступен только владельцу. Остальные пользователи имеют обычные права."
    )
    await message.reply_text(text, reply_markup=OWNER_PANEL_KEYBOARD)


async def show_owner_errors(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Журнал ошибок доступен только владельцу бота.")
        return
    entries = _read_recent_errors(10)
    if not entries:
        await message.reply_text("✅ Журнал ошибок пуст. На данный момент ошибок не зарегистрировано.", reply_markup=OWNER_PANEL_KEYBOARD)
        return
    lines = ["ПОСЛЕДНИЕ ОШИБКИ БОТА (до 10 записей):"]
    for item in entries:
        lines.append(
            f"\n[{item.get('time', 'время неизвестно')}] {item.get('event', 'bot')}\n"
            f"Пользователь: {item.get('user_id') or 'не определён'}\n"
            f"{item.get('error', 'текст ошибки отсутствует')}"
        )
    await message.reply_text("\n".join(lines)[:3900], reply_markup=OWNER_PANEL_KEYBOARD)


async def show_user_history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _has_access(user.id):
        await message.reply_text(_access_denied_text(user))
        return
    entries = _history_entries_for(user.id)
    if not entries:
        await message.reply_text("🗂 МОИ КОНСПЕКТЫ\n━━━━━━━━━━━━━━━━━━\nИстория пока пуста. После создания документов они появятся здесь для повторного скачивания.", reply_markup=main_keyboard_for(user.id))
        return
    keyboard, mapping = history_keyboard(entries)
    context.user_data["history_pick_map"] = mapping
    context.user_data["stage"] = "history_pick"
    lines = ["Ваши последние конспекты (хранятся до 15 документов):"]
    for item in entries:
        lines.append(f"• {item.get('created_at', '')[:16]} — {item.get('discipline') or 'Конспект'}: {item.get('topic') or item.get('title')}")
    await message.reply_text("\n".join(lines)[:3800], reply_markup=keyboard)


async def show_owner_reports(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _is_admin(user.id):
        await message.reply_text("Сообщения пользователей доступны только владельцу бота.")
        return
    reports = _read_user_reports(10)
    if not reports:
        await message.reply_text("📨 Сообщений пока нет. Когда пользователь отправит отчёт об ошибке, он появится здесь.", reply_markup=OWNER_PANEL_KEYBOARD)
        return
    lines = ["ПОСЛЕДНИЕ СООБЩЕНИЯ ПОЛЬЗОВАТЕЛЕЙ:"]
    for item in reports:
        lines.append(f"\n[{item.get('time','')}] ID {item.get('user_id','?')} @{item.get('username') or 'без username'}\n{item.get('message','')}")
    await message.reply_text("\n".join(lines)[:3900], reply_markup=OWNER_PANEL_KEYBOARD)


async def owner_error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Persist uncaught errors for the owner dashboard; do not expose stack traces to users."""
    error = context.error or RuntimeError("Неизвестная ошибка обработчика Telegram")
    user_id = None
    try:
        effective_user = update.effective_user if update else None
        user_id = effective_user.id if effective_user else None
    except Exception:
        user_id = None
    _record_error(error, user_id=user_id, event="unhandled_update")
    print("Зафиксирована необработанная ошибка; подробности доступны владельцу через /errors.")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    user = update.effective_user
    if not _has_access(user.id if user else None):
        if update.effective_message:
            await update.effective_message.reply_text(_access_denied_text(user))
        return

    text = (
        "🪖 ВОЕННЫЙ КОНСПЕКТ\n"
        "━━━━━━━━━━━━━━━━━━\n"
        "Учебный помощник | 4 бригада\n\n"
        "Здравствуйте! Добро пожаловать.\n\n"
        "📘 ЧТО УМЕЕТ БОТ\n"
        "• Подготавливает учебные конспекты Word по выбранной теме.\n"
        "• Ищет материалы в подключённых учебниках.\n"
        "• Сохраняет ваши предыдущие документы.\n\n"
        "👇 Выберите действие в меню ниже."
    )

    photo_path = Path(__file__).parent / "4_brigada.jpg"
    if photo_path.exists():
        with open(photo_path, "rb") as photo:
            await update.message.reply_photo(
                photo=photo,
                caption=text,
                reply_markup=main_keyboard_for(user.id if user else None),
            )
    else:
        await update.message.reply_text(
            text + "\n\n(Чтобы отображалась эмблема, добавьте файл 4_brigada.jpg рядом с bot.py.)",
            reply_markup=main_keyboard_for(user.id if user else None),
        )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not _has_access(user.id if user else None):
        if update.effective_message:
            await update.effective_message.reply_text(_access_denied_text(user))
        return
    context.user_data.clear()
    await update.message.reply_text(
        "Действие отменено. Выберите следующий шаг.",
        reply_markup=main_keyboard_for(user.id if user else None),
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message or not message.text:
        return
    user = update.effective_user
    if not _has_access(user.id if user else None):
        await message.reply_text(_access_denied_text(user))
        return

    text = message.text.strip()
    stage = context.user_data.get("stage")
    mode = context.user_data.get("mode")

    # User-facing helpers (all remain separate from the scheduling module).
    if text == BUTTON_HISTORY:
        await show_user_history(update, context)
        return
    if text == BUTTON_SEARCH_BOOKS:
        context.user_data.clear()
        context.user_data["stage"] = "book_search_query"
        await message.reply_text(
            "🔎 ПОИСК В УЧЕБНИКАХ\n"
            "━━━━━━━━━━━━━━━━━━\n\n"
            "Введите ключевые слова или фразу, которую хотите найти.\n\n"
            "Поддерживаются PDF и DOCX. Поиск по DJVU зависит от возможностей хостинга.\n\n"
            "Напишите запрос сообщением ниже. Для выхода нажмите «Отмена».",
            reply_markup=CANCEL_KEYBOARD,
        )
        return
    if text == BUTTON_REPORT_ERROR:
        context.user_data.clear()
        context.user_data["stage"] = "report_error"
        await message.reply_text(
            "🆘 СООБЩИТЬ ОБ ОШИБКЕ\n"
            "━━━━━━━━━━━━━━━━━━\n\n"
            "Опишите, пожалуйста, проблему:\n"
            "• что вы нажали;\n"
            "• что произошло;\n"
            "• какой текст ошибки появился.\n\n"
            "🔐 Не отправляйте пароли, токены и API-ключи.\n"
            "Чтобы выйти, нажмите «Отмена».",
            reply_markup=CANCEL_KEYBOARD,
        )
        return

    if text in (BUTTON_CANCEL, "/cancel"):
        await cancel(update, context)
        return

    if text == BUTTON_BACK_TO_MENU:
        context.user_data.clear()
        await message.reply_text("🏠 ГЛАВНОЕ МЕНЮ\n━━━━━━━━━━━━━━━━━━\nВыберите нужное действие:", reply_markup=main_keyboard_for(user.id if user else None))
        return

    if stage == "history_pick":
        mapping = context.user_data.get("history_pick_map", {})
        entry_id = mapping.get(text)
        if not entry_id:
            await message.reply_text("Выберите конспект кнопкой из списка или нажмите «Главное меню».")
            return
        item, history_path = _history_item_by_id(user.id, entry_id)
        if not item or not history_path:
            await message.reply_text("Этот файл больше не найден в истории.", reply_markup=main_keyboard_for(user.id))
            context.user_data.clear()
            return
        with history_path.open("rb") as fh:
            await message.reply_document(document=fh, filename=item.get("filename") or history_path.name,
                                         caption=f"Конспект из истории от {item.get('created_at','')[:16]}.")
        return

    if stage == "book_search_query":
        query = text.strip()
        if len(query) < 3:
            await message.reply_text("Введите запрос длиннее двух символов.", reply_markup=CANCEL_KEYBOARD)
            return
        if _restricted_book_query(query):
            await message.reply_text("Этот запрос нельзя обработать. Поиск доступен по общим учебным, медицинским, правовым, строевым и другим разрешённым материалам.", reply_markup=BOOKS_KEYBOARD)
            context.user_data.clear()
            return
        status = await message.reply_text("🔎 Ищу совпадения в доступных учебниках…\nПроверяю названия и текст документов.")
        try:
            results = _search_safe_book_chunks(query, max_results=5)
            if not results:
                await status.edit_text("Совпадений не найдено. Попробуйте другие слова или откройте нужный раздел учебников.")
            else:
                blocks = [f"Результаты поиска: {query}"]
                for i, (label, snippet) in enumerate(results, start=1):
                    blocks.append(f"\n{i}. {label}\n{snippet}")
                await status.edit_text("\n".join(blocks)[:3900])
        except Exception as exc:
            _record_error(exc, user_id=user.id, event="book_search")
            await status.edit_text("Не удалось выполнить поиск. Ошибка записана для владельца.")
        context.user_data.clear()
        await message.reply_text("Можно выполнить ещё один поиск или вернуться в меню.", reply_markup=BOOKS_KEYBOARD)
        return

    if stage == "report_error":
        report_text = text.strip()
        if len(report_text) < 5:
            await message.reply_text("Напишите немного подробнее, минимум 5 символов.", reply_markup=CANCEL_KEYBOARD)
            return
        if any(k in report_text.casefold() for k in ("gemini_api_key=", "telegram_token=", "api_key=")):
            await message.reply_text("Похоже, сообщение содержит секретный ключ. Удалите его из текста и отправьте описание ошибки без ключей.", reply_markup=CANCEL_KEYBOARD)
            return
        try:
            report = _save_user_report(user.id, user.username, report_text)
            delivered = False
            for owner_id in OWNER_IDS:
                try:
                    await context.bot.send_message(
                        chat_id=owner_id,
                        text=(f"Сообщение об ошибке от пользователя ID {user.id} (@{user.username or 'без username'})\n"
                              f"Время: {report['time']}\n\n{report['message'][:2500]}")
                    )
                    delivered = True
                except Exception as exc:
                    _record_error(exc, user_id=user.id, event="notify_owner_report")
            context.user_data.clear()
            await message.reply_text(
                "Сообщение сохранено. Владелец получил уведомление." if delivered else "Сообщение сохранено; владелец сможет увидеть его в кабинете.",
                reply_markup=main_keyboard_for(user.id),
            )
        except OSError as exc:
            _record_error(exc, user_id=user.id, event="save_user_report")
            context.user_data.clear()
            await message.reply_text("⚠️ Не удалось сохранить сообщение. Попробуйте ещё раз позже.", reply_markup=main_keyboard_for(user.id))
        return

    if text == BUTTON_OWNER_DASHBOARD:
        if not user or not _is_admin(user.id):
            await message.reply_text("Кабинет владельца доступен только владельцу бота.")
            return
        context.user_data.clear()
        await show_owner_dashboard(update, context)
        return

    owner_panel_buttons = {
        BUTTON_OWNER_GRANT, BUTTON_OWNER_REVOKE, BUTTON_OWNER_USERS,
        BUTTON_OWNER_ERRORS, BUTTON_OWNER_REPORTS, BUTTON_OWNER_STATS, BUTTON_OWNER_BACK,
    }
    if text in owner_panel_buttons:
        if not user or not _is_admin(user.id):
            await message.reply_text("Эта функция доступна только владельцу бота.")
            return
        context.user_data.clear()
        if text == BUTTON_OWNER_GRANT:
            context.user_data["owner_action"] = "grant"
            await message.reply_text(
                "Введите Telegram ID пользователя, которому нужно выдать обычный доступ.\n"
                "Отмена: /cancel", reply_markup=CANCEL_KEYBOARD
            )
            return
        if text == BUTTON_OWNER_REVOKE:
            context.user_data["owner_action"] = "revoke"
            await message.reply_text(
                "Введите Telegram ID пользователя, у которого нужно отозвать доступ.\n"
                "Отмена: /cancel", reply_markup=CANCEL_KEYBOARD
            )
            return
        if text == BUTTON_OWNER_USERS:
            await list_authorized_users(update, context)
            return
        if text == BUTTON_OWNER_ERRORS:
            await show_owner_errors(update, context)
            return
        if text == BUTTON_OWNER_REPORTS:
            await show_owner_reports(update, context)
            return
        if text == BUTTON_OWNER_STATS:
            await show_owner_dashboard(update, context)
            return
        if text == BUTTON_OWNER_BACK:
            await message.reply_text("Главное меню:", reply_markup=main_keyboard_for(user.id))
            return

    if context.user_data.get("owner_action"):
        if not user or not _is_admin(user.id):
            context.user_data.clear()
            await message.reply_text("Эта функция доступна только владельцу бота.")
            return
        action = context.user_data.pop("owner_action", None)
        if not text.isdigit():
            await message.reply_text(
                "Нужен числовой Telegram ID. Попробуйте ещё раз или нажмите «Отмена».",
                reply_markup=CANCEL_KEYBOARD,
            )
            context.user_data["owner_action"] = action
            return
        target_id = int(text)
        if target_id in OWNER_IDS:
            await message.reply_text(
                "Это ID владельца. Операция не требуется и недоступна.",
                reply_markup=OWNER_PANEL_KEYBOARD,
            )
            return
        if action == "grant":
            if target_id in AUTHORIZED_USER_IDS:
                result_text = f"У пользователя {target_id} уже есть доступ."
            else:
                AUTHORIZED_USER_IDS.add(target_id)
                try:
                    _save_authorized_users()
                    result_text = f"Обычный доступ выдан пользователю {target_id}."
                except OSError as exc:
                    AUTHORIZED_USER_IDS.discard(target_id)
                    _record_error(exc, user_id=user.id, event="owner_grant_access")
                    result_text = "Не удалось сохранить доступ. Проверьте права на запись в папку проекта."
        elif action == "revoke":
            if target_id not in AUTHORIZED_USER_IDS:
                result_text = f"У пользователя {target_id} нет доступа в списке."
            else:
                AUTHORIZED_USER_IDS.remove(target_id)
                try:
                    _save_authorized_users()
                    result_text = f"Доступ пользователя {target_id} отозван."
                except OSError as exc:
                    AUTHORIZED_USER_IDS.add(target_id)
                    _record_error(exc, user_id=user.id, event="owner_revoke_access")
                    result_text = "Не удалось сохранить изменение. Проверьте права на запись в папку проекта."
        else:
            result_text = "Неизвестное действие кабинета."
        await message.reply_text(result_text, reply_markup=OWNER_PANEL_KEYBOARD)
        return

    if text == BUTTON_BOOKS:
        await message.reply_text(
            "📚 Выберите раздел учебников:",
            reply_markup=BOOKS_KEYBOARD,
        )
        return

    if text in SAFE_BOOK_BUTTON_FOLDERS:
        folder_name = SAFE_BOOK_BUTTON_FOLDERS[text]
        paths = list_safe_book_paths(folder_name)
        if not paths:
            await message.reply_text(
                "В этом разделе пока не найдено подходящих PDF/DOCX/DJVU. "
                "Можно оставить файлы в корне books/ или подпапках; имя/путь должен "
                "позволять отнести книгу к этому учебному разделу.",
                reply_markup=BOOKS_KEYBOARD,
            )
            return
        await message.reply_text(
            "Выберите документ:",
            reply_markup=book_list_keyboard([safe_book_display_name(path) for path in paths]),
        )
        return

    # Отправлять пользователю можно только документы из разрешённых общих разделов.
    safe_paths = list_safe_book_paths()
    selected_path = next((path for path in safe_paths if safe_book_display_name(path) == text), None)
    if selected_path is not None:
        try:
            with selected_path.open("rb") as document:
                await message.reply_document(document=document, filename=selected_path.name)
        except Exception:
            traceback.print_exc()
            await message.reply_text(
                "Не удалось отправить документ. Проверьте журнал ошибок сервера.",
                reply_markup=BOOKS_KEYBOARD,
            )
        return

    if text == BUTTON_NEW:
        context.user_data.clear()
        context.user_data.update({"mode": "new", "stage": "preparation_choice"})
        await message.reply_text(
            "Выберите раздел программы, по которому нужно создать конспект:",
            reply_markup=PREPARATION_KEYBOARD,
        )
        return

    if text in DISCIPLINES and text not in ("Техническая подготовка", "Специальная подготовка", "Подготовка по специальности"):
        context.user_data.clear()
        context.user_data.update({
            "mode": "new",
            "stage": "lookup_request",
            "selected_discipline": text,
        })
        await message.reply_text(
            f"Выбран раздел: {text}\n\nВведите номер темы и занятия, например: Тема 1, занятие 1.",
            reply_markup=CANCEL_KEYBOARD,
        )
        return

    if text in (BUTTON_SPECIAL, BUTTON_TECHNICAL):
        category = "technical" if text == BUTTON_TECHNICAL else "special"
        context.user_data.clear()
        context.user_data.update({
            "mode": "new",
            "stage": "special_lookup_request",
            "special_mode": True,
            "program_category": category,
            "discipline": "Техническая подготовка" if category == "technical" else "Специальная подготовка",
        })
        title = "Техническая подготовка" if category == "technical" else "Подготовка по специальности"
        await message.reply_text(
            f"{title}\n\nВведите только номер темы и номер занятия из программы.\n"
            "Например: Тема 1, занятие 1.\n\n"
            "Если есть несколько совпадений, я покажу варианты кнопками — выберите нужный.",
            reply_markup=CANCEL_KEYBOARD,
        )
        return

    if text == BUTTON_HELP:
        await message.reply_text(
            "ℹ️ ПОМОЩЬ\n"
            "━━━━━━━━━━━━━━━━━━\n\n"
            "📝 СОЗДАНИЕ КОНСПЕКТА\n"
            "1. Нажмите «Создать новый конспект».\n"
            "2. Выберите дисциплину, тему и занятие.\n"
            "3. Укажите место, продолжительность и учебные материалы.\n"
            "4. Дождитесь готового файла Word.\n\n"
            "📚 УЧЕБНИКИ\n"
            "Откройте раздел «Учебники» или воспользуйтесь поиском.\n\n"
            "🗂 МОИ КОНСПЕКТЫ\n"
            "Здесь можно повторно скачать ранее созданные документы.\n\n"
            "⚠️ Перед использованием проверьте тему, содержание и соответствие утверждённой программе.\n\n"
            "Разработчик: @sintromotmen1",
            reply_markup=main_keyboard_for(user.id if user else None),
        )
        return

    if mode == "new":
        if stage == "special_lookup_request":
            topic_match = re.search(r"тема\s*№?\s*(\d+)", text, re.I)
            lesson_match = re.search(r"занятие\s*№?\s*(\d+)", text, re.I)
            if not topic_match or not lesson_match:
                await message.reply_text(
                    "Не смог распознать номера. Напишите, например: Тема 1, занятие 1.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            category = context.user_data.get("program_category", "special")
            try:
                matches = find_special_program_lessons(topic_match.group(1), lesson_match.group(1), category)
            except Exception as error:
                print(f"Ошибка поиска темы в программе: {error}")
                await message.reply_text(
                    "Не удалось открыть программу. Проверьте, что файл программы находится рядом с bot.py.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            if not matches:
                await message.reply_text(
                    "В выбранном разделе программы не нашёл такую тему и занятие. Проверьте номера и отправьте ещё раз.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            if len(matches) > 1:
                context.user_data["special_matches"] = matches
                lines = ["Нашлось несколько вариантов. Выберите нужный вариант кнопкой:"]
                for i, item in enumerate(matches, start=1):
                    lines.append(f"{i}. {item['specialty']} — тема {item['topic_number']}: {item['topic']}; занятие {item['lesson_number']}: {item['lesson_title']}")
                context.user_data["stage"] = "special_choose_match"
                await message.reply_text("\n".join(lines), reply_markup=program_match_keyboard(len(matches)))
                return
            found = matches[0]
            context.user_data.update({
                "discipline": "Техническая подготовка" if category == "technical" else "Специальная подготовка",
                "specialty": found["specialty"],
                "topic_number": found["topic_number"],
                "topic": found["topic"],
                "lesson_number": found["lesson_number"],
                "lesson_title": found["lesson_title"],
                "lesson_type": found["lesson_type"],
                "stage": "confirm_program_match",
            })
            await message.reply_text(
                "Нашёл в программе:\n\n"
                f"Раздел: {context.user_data['discipline']}\nСпециальность/подраздел: {found['specialty']}\n"
                f"Тема № {found['topic_number']}: {found['topic']}\n"
                f"Занятие № {found['lesson_number']}: {found['lesson_title']}\n"
                f"Вид занятия: {found['lesson_type']}\n\nЭто нужное занятие?",
                reply_markup=ReplyKeyboardMarkup([[BUTTON_CONFIRM_YES], [BUTTON_CONFIRM_NO], [BUTTON_CANCEL]], resize_keyboard=True),
            )
            return
        if stage == "special_choose_match":
            matches = context.user_data.get("special_matches", [])
            m = re.fullmatch(r"Выбрать вариант\s+(\d+)", text, re.I)
            if not m or not (1 <= int(m.group(1)) <= len(matches)):
                await message.reply_text("Выберите один из вариантов кнопками ниже.", reply_markup=program_match_keyboard(len(matches)))
                return
            chosen = matches[int(m.group(1)) - 1]
            category = context.user_data.get("program_category", "special")
            context.user_data.update({
                "discipline": "Техническая подготовка" if category == "technical" else "Специальная подготовка",
                "specialty": chosen["specialty"],
                "topic_number": chosen["topic_number"],
                "topic": chosen["topic"],
                "lesson_number": chosen["lesson_number"],
                "lesson_title": chosen["lesson_title"],
                "lesson_type": chosen["lesson_type"],
                "stage": "confirm_program_match",
            })
            await message.reply_text(
                "Проверьте выбранный вариант:\n\n"
                f"Раздел: {context.user_data['discipline']}\nСпециальность/подраздел: {chosen['specialty']}\n"
                f"Тема № {chosen['topic_number']}: {chosen['topic']}\n"
                f"Занятие № {chosen['lesson_number']}: {chosen['lesson_title']}\n"
                f"Вид занятия: {chosen['lesson_type']}\n\nЭто нужное занятие?",
                reply_markup=ReplyKeyboardMarkup([[BUTTON_CONFIRM_YES], [BUTTON_CONFIRM_NO], [BUTTON_CANCEL]], resize_keyboard=True),
            )
            return
        if stage == "lookup_request":
            parsed = parse_lookup_request(text)
            if parsed:
                discipline, topic_number, lesson_number = parsed
            else:
                topic_match = re.search(r"тема\s*№?\s*(\d+)", text, re.I)
                lesson_match = re.search(r"занятие\s*№?\s*(\d+)", text, re.I)
                discipline = context.user_data.get("selected_discipline")
                if not (discipline and topic_match and lesson_match):
                    await message.reply_text(
                        "Не смог распознать номера. Напишите: Тема 1, занятие 1.",
                        reply_markup=CANCEL_KEYBOARD,
                    )
                    return
                topic_number, lesson_number = topic_match.group(1), lesson_match.group(1)
            try:
                found = find_program_lesson(discipline, topic_number, lesson_number)
            except Exception as error:
                print(f"Ошибка поиска темы в программе: {error}")
                await message.reply_text(
                    "Не удалось открыть программу. Проверьте, что файл «8. Программа БП адн рус 2026(1).docx» находится рядом с bot.py.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            if not found:
                await message.reply_text(
                    "В программе не нашёл совпадение для этих данных. Проверьте название подготовки и номера темы/занятия и отправьте ещё раз.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            if "multiple" in found:
                options = found["multiple"][:10]
                context.user_data["program_matches"] = options
                lines = ["Нашлось несколько вариантов. Выберите нужный кнопкой:"]
                for i, item in enumerate(options, start=1):
                    lines.append(f"{i}. {item['discipline']}: тема {item['topic_number']} — {item['topic']}; занятие {item['lesson_number']} — {item.get('lesson_title', 'название не найдено')}")
                context.user_data["stage"] = "choose_program_match"
                await message.reply_text("\n".join(lines), reply_markup=program_match_keyboard(len(options)))
                return
            context.user_data.update({
                "discipline": found["discipline"].title(),
                "topic_number": found["topic_number"],
                "topic": found["topic"],
                "lesson_number": found["lesson_number"],
                "lesson_title": found["lesson_title"],
                "lesson_type": found["lesson_type"],
                "stage": "confirm_program_match",
            })
            await message.reply_text(
                "Проверьте, правильно ли нашёл в программе:\n\n"
                f"Подготовка: {found['discipline']}\n"
                f"Тема № {found['topic_number']}: {found['topic']}\n"
                f"Занятие № {found['lesson_number']}\n"
                f"Название занятия: {found['lesson_title']}\n"
                f"Вид занятия: {found['lesson_type']}\n\n"
                "Это нужная тема и занятие?",
                reply_markup=ReplyKeyboardMarkup(
                    [[BUTTON_CONFIRM_YES], [BUTTON_CONFIRM_NO], [BUTTON_CANCEL]],
                    resize_keyboard=True,
                ),
            )
            return
        if stage == "choose_program_match":
            matches = context.user_data.get("program_matches", [])
            m = re.fullmatch(r"Выбрать вариант\s+(\d+)", text, re.I)
            if not m or not (1 <= int(m.group(1)) <= len(matches)):
                await message.reply_text("Выберите один из вариантов кнопками ниже.", reply_markup=program_match_keyboard(min(len(matches), 10)))
                return
            chosen = matches[int(m.group(1)) - 1]
            context.user_data.update({
                "discipline": chosen["discipline"].title(),
                "topic_number": chosen["topic_number"],
                "topic": chosen["topic"],
                "lesson_number": chosen["lesson_number"],
                "lesson_title": chosen["lesson_title"],
                "lesson_type": chosen["lesson_type"],
                "stage": "confirm_program_match",
            })
            await message.reply_text(
                "Проверьте выбранный вариант:\n\n"
                f"Подготовка: {chosen['discipline']}\n"
                f"Тема № {chosen['topic_number']}: {chosen['topic']}\n"
                f"Занятие № {chosen['lesson_number']}: {chosen['lesson_title']}\n"
                f"Вид занятия: {chosen['lesson_type']}\n\nЭто нужная тема и занятие?",
                reply_markup=ReplyKeyboardMarkup([[BUTTON_CONFIRM_YES], [BUTTON_CONFIRM_NO], [BUTTON_CANCEL]], resize_keyboard=True),
            )
            return

        if stage == "confirm_program_match":
            if text == BUTTON_CONFIRM_NO or text.casefold() in ("нет", "неверно", "изменить"):
                if context.user_data.get("special_mode"):
                    context.user_data.update({"stage": "special_lookup_request"})
                    await message.reply_text(
                        "Хорошо. Отправьте номера ещё раз в формате: Тема 1, занятие 1.",
                        reply_markup=CANCEL_KEYBOARD,
                    )
                else:
                    context.user_data.update({"stage": "lookup_request"})
                    await message.reply_text(
                        "Хорошо. Отправьте заново: название подготовки, номер темы и номер занятия.\n"
                        "Например: Тактическая подготовка, тема 2, занятие 1.",
                        reply_markup=CANCEL_KEYBOARD,
                    )
                return
            if text != BUTTON_CONFIRM_YES and text.casefold() not in ("да", "верно", "правильно"):
                await message.reply_text("Нажмите «Да, верно» или «Нет, изменить».", reply_markup=ReplyKeyboardMarkup(
                    [[BUTTON_CONFIRM_YES], [BUTTON_CONFIRM_NO], [BUTTON_CANCEL]], resize_keyboard=True
                ))
                return
            context.user_data["stage"] = "audience"
            await message.reply_text(
                "Подтверждено. выберите подразделение или введите свой вариант:",
                reply_markup=choice_keyboard(["1-й артиллерийской батареи", "2-й артиллерийской батареи", "3-й артиллерийской батареи"], add_custom=True),
            )
            return
        if stage == "lesson_type":
            context.user_data["lesson_type"] = text
            context.user_data["stage"] = "audience"
            await message.reply_text(
                "👥 выберите батарею или введите своё подразделение:",
                reply_markup=choice_keyboard(["1-й артиллерийской батареи", "2-й артиллерийской батареи", "3-й артиллерийской батареи"], add_custom=True),
            )
            return
        if stage == "audience":
            if text == BUTTON_CUSTOM:
                context.user_data["stage"] = "audience_custom"
                await message.reply_text("✍️ Введите подразделение или категорию обучаемых:", reply_markup=CANCEL_KEYBOARD)
                return
            context.user_data["audience"] = text
            context.user_data["duration"] = ""
            context.user_data["stage"] = "location"
            await message.reply_text("Укажите место проведения занятия:", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "audience_custom":
            context.user_data["audience"] = text
            context.user_data["duration"] = ""
            context.user_data["stage"] = "location"
            await message.reply_text("Укажите место проведения занятия:", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "location":
            context.user_data["location"] = text
            context.user_data["stage"] = "manuals"
            await message.reply_text(
                "📚 выберите руководство/пособие или введите свой вариант. Можно указать несколько пунктов текстом.",
                reply_markup=choice_keyboard(MANUAL_OPTIONS, add_custom=True, add_skip=True),
            )
            return
        if stage == "manuals":
            if text == BUTTON_CUSTOM:
                context.user_data["stage"] = "manuals_custom"
                await message.reply_text("✍️ Перечислите руководства и пособия. Если несколько — укажите через запятую или с новой строки:", reply_markup=CANCEL_KEYBOARD)
                return
            context.user_data["manuals"] = "" if text == BUTTON_SKIP or text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "materials"
            await message.reply_text(
                "🧰 выберите материальное обеспечение или введите свой вариант:",
                reply_markup=choice_keyboard(MATERIAL_OPTIONS, add_custom=True, add_skip=True),
            )
            return
        if stage == "manuals_custom":
            context.user_data["manuals"] = "" if text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "materials"
            await message.reply_text(
                "🧰 выберите материальное обеспечение или введите свой вариант:",
                reply_markup=choice_keyboard(MATERIAL_OPTIONS, add_custom=True, add_skip=True),
            )
            return
        if stage == "materials":
            if text == BUTTON_CUSTOM:
                context.user_data["stage"] = "materials_custom"
                await message.reply_text("✍️ Перечислите материальное обеспечение через запятую или с новой строки:", reply_markup=CANCEL_KEYBOARD)
                return
            context.user_data["materials"] = "" if text == BUTTON_SKIP or text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "goals"
            await message.reply_text("🧩 дополнительные пожелания к конспекту. Если их нет, нажмите «Не указывать» или напишите «нет»:", reply_markup=choice_keyboard([], add_skip=True))
            return
        if stage == "materials_custom":
            context.user_data["materials"] = "" if text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "goals"
            await message.reply_text("🧩 дополнительные пожелания к конспекту. Если их нет, нажмите «Не указывать» или напишите «нет»:", reply_markup=choice_keyboard([], add_skip=True))
            return
        if stage == "goals":
            context.user_data["goals"] = "" if text == BUTTON_SKIP or text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "approval_position"
            await message.reply_text(
                "Укажите должность утверждающего. Например: Командир артиллерийского дивизиона.\n"
                "Можно ввести другую должность.",
                reply_markup=CANCEL_KEYBOARD,
            )
            return
        if stage == "approval_position":
            context.user_data["approval_position"] = text.strip()
            context.user_data["stage"] = "approval_identity"
            await message.reply_text(
                "Введите воинское звание и ФИО утверждающего одной строкой.\n"
                "Например: Капитан М. Еспенов",
                reply_markup=CANCEL_KEYBOARD,
            )
            return
        if stage == "approval_identity" or (stage == "retry_generation" and text == BUTTON_RETRY):
            if stage == "approval_identity":
                identity = text.strip()
                # Recognize common multi-word military ranks; the remaining text is the name.
                rank_patterns = [
                    r"старший\s+лейтенант", r"младший\s+лейтенант", r"старший\s+сержант",
                    r"младший\s+сержант", r"сержант\s+3\s+класса", r"сержант\s+2\s+класса",
                    r"сержант\s+1\s+класса", r"главный\s+сержант", r"штаб-сержант",
                    r"старшина", r"подполковник", r"капитан", r"лейтенант", r"сержант",
                    r"майор", r"полковник", r"генерал", r"ефрейтор", r"рядовой",
                ]
                matched = None
                for pattern in rank_patterns:
                    m = re.match(r"^\s*(" + pattern + r")\s+(.+?)\s*$", identity, flags=re.I)
                    if m:
                        matched = (m.group(1).strip(), m.group(2).strip())
                        break
                if not matched:
                    parts = identity.split(maxsplit=1)
                    if len(parts) < 2:
                        await message.reply_text(
                            "Укажите и звание, и ФИО в одной строке. Например: Майор К. С. Серикпаев",
                            reply_markup=CANCEL_KEYBOARD,
                        )
                        return
                    matched = (parts[0], parts[1])
                context.user_data["rank"], context.user_data["full_name"] = matched
            status = await message.reply_text("⚙️ ГОТОВЛЮ КОНСПЕКТ\n━━━━━━━━━━━━━━━━━━\nПодбираю материал и заполняю Word-шаблон. Это может занять некоторое время…")
            output_path = None
            generation_succeeded = False
            try:
                data = generate_lesson_data(
                    context.user_data["discipline"],
                    context.user_data["topic"],
                    context.user_data["lesson_title"],
                    context.user_data["lesson_type"],
                    context.user_data["audience"],
                    context.user_data["duration"],
                    (f"Специальность по программе: {context.user_data.get('specialty')}. " if context.user_data.get("specialty") else "") + context.user_data["goals"],
                    context.user_data["manuals"],
                    context.user_data["materials"],
                )
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp:
                    output_path = temp.name

                create_from_template(
                    context.user_data["rank"],
                    context.user_data["full_name"],
                    context.user_data["discipline"],
                    context.user_data["topic_number"],
                    context.user_data["topic"],
                    context.user_data["lesson_number"],
                    context.user_data["lesson_title"],
                    context.user_data["lesson_type"],
                    context.user_data["audience"],
                    context.user_data["location"],
                    context.user_data["duration"],
                    context.user_data["goals"],
                    data,
                    output_path,
                    context.user_data.get("approval_position", "Командир артиллерийского дивизиона"),
                )

                quality_warnings = validate_generated_docx(
                    output_path,
                    expected_topic=context.user_data.get("topic", ""),
                    expected_title=context.user_data.get("lesson_title", ""),
                )
                with open(output_path, "rb") as document_file:
                    await message.reply_document(
                        document=document_file,
                        filename="Plan-konspekt.docx",
                        caption=(
                            "✅ КОНСПЕКТ ГОТОВ\n"
                            "━━━━━━━━━━━━━━━━━━\n"
                            f"📘 Дисциплина: {context.user_data['discipline']}\n"
                            f"📑 Тема № {context.user_data['topic_number']}: {context.user_data['topic']}\n"
                            f"📝 Занятие № {context.user_data['lesson_number']}: {context.user_data['lesson_title']} ({context.user_data['lesson_type']})\n\n"
                            "Файл Word приложен к сообщению. Перед использованием проверьте материал и распределение времени."
                            + ("\n\nПроверка документа:\n• " + "\n• ".join(quality_warnings) if quality_warnings else "\n\nАвтоматическая проверка структуры пройдена.")
                        ),
                    )
                try:
                    _archive_document(
                        user.id,
                        output_path,
                        context.user_data.get("lesson_title", "Конспект"),
                        context.user_data.get("discipline", ""),
                        context.user_data.get("topic", ""),
                    )
                except Exception as archive_error:
                    _record_error(archive_error, user_id=user.id, event="archive_document")
                await status.edit_text("✅ Готово!\nДокумент создан на основе вашего Word-шаблона.")
                generation_succeeded = True
                _record_generated_doc(user.id, "конспект")
            except Exception as error:
                _record_error(error, user_id=user.id if user else None, event="lesson_generation")
                print("\n=== ОШИБКА СОЗДАНИЯ КОНСПЕКТА ===")
                traceback.print_exc()
                print("=== КОНЕЦ ОШИБКИ ===\n")
                error_text = str(error).upper()
                is_quota_error = "429" in error_text or "RESOURCE_EXHAUSTED" in error_text
                is_temporary_error = any(marker in error_text for marker in ("503", "UNAVAILABLE", "500", "502", "504"))
                if is_quota_error:
                    user_message = (
                        "Достигнут лимит Gemini API (429). Подожди сброса лимита "
                        "или проверь квоты и доступность модели в Google AI Studio."
                    )
                elif is_temporary_error:
                    user_message = (
                        "ИИ-сервис временно недоступен. Ваши ответы сохранены — попробуйте позже."
                    )
                else:
                    user_message = (
                        "Не удалось создать конспект. Проверьте подключение к интернету, API-ключи "
                        "и наличие template.docx рядом с bot.py. Подробности ошибки есть в командной строке."
                    )
                # edit_text принимает только inline-клавиатуру, а ReplyKeyboardMarkup
                # отправляется отдельным сообщением через reply_text.
                context.user_data["stage"] = "retry_generation"
                if is_quota_error or is_temporary_error:
                    status_text = (
                        "Это очень плохая ошибка напиши в Поддержку."
                    )
                else:
                    status_text = (
                        user_message
                        + "\\n\\nТвои ответы сохранены. После исправления проблемы нажми «Повторить создание»."
                    )
                await status.edit_text(status_text)
                await message.reply_text(
                    "Выбери действие:",
                    reply_markup=ReplyKeyboardMarkup(
                        [[BUTTON_RETRY], [BUTTON_CANCEL]], resize_keyboard=True
                    ),
                )
            finally:
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
                if generation_succeeded:
                    context.user_data.clear()
            return

    await message.reply_text("🏠 ГЛАВНОЕ МЕНЮ\n━━━━━━━━━━━━━━━━━━\nВыберите нужное действие:", reply_markup=main_keyboard_for(user.id if user else None))



async def issue_site_code_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Issues a short-lived one-time code for the website to already-authorized users."""
    user = update.effective_user
    message = update.effective_message
    if not user or not message:
        return
    if not _has_access(user.id):
        await message.reply_text(_access_denied_text(user))
        return
    try:
        code, ttl_seconds = issue_site_login_code(user.id)
    except Exception as exc:
        _record_error(exc, user_id=user.id, event="site_login_code")
        await message.reply_text("Не удалось создать код входа. Попробуйте позже.")
        return
    minutes = max(1, ttl_seconds // 60)
    await message.reply_text(
        "🌐 ВХОД НА САЙТ «ВОЕННЫЙ КОНСПЕКТ»\n"
        "━━━━━━━━━━━━━━━━━━\n\n"
        f"Ваш Telegram ID: {user.id}\n"
        f"Одноразовый код: {code}\n\n"
        f"Введите ID и код на сайте. Код действует {minutes} минут и используется один раз.\n"
        "Никому не пересылайте код. Если вы не запрашивали вход, просто проигнорируйте сообщение."
    )

def main(run_in_thread=False):
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    # /id доступна всем, чтобы человек без доступа мог сообщить свой ID владельцу.
    app.add_handler(CommandHandler("id", show_my_id))
    app.add_handler(CommandHandler("sitecode", issue_site_code_command))
    app.add_handler(CommandHandler("myid", show_my_id))
    app.add_handler(CommandHandler("grant", grant_access))
    app.add_handler(CommandHandler("revoke", revoke_access))
    app.add_handler(CommandHandler("users", list_authorized_users))
    app.add_handler(CommandHandler("access_help", access_help))
    app.add_handler(CommandHandler("owner", show_owner_dashboard))
    app.add_handler(CommandHandler("stats", show_owner_dashboard))
    app.add_handler(CommandHandler("errors", show_owner_errors))
    app.add_handler(CommandHandler("reports", show_owner_reports))
    app.add_handler(CommandHandler("history", show_user_history))
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(owner_error_handler)
    print("Бот «Военный Конспект» запущен.")
    if not OWNER_IDS:
        print("ВНИМАНИЕ: не настроен OWNER_ID в .env. Напишите боту /id, затем укажите свой ID в OWNER_ID и перезапустите бота.")
    else:
        print(f"Владелец доступа настроен; обычных пользователей с доступом: {len(AUTHORIZED_USER_IDS)}")
    detected_books = list_safe_book_paths()
    print(f"Найдено безопасных документов в books/: {len(detected_books)}")
    for detected_path in detected_books:
        print(f"  - {safe_book_display_name(detected_path)}")
    if not BOOKS_ROOT.exists():
        print(f"Папка библиотеки не найдена: {BOOKS_ROOT}")
    # When hosted together with the web API, this runs in a worker thread and must not register OS signal handlers.
    if run_in_thread:
        app.run_polling(stop_signals=None)
    else:
        app.run_polling()


if __name__ == "__main__":
    main()
