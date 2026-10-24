import os
import json
import tempfile
import time
import random
import traceback
import re
import shutil
import subprocess
from functools import lru_cache
import urllib.request
import urllib.error
from pathlib import Path

from dotenv import load_dotenv
from docx import Document
from pypdf import PdfReader
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw, ImageFont
from google import genai

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
BUTTON_CONFIRM_YES = "✅ Да, верно"
BUTTON_CONFIRM_NO = "❌ Нет, изменить"

BUTTON_NEW = "📝 Создать новый конспект"
BUTTON_SPECIAL = "🎯 По специальности"
BUTTON_TECHNICAL = "🛠 Техническая подготовка"
BUTTON_PICK_PREFIX = "Выбрать вариант "
BUTTON_BOOKS = "📚 Учебники"
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

    required = [
        "objectives", "questions", "question_sections",
        "intro_minutes", "main_minutes", "conclusion_minutes",
        "intro_leader", "intro_learners", "main_learners",
        "conclusion_leader", "conclusion_learners", "review_questions", "self_study",
    ]
    missing = [key for key in required if key not in data]
    if missing:
        raise ValueError(f"Модель {provider_used} не вернула обязательные поля: " + ", ".join(missing))

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
    # Схемы добавляются внутрь ячейки «Краткое содержание (действия руководителя)».
    # Каждая подпись связывает схему с соответствующим учебным вопросом.
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


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()

    text = (
        "🪖 ВОЕННЫЙ КОНСПЕКТ | 4 БРИГАДА\n\n"
        "Здравствуйте, товарищ военнослужащий!\n\n"
        "Добро пожаловать в помощник по подготовке учебных конспектов.\n\n"
        "📚 Создавайте новые конспекты в формате Word по темам и занятиям "
        "из программы боевой подготовки.\n\n"
        "Выберите необходимое действие в меню ниже.\n\n"
        "С уважением,\nВоенный Конспект | 4 бригада"
    )

    photo_path = Path(__file__).parent / "4_brigada.jpg"
    if photo_path.exists():
        with open(photo_path, "rb") as photo:
            await update.message.reply_photo(
                photo=photo,
                caption=text,
                reply_markup=KEYBOARD,
            )
    else:
        await update.message.reply_text(
            text + "\n\n(Чтобы отображалась эмблема, добавьте файл 4_brigada.jpg рядом с bot.py.)",
            reply_markup=KEYBOARD,
        )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Действие отменено. Выберите следующий шаг.",
        reply_markup=KEYBOARD,
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message or not message.text:
        return

    text = message.text.strip()
    stage = context.user_data.get("stage")
    mode = context.user_data.get("mode")

    if text in (BUTTON_CANCEL, "/cancel"):
        await cancel(update, context)
        return

    if text == BUTTON_BACK_TO_MENU:
        context.user_data.clear()
        await message.reply_text("Главное меню:", reply_markup=KEYBOARD)
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
            "📘 ПОМОЩЬ — РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ\n"
            "Бот «4 бригада»\n\n"
            "Бот помогает подготовить учебный конспект в формате Microsoft Word "
            "━━━━━━━━━━━━━━━━━━\n"
            "📝 1. СОЗДАНИЕ КОНСПЕКТА\n\n"
            "Нажмите «Создать новый конспект», выберите «По специальности» или «Техническая подготовка», затем введите номер темы и занятия. При совпадениях выберите вариант кнопкой.\n\n"
            "① Бот найдёт запись в программе и попросит подтвердить её.\n"
            "Пример: Тактическая подготовка, тема 2, занятие 1. Бот найдёт запись в программе и попросит подтвердить её.\n\n"
            "② Личный состав — укажите подразделение или категорию обучаемых.\n"
            "Пример: 3 артиллерийская батарея артиллерийского дивизиона.\n\n"
            "③ Продолжительность — укажите время занятия. Например, 45 минут или 2 часа.\n\n"
            "④ Место проведения — например, учебный класс или учебная площадка.\n\n"
            "⑤ Руководство и пособия — перечислите используемые руководства, учебники и пособия.\n\n"
            "⑥ Материальное обеспечение — укажите необходимые учебные материалы, оборудование и наглядные пособия.\n\n"
            "⑦ Дополнительные пожелания — при необходимости укажите требования к содержанию. Если их нет, напишите «Нет».\n\n"
            "⑧–⑨ Данные утверждающего — введите воинское звание и ФИО.\n\n"
            "━━━━━━━━━━━━━━━━━━\n\n"
            "📄 2. ПОЛУЧЕНИЕ ДОКУМЕНТА\n\n"
            "После ввода данных подготовит содержание, а бот отправит готовый файл Word.\n"
            "Перед использованием проверьте тему, подразделение, время, место проведения, "
            "руководства, материальное обеспечение и соответствие конспекта утверждённой программе.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "💡 3. РЕКОМЕНДАЦИИ\n\n"
            "• Формулируйте тему точно и по возможности без сокращений.\n"
            "• Указывайте полное наименование подразделения.\n"
            "• Используйте актуальные руководства и пособия.\n"
            "• Автоматически созданный текст может содержать неточности — обязательно проверьте документ.\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "👨‍💻 РАЗРАБОТЧИК\n"
            "Telegram: @sintromotmen1\n"
            "По вопросам работы бота, предложениям и сообщениям об ошибках обращайтесь к разработчику.",
            reply_markup=KEYBOARD,
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
            status = await message.reply_text("⏳ ИИ готовит содержание. Подождите немного…")
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

                with open(output_path, "rb") as document_file:
                    await message.reply_document(
                        document=document_file,
                        filename="Plan-konspekt.docx",
                        caption=(
                            "✅ Конспект подготовлен в исходном Word-шаблоне.\n"
                            f"Дисциплина: {context.user_data['discipline']}\n"
                            f"Тема № {context.user_data['topic_number']}: {context.user_data['topic']}\n"
                            f"Занятие № {context.user_data['lesson_number']}: {context.user_data['lesson_title']} ({context.user_data['lesson_type']})\n"
                            "Проверьте содержание и распределение времени перед использованием."
                        ),
                    )
                await status.edit_text("Готово! Документ оформлен на основе вашего шаблона.")
                generation_succeeded = True
            except Exception as error:
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

    await message.reply_text("Чтобы начать, выберите действие ниже.", reply_markup=KEYBOARD)


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Бот «Военный Конспект» запущен.")
    detected_books = list_safe_book_paths()
    print(f"Найдено безопасных документов в books/: {len(detected_books)}")
    for detected_path in detected_books:
        print(f"  - {safe_book_display_name(detected_path)}")
    if not BOOKS_ROOT.exists():
        print(f"Папка библиотеки не найдена: {BOOKS_ROOT}")
    app.run_polling()


if __name__ == "__main__":
    main()
