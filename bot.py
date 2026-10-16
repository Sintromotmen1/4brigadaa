import os
import json
import tempfile
import time
import random
import traceback
import re
import urllib.request
import urllib.error
from pathlib import Path

from dotenv import load_dotenv
from docx import Document
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
    raise ValueError("Не найден TELEGRAM_TOKEN в файле .env")
if not GEMINI_API_KEY:
    raise ValueError("Не найден GEMINI_API_KEY в файле .env")
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
BUTTON_BOOKS = "📚 Учебники"
BUTTON_ARTILLERY_BOOKS = "📚 Артиллерийские учебники"
BUTTON_MINBAT_BOOKS = "📘 МинБат учебники"
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
    lesson_pattern = re.compile(r"^занятие\s*№?\s*(\d+)\b\s*(.*)$", re.I | re.S)
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
            current_topic = {"number": match.group(1), "title": match.group(2).strip(" .:—-"), "section": current_section}
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

BOOKS_KEYBOARD = ReplyKeyboardMarkup(
    [[BUTTON_ARTILLERY_BOOKS], [BUTTON_MINBAT_BOOKS], [BUTTON_BACK_TO_MENU]],
    resize_keyboard=True,
)

# Названия книг выводятся отдельными кнопками в соответствующем разделе.
ARTILLERY_BOOKS = [
    "курс подготовки артиллерии.docx",
    "Курс стрельб ВС РК.docx",
    "Правила боевого применения Сухопутных войск Вооруженных Сил .docx",
    "Руководство по боевой работе огневых подразделений артиллерии.docx",
    "сборник нормативов.docx",
]
MINBAT_BOOKS = []

# Библиотека для обычных учебных и справочных материалов.
# Помещайте документы в папку books/general рядом с bot.py.
GENERAL_BOOKS_DIR = Path(__file__).parent / "books" / "general"
GENERAL_BOOKS = []  # Пример: ["история.docx", "общая_подготовка.pdf"]


def book_list_keyboard(book_names):
    rows = [[name] for name in book_names]
    rows.append([BUTTON_BACK_TO_MENU])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)
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


def generate_lesson_data(discipline, topic, lesson_title, lesson_type, audience, duration, goals, manuals, materials):
    """Просит Gemini подготовить содержание для разделов исходного Word-шаблона."""
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
Вид занятия: {lesson_type}
Категория обучаемых: {audience}
Продолжительность: {duration}
Руководства и пособия (предоставлены пользователем): {manuals}
Материальное обеспечение (предоставлено пользователем): {materials}
Дополнительные пожелания: {goals or "Не указаны"}

Верни только корректный JSON со следующими ключами:
{{
  "objectives": ["цель 1", "цель 2", "цель 3", "цель 4"],
  "questions": ["учебный вопрос 1", "учебный вопрос 2", "учебный вопрос 3"],
  "materials": ["материальное обеспечение 1", "материальное обеспечение 2"],
  "manuals": "руководства и учебные пособия; не выдумывай конкретные названия",
  "intro_minutes": 10,
  "main_minutes": 30,
  "conclusion_minutes": 5,
  "intro_leader": "действия руководителя во вводной части",
  "intro_learners": "действия обучаемых во вводной части",
  "main_leader": "содержательное объяснение учебных вопросов с понятными подзаголовками",
  "main_learners": "действия обучаемых в основной части",
  "conclusion_leader": "подведение итогов и задание на самостоятельную подготовку",
  "conclusion_learners": "действия обучаемых в заключительной части",
  "review_questions": ["вопрос для закрепления 1", "вопрос 2", "вопрос 3"],
  "self_study": ["задание 1", "задание 2"],
  "diagrams": [
    {{"question_index": 1, "title": "Краткое название схемы", "nodes": ["Понятие", "Основные элементы", "Проверка понимания"]}}
  ]
}}

Требования:
- objectives: ровно 3 короткие учебные цели.
- questions: от 2 до 3 учебных вопросов.
- Время вводной, основной и заключительной частей в сумме должно примерно соответствовать
  указанной продолжительности. Если длительность указана в часах, переведи её в минуты.
- main_leader — это основная часть графы «Краткое содержание (действия руководителя)».
  Не ограничивайся кратким пересказом: подробно раскрой КАЖДЫЙ учебный вопрос по порядку.
  Для каждого вопроса дай отдельный заголовок и несколько содержательных абзацев:
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
- diagrams: создай 1–3 простые учебные схемы, соответствующие теме и учебным вопросам. Для каждой укажи question_index (номер вопроса начиная с 1), короткий title и nodes (от 3 до 5 коротких подписей, не длиннее 7 слов каждая). Схемы должны объяснять общие понятия, классификацию, этапы обучения или меры безопасности. Не включай координаты, расчёты наведения, процедуры применения оружия или иные оперативные инструкции. Если схема неуместна, верни пустой список.
- Не используй смайлики, эмодзи или декоративные символы в любом тексте конспекта.
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
        "objectives", "questions",
        "intro_minutes", "main_minutes", "conclusion_minutes",
        "intro_leader", "intro_learners", "main_leader", "main_learners",
        "conclusion_leader", "conclusion_learners", "review_questions", "self_study",
    ]
    missing = [key for key in required if key not in data]
    if missing:
        raise ValueError(f"Модель {provider_used} не вернула обязательные поля: " + ", ".join(missing))
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
    """Рисует простую последовательную учебную схему, не меняя Word-шаблон."""
    nodes = [str(x).strip()[:90] for x in diagram.get("nodes", []) if str(x).strip()][:5]
    if len(nodes) < 3:
        return False
    width = 1200
    box_w, box_h = 940, 112
    gap = 48
    margin_top, margin_bottom = 95, 60
    height = margin_top + len(nodes) * box_h + (len(nodes) - 1) * gap + margin_bottom
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title = str(diagram.get("title", "Учебная схема"))[:100]
    draw.text((width // 2, 25), title, fill="#1f2937", font=_diagram_font(34, True), anchor="ma")
    x0 = (width - box_w) // 2
    for i, node in enumerate(nodes):
        y0 = margin_top + i * (box_h + gap)
        y1 = y0 + box_h
        draw.rounded_rectangle((x0, y0, x0 + box_w, y1), radius=18, outline="#334155", width=4, fill="#eef4f8")
        # Перенос текста на две строки, чтобы подписи не выходили за границы блока.
        font = _diagram_font(28)
        words = node.split()
        lines, line = [], ""
        for word in words:
            candidate = (line + " " + word).strip()
            if draw.textbbox((0, 0), candidate, font=font)[2] > box_w - 70 and line:
                lines.append(line)
                line = word
            else:
                line = candidate
        if line:
            lines.append(line)
        lines = lines[:2]
        line_heights = [draw.textbbox((0, 0), part, font=font)[3] for part in lines]
        total_h = sum(line_heights) + max(0, len(lines) - 1) * 8
        yy = y0 + (box_h - total_h) // 2
        for part, line_h in zip(lines, line_heights):
            draw.text((width // 2, yy), part, fill="#111827", font=font, anchor="ma")
            yy += line_h + 8
        if i < len(nodes) - 1:
            cx = width // 2
            top = y1 + 6
            bottom = y1 + gap - 7
            draw.line((cx, top, cx, bottom), fill="#475569", width=5)
            draw.polygon([(cx - 12, bottom - 14), (cx + 12, bottom - 14), (cx, bottom + 2)], fill="#475569")
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
        picture_paragraph.add_run().add_picture(str(image_path), width=Inches(3.0))


def create_from_template(
    rank, full_name, discipline, topic_number, topic, lesson_number, lesson_title, lesson_type, audience, location,
    duration, goals, data, output_path
):
    """Заполняет титульные поля и ВСЕ части конспекта в исходном Word-шаблоне."""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            "Не найден template.docx. Положи его рядом с bot.py."
        )

    doc = Document(str(TEMPLATE_PATH))

    # Блок «УТВЕРЖДАЮ»: сохраняем расположение в правой верхней части,
    # звание и ФИО размещаем НА ОДНОЙ строке, как в образце пользователя.
    replace_approval_fields(doc, rank, full_name)
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.shared import Inches

    approval_label = find_paragraph(doc, "УТВЕРЖДАЮ")
    approval_title = find_paragraph(doc, "Командир артиллерийского дивизиона")
    approval_line = find_paragraph(doc, rank)
    approval_date = find_paragraph(doc, "«")

    # Заголовок, должность и дата остаются в правой части страницы.
    for paragraph in (approval_label, approval_title, approval_date):
        if paragraph is not None:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Звание слева, ФИО справа на одной строке; не переносим их друг под друга.
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
    set_if_found(doc, "Занятие№1:", f"Занятие № {lesson_number}: {lesson_title} ({lesson_type}).")
    set_if_found(doc, "Время:", f"Время: {duration}.")
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
    await update.message.reply_text(
        "Здравствуйте! Это бот «Военный Конспект».\n\n"
        "Выберите действие:\n"
        "• создать новый конспект Word с помощью Google Gemini;\n"
        "• открыть помощь по работе с ботом.",
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

    if text == BUTTON_ARTILLERY_BOOKS:
        await message.reply_text(
            "📚 Артиллерийские учебники. Выберите название:",
            reply_markup=book_list_keyboard(ARTILLERY_BOOKS),
        )
        return

    if text == BUTTON_MINBAT_BOOKS:
        if not MINBAT_BOOKS:
            await message.reply_text(
                "📘 В разделе «МинБат учебники» пока нет названий. "
                "Добавьте их в список MINBAT_BOOKS в bot.py.",
                reply_markup=BOOKS_KEYBOARD,
            )
            return
        await message.reply_text(
            "📘 МинБат учебники. Выберите название:",
            reply_markup=book_list_keyboard(MINBAT_BOOKS),
        )
        return

    if text in ARTILLERY_BOOKS or text in MINBAT_BOOKS:
        # В этих разделах пока показываем выбор названия. Отправку документов
        # включайте только для обычных учебных/справочных материалов,
        # размещённых в books/general.
        await message.reply_text(
            f"Вы выбрали: {text}\n\n"
            "Для отправки обычного учебного или справочного документа добавьте "
            "его в папку books/general и укажите имя файла в списке GENERAL_BOOKS.",
            reply_markup=book_list_keyboard(ARTILLERY_BOOKS if text in ARTILLERY_BOOKS else MINBAT_BOOKS),
        )
        return

    if text in GENERAL_BOOKS:
        file_path = GENERAL_BOOKS_DIR / text
        if not file_path.is_file():
            await message.reply_text(
                f"Файл не найден: {text}\nПроверьте, что он находится в папке books/general.",
                reply_markup=KEYBOARD,
            )
            return
        try:
            with file_path.open("rb") as document:
                await message.reply_document(document=document, filename=file_path.name)
        except Exception:
            await message.reply_text("Не удалось отправить файл. Проверьте его формат и права доступа.", reply_markup=KEYBOARD)
        return

    if text == BUTTON_NEW:
        context.user_data.clear()
        context.user_data.update({"mode": "new", "stage": "lookup_request"})
        await message.reply_text(
            "📚 Напишите в одном сообщении направление подготовки, номер темы и номер занятия.\n\n"
            "Например: Тактическая подготовка, тема 2, занятие 1.\n"
            "Я найду запись в программе и спрошу, правильно ли выбраны тема и занятие.",
            reply_markup=CANCEL_KEYBOARD,
        )
        return

    if text == BUTTON_HELP:
        await message.reply_text(
            "📘 ПОМОЩЬ — РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ\n"
            "Бот «Военный Конспект»\n\n"
            "Бот помогает подготовить учебный конспект в формате Microsoft Word "
            "на основе установленного шаблона с использованием Google Gemini.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "📝 1. СОЗДАНИЕ КОНСПЕКТА\n\n"
            "Нажмите «Создать новый конспект». Бот последовательно задаст вопросы.\n\n"
            "① Отправьте в одном сообщении направление подготовки, номер темы и номер занятия.\n"
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
            "После ввода данных Gemini подготовит содержание, а бот отправит готовый файл Word.\n"
            "Перед использованием проверьте тему, подразделение, время, место проведения, "
            "руководства, материальное обеспечение и соответствие конспекта утверждённой программе.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "⚙️ 3. ЕСЛИ ВОЗНИКЛА ОШИБКА\n\n"
            "• Проверьте подключение к интернету и повторите попытку.\n"
            "• Ошибка Gemini 503 UNAVAILABLE может означать временную перегрузку сервиса. Попробуйте позже.\n"
            "• Если ошибка повторяется, отправьте разработчику её текст.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "💡 4. РЕКОМЕНДАЦИИ\n\n"
            "• Формулируйте тему точно и по возможности без сокращений.\n"
            "• Указывайте полное наименование подразделения.\n"
            "• Используйте актуальные руководства и пособия.\n"
            "• Автоматически созданный текст может содержать неточности — обязательно проверьте документ.\n"
            "• Не отправляйте в бот пароли, API-ключи и другие конфиденциальные данные.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "👨‍💻 РАЗРАБОТЧИК\n"
            "Telegram: @sintromotmen1\n"
            "По вопросам работы бота, предложениям и сообщениям об ошибках обращайтесь к разработчику.",
            reply_markup=KEYBOARD,
        )
        return

    if mode == "new":
        if stage == "lookup_request":
            parsed = parse_lookup_request(text)
            if not parsed:
                await message.reply_text(
                    "Не смог распознать запрос. Напишите, например: Тактическая подготовка, тема 2, занятие 1.",
                    reply_markup=CANCEL_KEYBOARD,
                )
                return
            discipline, topic_number, lesson_number = parsed
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
                options = found["multiple"][:5]
                lines = ["Нашлось несколько совпадений. Уточните название подготовки и отправьте запрос ещё раз:"]
                for item in options:
                    lines.append(f"• {item['discipline']}: тема {item['topic_number']} — {item['topic']}; занятие {item['lesson_number']} — {item.get('lesson_title', 'название не найдено')}")
                await message.reply_text("\n".join(lines), reply_markup=CANCEL_KEYBOARD)
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
        if stage == "confirm_program_match":
            if text == BUTTON_CONFIRM_NO or text.casefold() in ("нет", "неверно", "изменить"):
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
                "Подтверждено. Шаг 2 — выберите подразделение или введите свой вариант:",
                reply_markup=choice_keyboard(["1-й артиллерийской батареи", "2-й артиллерийской батареи", "3-й артиллерийской батареи"], add_custom=True),
            )
            return
        if stage == "lesson_type":
            context.user_data["lesson_type"] = text
            context.user_data["stage"] = "audience"
            await message.reply_text(
                "👥 Шаг 3 — выберите батарею или введите своё подразделение:",
                reply_markup=choice_keyboard(["1-й артиллерийской батареи", "2-й артиллерийской батареи", "3-й артиллерийской батареи"], add_custom=True),
            )
            return
        if stage == "audience":
            if text == BUTTON_CUSTOM:
                context.user_data["stage"] = "audience_custom"
                await message.reply_text("✍️ Введите подразделение или категорию обучаемых:", reply_markup=CANCEL_KEYBOARD)
                return
            context.user_data["audience"] = text
            context.user_data["stage"] = "duration"
            await message.reply_text("⏱️ Шаг 4 — введите продолжительность свободным текстом (например, 45 минут или 2 часа):", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "audience_custom":
            context.user_data["audience"] = text
            context.user_data["stage"] = "duration"
            await message.reply_text("⏱️ Шаг 4 — введите продолжительность свободным текстом (например, 45 минут или 2 часа):", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "duration":
            context.user_data["duration"] = text
            context.user_data["stage"] = "location"
            await message.reply_text("📍 Шаг 5 — укажите место проведения занятия:", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "location":
            context.user_data["location"] = text
            context.user_data["stage"] = "manuals"
            await message.reply_text(
                "📚 Шаг 6 — выберите руководство/пособие или введите свой вариант. Можно указать несколько пунктов текстом.",
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
                "🧰 Шаг 7 — выберите материальное обеспечение или введите свой вариант:",
                reply_markup=choice_keyboard(MATERIAL_OPTIONS, add_custom=True, add_skip=True),
            )
            return
        if stage == "manuals_custom":
            context.user_data["manuals"] = "" if text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "materials"
            await message.reply_text(
                "🧰 Шаг 7 — выберите материальное обеспечение или введите свой вариант:",
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
            await message.reply_text("🧩 Шаг 8 — дополнительные пожелания к конспекту. Если их нет, нажмите «Не указывать» или напишите «нет»:", reply_markup=choice_keyboard([], add_skip=True))
            return
        if stage == "materials_custom":
            context.user_data["materials"] = "" if text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "goals"
            await message.reply_text("🧩 Шаг 8 — дополнительные пожелания к конспекту. Если их нет, нажмите «Не указывать» или напишите «нет»:", reply_markup=choice_keyboard([], add_skip=True))
            return
        if stage == "goals":
            context.user_data["goals"] = "" if text == BUTTON_SKIP or text.lower() in ("нет", "-", "не указано") else text
            context.user_data["stage"] = "approval_rank"
            await message.reply_text("🖋️ Шаг 9 — введите воинское звание утверждающего:", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "approval_rank":
            context.user_data["rank"] = text
            context.user_data["stage"] = "approval_name"
            await message.reply_text("🖋️ Шаг 10 — введите ФИО утверждающего:", reply_markup=CANCEL_KEYBOARD)
            return
        if stage == "approval_name" or (stage == "retry_generation" and text == BUTTON_RETRY):
            if stage == "approval_name":
                context.user_data["full_name"] = text
            status = await message.reply_text("⏳ ИИ готовит содержание и заполняет ваш Word-шаблон. Подождите немного…")
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
                    context.user_data["goals"],
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
    app.run_polling()


if __name__ == "__main__":
    main()
