import os
import re
import json
import time
import unicodedata
from pathlib import Path

from dotenv import load_dotenv
from docx import Document

load_dotenv()

# Files are expected in the same folder as this script.
SOURCE_FILE = Path("8. Программа БП адн рус 2026(1).docx")
TEMPLATE_FILE = Path("template(1).docx")
OUTPUT_DIR = Path("Готовые_конспекты")
INDEX_FILE = OUTPUT_DIR / "index.json"

# Select provider via .env: AI_PROVIDER=mistral or AI_PROVIDER=gemini
AI_PROVIDER = os.getenv("AI_PROVIDER", "auto").strip().lower()
MISTRAL_MODEL = os.getenv("MISTRAL_MODEL", "mistral-small-latest")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

def clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip()

def safe_filename(s, limit=95):
    s = unicodedata.normalize("NFKD", s)
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" ._")
    return (s[:limit] or "Конспект")

def extract_lessons(path):
    doc = Document(path)
    lessons = []
    chapter = "Раздел программы"
    discipline = "Общая подготовка"
    current_topic = "Тема не указана"
    topic_number = ""
    topic_count = 0

    for p in doc.paragraphs:
        text = clean(p.text)
        if not text:
            continue

        if re.match(r"^ГЛАВА\b", text, re.I):
            chapter = text

        # Uppercase headings often identify a discipline/section.
        if (len(text) >= 4 and text == text.upper()
                and not text.startswith(("ТЕМА", "ЗАНЯТИЕ", "ГЛАВА"))
                and len(text) < 150
                and re.search(r"[А-ЯЁ]", text)):
            discipline = text.title()

        mt = re.match(r"^Тема\s*№?\s*([\d.]+)?\s*[:.]?\s*(.*)$", text, re.I)
        if mt:
            topic_count += 1
            topic_number = (mt.group(1) or str(topic_count)).strip(".")
            current_topic = clean(mt.group(2)) or f"Тема {topic_number}"
            continue

        ml = re.match(r"^Занятие\s*№?\s*([\d.]+)?\s*[:.]?\s*(.*)$", text, re.I)
        if ml:
            lesson_no = (ml.group(1) or str(sum(1 for x in lessons if x["topic"] == current_topic) + 1)).strip(".")
            lesson_title = clean(ml.group(2)) or current_topic
            # Avoid adding lines that are clearly only a repeated table label.
            if len(lesson_title) < 3:
                lesson_title = current_topic
            lessons.append({
                "chapter": chapter,
                "discipline": discipline,
                "topic_no": topic_number,
                "topic": current_topic,
                "lesson_no": lesson_no,
                "lesson": lesson_title,
                "source_text": text
            })

    # Stable IDs and duplicate handling
    for i, item in enumerate(lessons, 1):
        item["id"] = i
    return lessons

def call_ai(lesson):
    system_prompt = """
Ты помогаешь оформить учебный план-конспект по предоставленной программе. Пиши на русском языке,
официально, грамотно, без выдуманных ссылок на нормативные документы. Верни только корректный JSON.
Сделай учебный материал общеобразовательным, теоретическим и безопасным.
Не создавай пошаговые инструкции по ведению боя, применению или наведению оружия, выбору целей,
стрельбе, поражению людей/объектов, установке/обезвреживанию мин и взрывчатых устройств,
тактическому проникновению, обходу охраны или другим действиям, способным причинить вред.
Если тема относится к таким областям, изложи только общие понятия, правовые рамки, безопасность,
профилактику несчастных случаев, обязанности по соблюдению команд и обращению за инструктажем
к уполномоченному инструктору. Не добавляй опасных технических деталей.
Не приписывай источнику сведения, которых в названии занятия нет; если информации мало,
подготовь базовый план занятия с пометкой, что конкретные детали уточняет руководитель.
"""
    user_prompt = f"""
Подготовь один план-конспект для учебного занятия по шаблону.
Раздел программы: {lesson['chapter']}
Дисциплина/раздел: {lesson['discipline']}
Тема № {lesson['topic_no']}: {lesson['topic']}
Занятие № {lesson['lesson_no']}: {lesson['lesson']}
Исходная строка программы: {lesson['source_text']}

Верни JSON со следующими ключами:
"objectives": массив из 3-4 учебных целей;
"study_questions": массив из 2-4 учебных вопросов;
"duration_minutes": число минут (если длительность указана в названии занятия, переведи часы в минуты; иначе 45);
"place": место проведения;
"method": метод проведения;
"manuals": массив руководств/пособий. Не выдумывай конкретные номера приказов, страниц и названия книг;
"materials": массив учебных материалов;
"intro_leader": текст действий руководителя во вводной части;
"intro_learners": текст действий обучаемых во вводной части;
"main_leader": подробный, но безопасный учебный материал с подзаголовками и пояснениями, примерно 500-900 слов; для опасных тем только общая теория и безопасность;
"main_learners": текст действий обучаемых в основной части;
"conclusion_leader": текст действий руководителя в заключительной части;
"conclusion_learners": текст действий обучаемых в заключительной части;
"homework": задание для самостоятельной подготовки.
"""
    if AI_PROVIDER == "gemini":
        from google import genai
        client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=system_prompt + "\n\n" + user_prompt,
            config={"response_mime_type": "application/json"}
        )
        raw = response.text
    else:
        from mistralai import Mistral
        client = Mistral(api_key=os.getenv("MISTRAL_API_KEY"))
        response = client.chat.complete(
            model=MISTRAL_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.2,
        )
        raw = response.choices[0].message.content

    if isinstance(raw, list):
        raw = "".join(str(part) for part in raw)
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I)
    return json.loads(raw)

def set_paragraph_text(paragraph, text):
    # Keep the paragraph's style while replacing its text.
    paragraph.text = str(text or "")

def replace_paragraph_prefix(doc, prefix, replacement):
    for p in doc.paragraphs:
        if clean(p.text).lower().startswith(prefix.lower()):
            set_paragraph_text(p, replacement)
            return True
    return False

def fill_template(template_path, output_path, lesson, data):
    doc = Document(template_path)

    replace_paragraph_prefix(
        doc, "проведения занятия по",
        f"проведения занятия по {lesson['discipline'].lower()}"
    )
    replace_paragraph_prefix(
        doc, "с топогеодезистами",
        "с личным составом подразделения"
    )
    replace_paragraph_prefix(
        doc, "на «___»",
        "на «___» __________ 2026 года"
    )

    # Replace topic and lesson lines.
    replace_paragraph_prefix(
        doc, "тема",
        f"Тема № {lesson['topic_no']}: «{lesson['topic']}»"
    )
    replace_paragraph_prefix(
        doc, "занятие",
        f"Занятие № {lesson['lesson_no']}: {lesson['lesson']}"
    )

    # Replace only numbered lines between "Учебные цели" and "Учебные вопросы".
    paragraphs = list(doc.paragraphs)
    goals = data.get("objectives", [])
    goal_heading_idx = next((i for i, p in enumerate(paragraphs) if clean(p.text).lower().startswith("учебные цели")), None)
    q_heading_idx = next((i for i, p in enumerate(paragraphs) if clean(p.text).lower().startswith("учебные вопросы")), None)
    if goal_heading_idx is not None and q_heading_idx is not None:
        goal_lines = [p for p in paragraphs[goal_heading_idx + 1:q_heading_idx]
                      if re.match(r"^\d+\.", clean(p.text))]
        for i, p in enumerate(goal_lines):
            set_paragraph_text(p, f"{i+1}. {goals[i]}" if i < len(goals) else "")
        anchor = paragraphs[q_heading_idx]
        for i in range(len(goal_lines), len(goals)):
            newp = anchor.insert_paragraph_before(f"{i+1}. {goals[i]}")
            newp.style = anchor.style

    # Replace the existing study-question paragraph(s).
    questions = data.get("study_questions", [])
    paragraphs = list(doc.paragraphs)
    q_heading_idx = next((i for i, p in enumerate(paragraphs) if clean(p.text).lower().startswith("учебные вопросы")), None)
    time_idx = next((i for i, p in enumerate(paragraphs) if clean(p.text).lower().startswith("время:")), None)
    if q_heading_idx is not None:
        end_idx = time_idx if time_idx is not None else min(q_heading_idx + 3, len(paragraphs))
        old_qs = [p for p in paragraphs[q_heading_idx + 1:end_idx] if clean(p.text)]
        if old_qs:
            set_paragraph_text(old_qs[0], "\n".join(f"{i+1}. {q}" for i, q in enumerate(questions)))
            for p in old_qs[1:]:
                set_paragraph_text(p, "")
        elif questions:
            newp = paragraphs[q_heading_idx].insert_paragraph_before("\n".join(f"{i+1}. {q}" for i, q in enumerate(questions)))
            newp.style = paragraphs[q_heading_idx].style
    replace_paragraph_prefix(doc, "время:", f"Время: {int(data.get('duration_minutes',45))} минут.")
    replace_paragraph_prefix(doc, "место:", f"Место: {data.get('place','учебный класс')}")
    replace_paragraph_prefix(doc, "метод:", f"Метод: {data.get('method','лекция с обсуждением')}")
    manuals = data.get("manuals", ["Материалы учебной программы"])
    replace_paragraph_prefix(doc, "руководства и пособия:", "Руководства и пособия: " + "; ".join(map(str, manuals)))
    materials = data.get("materials", ["Учебная программа, доска/экран, раздаточные материалы"])
    replace_paragraph_prefix(doc, "материальное обеспечение:", "Материальное обеспечение: " + "; ".join(map(str, materials)))

    duration = int(data.get('duration_minutes', 45))
    intro_minutes = max(5, round(duration * 0.1))
    main_minutes = max(25, round(duration * 0.8))
    conclusion_minutes = max(5, duration - intro_minutes - main_minutes)

    if doc.tables:
        table = doc.tables[0]
        # Expected template layout: intro, main, conclusion rows.
        if len(table.rows) >= 4:
            intro, main, conclusion = table.rows[1], table.rows[2], table.rows[3]
            intro.cells[0].text = "I. Вводная часть"
            intro.cells[1].text = f"{intro_minutes} минут"
            intro.cells[2].text = str(data.get("intro_leader", "Проверяю готовность, объявляю тему, цели и порядок занятия."))
            intro.cells[3].text = str(data.get("intro_learners", "Слушают руководителя, готовятся к занятию, уточняют организационные вопросы."))

            questions_text = "\n".join(f"{i+1}. {q}" for i, q in enumerate(questions))
            main.cells[0].text = "II. Основная часть\n" + questions_text
            main.cells[1].text = f"{main_minutes} минут"
            main.cells[2].text = str(data.get("main_leader", "Излагаю основные понятия темы, поясняю определения и отвечаю на вопросы."))
            main.cells[3].text = str(data.get("main_learners", "Слушают, ведут записи, отвечают на вопросы, выполняют учебные задания под руководством преподавателя."))

            conclusion.cells[0].text = "III. Заключительная часть"
            conclusion.cells[1].text = f"{conclusion_minutes} минут"
            conclusion.cells[2].text = str(data.get("conclusion_leader", "Подвожу итоги, проверяю усвоение материала и довожу задание на самостоятельную подготовку: " + str(data.get("homework","повторить основные понятия темы."))))
            conclusion.cells[3].text = str(data.get("conclusion_learners", "Отвечают на контрольные вопросы, записывают задание и уточняют непонятные положения."))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)

def main():
    global AI_PROVIDER
    if AI_PROVIDER == "auto":
        if os.getenv("MISTRAL_API_KEY"):
            AI_PROVIDER = "mistral"
        elif os.getenv("GEMINI_API_KEY"):
            AI_PROVIDER = "gemini"
        else:
            raise RuntimeError("В .env нет MISTRAL_API_KEY или GEMINI_API_KEY")
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Не найден файл программы: {SOURCE_FILE.resolve()}")
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"Не найден шаблон: {TEMPLATE_FILE.resolve()}")

    if AI_PROVIDER == "gemini" and not os.getenv("GEMINI_API_KEY"):
        raise RuntimeError("В .env не найден GEMINI_API_KEY")
    if AI_PROVIDER != "gemini" and not os.getenv("MISTRAL_API_KEY"):
        raise RuntimeError("В .env не найден MISTRAL_API_KEY")

    OUTPUT_DIR.mkdir(exist_ok=True)
    lessons = extract_lessons(SOURCE_FILE)
    if not lessons:
        raise RuntimeError("Не удалось найти строки «Тема» и «Занятие» в программе.")

    index = []
    for lesson in lessons:
        folder = OUTPUT_DIR / safe_filename(lesson["chapter"], 65) / safe_filename(lesson["discipline"], 65)
        filename = f"{lesson['id']:03d}_Тема_{safe_filename(lesson['topic'], 45)}_Занятие_{safe_filename(lesson['lesson'], 45)}.docx"
        target = folder / filename

        if target.exists():
            print(f"[УЖЕ ЕСТЬ] {target}")
            index.append({"id": lesson["id"], "topic": lesson["topic"], "lesson": lesson["lesson"], "file": str(target)})
            continue

        print(f"[{lesson['id']}/{len(lessons)}] Генерирую: {lesson['topic']} — {lesson['lesson']}")
        try:
            data = call_ai(lesson)
            fill_template(TEMPLATE_FILE, target, lesson, data)
            index.append({"id": lesson["id"], "chapter": lesson["chapter"], "discipline": lesson["discipline"], "topic": lesson["topic"], "lesson": lesson["lesson"], "file": str(target)})
            print(f"[ГОТОВО] {target}")
            time.sleep(1.5)
        except Exception as exc:
            print(f"[ОШИБКА] {lesson['id']}: {exc}")
            # Keep going; the next run can resume by skipping completed files.
            continue

        INDEX_FILE.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")

    INDEX_FILE.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nГотово. Найдено занятий: {len(lessons)}. Индекс: {INDEX_FILE.resolve()}")
    print("Проверьте несколько документов вручную: автоматическая генерация требует редакторской проверки.")

if __name__ == "__main__":
    main()
