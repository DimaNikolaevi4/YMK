# -*- coding: utf-8 -*-
"""Импорт образца «Перечня материалов» в JSON для generate_perechen_docx.py.

Одноразовый инструмент: разбирает конкретный образец (по умолчанию —
«Перечни/перечень ПМ 02 МДК 02.01.docx») и раскладывает его по блокам схемы:
meta, control, requirements, works, course_project, criteria, student_lists.

Тексты переносятся дословно; нормализуются только пробелы и разбивка
склеенных работ («Практическая работа №1 … Практическая работа №2 …»)
на отдельные пункты.

Запуск из корня репозитория:
    python scripts/import_perechen_etalon.py \
        --src "Перечни/перечень ПМ 02 МДК 02.01.docx" \
        --out scripts/data_perechen_etalon.json
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document

# --- метаданные образца (в документе размазаны по абзацам, проще зафиксировать) ---
META = {
    "study_year": "2022-2023",
    "doc_kind": "mdk",
    "pm_code": "ПМ 02.",
    "pm_title": "Теоретические основы организации монтажа, ремонта, наладки систем "
    "автоматического управления, средств измерений и мехатронных систем",
    "mdk_code": "МДК 02.01",
    "mdk_title": "Теоретические основы организации монтажа, ремонта, наладки систем "
    "автоматического управления, средств измерений и мехатронных систем",
    "education_kind": "специальность",
    "education_code": "15.02.07",
    "education_title": "Автоматизация технологических процессов и производств",
    "groups": ["АТП-41", "АТП-42"],
    "teachers": ["Бардаков Д.Н.", "Озеров И.Н."],
    "year_for_print": "2022",
}

APPROVAL = {
    "deputy_director": "Т.В. Якимова",
    "commission_chair": "А.Н. Ткаченко",
    "commission_name": "цикловой комиссии технических дисциплин",
    "protocol_number": "",
    "protocol_date": "",
}

# Заголовки подсекций требований → ключи схемы
REQ_HEADINGS = [
    ("практический опыт", "обязательная", "practice_obligatory"),
    ("должен уметь", None, "skills_obligatory"),
    ("должен знать", None, "knowledge_obligatory"),
    ("практический опыт", "вариативная", "practice_variative"),
    ("уметь", "вариативная", "skills_variative"),
    ("знать", "вариативная", "knowledge_variative"),
]

WORK_MARK = re.compile(
    r"(?=(?:Практическая|Самостоятельная)\s+работ[аы]?\s*№\s*\d+)", re.IGNORECASE
)


def norm(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text.replace(" »", "»")


def merge_continuations(items: list[str]) -> list[str]:
    """Склеить пункты, разорванные переносом строки в исходнике."""
    out: list[str] = []
    for item in items:
        if out and not out[-1].rstrip().endswith((";", ".", ":", "?", "!")):
            out[-1] = f"{out[-1]} {item}"
        else:
            out.append(item)
    return out


def split_works(text: str) -> list[str]:
    """«Практическая работа №1 А … Практическая работа №2 Б» → два пункта."""
    flat = norm(text)
    if not flat:
        return []
    items = [norm(p) for p in WORK_MARK.split(flat) if norm(p)]
    return [i[:-1] if i.endswith("»»") else i for i in items]


def is_req_heading(text: str) -> str | None:
    low = norm(text).lower()
    if not low.endswith(":"):
        return None
    for needle, qualifier, key in REQ_HEADINGS:
        if needle not in low:
            continue
        if qualifier and qualifier not in low:
            continue
        if not qualifier and "вариативная" in low:
            continue
        return key
    return None


def parse_requirements(doc: Document) -> dict:
    out: dict = {"heading": None, "intro": "", "practice_obligatory": []}
    current: str | None = None
    started = False
    for par in doc.paragraphs:
        text = norm(par.text)
        if not started:
            if text.lower().startswith("в результате изучения"):
                started = True
                out["intro"] = text
            continue
        if text.lower().startswith("перечень самостоятельных"):
            break
        key = is_req_heading(text)
        if key:
            current = key
            out.setdefault(current, [])
            continue
        if not text or current is None:
            continue
        item = text.lstrip("-–—").strip()
        if item:
            out[current].append(item)
    for _, _, key in REQ_HEADINGS:
        if isinstance(out.get(key), list):
            out[key] = merge_continuations(out[key])
    return out


def parse_works(doc: Document) -> dict:
    table = doc.tables[1]
    rows = table.rows
    header = [norm(c.text) for c in rows[0].cells]
    mdk_header = norm(rows[1].cells[0].text)
    topics = []
    footer = dict(topic="", text="")
    for row in rows[2:]:
        cells = row.cells
        topic = norm(cells[0].text)
        practical = split_works(cells[1].text)
        lab = split_works(cells[2].text)
        independent = split_works(cells[3].text)
        if topic.lower().startswith("вид контроля"):
            footer = {"topic": topic, "text": ""}
            continue
        if not topic:
            continue
        topics.append(
            {
                "topic": topic,
                "practical": practical,
                "lab": lab,
                "independent": independent,
            }
        )
        _ = (practical, lab, independent)
    return {
        "mdk_header": mdk_header,
        "topics": topics,
        "footer": footer,
    }


COURSE_START = "курсовое проектирование по"
COURSE_END = "критерии оценки знаний"
PART_MARK = re.compile(
    r"^[123]\s+(Пояснительная записка|Графическая часть|Информационная часть)", re.IGNORECASE
)


def parse_course_project(doc: Document) -> dict | None:
    lines: list[str] = []
    inside = False
    for par in doc.paragraphs:
        text = norm(par.text)
        low = text.lower()
        if not inside:
            if low.startswith(COURSE_START):
                inside = True
                lines.append(text)
            continue
        if low.startswith(COURSE_END):
            break
        lines.append(text)

    if not lines:
        return None

    title = lines[0]
    rest = lines[1:]
    body = [ln for ln in rest if ln]

    # «Примерное задание» / «Задание» идут до вводных строк задания
    lead: list[str] = []
    while body and body[0] in {"Примерное задание", "Задание"}:
        lead.append(body.pop(0))

    intro_center: list[str] = []
    while body and body[0].lower().startswith(("на курсовой проект", "по профессиональному модулю")):
        intro_center.append(body.pop(0))

    intro: list[str] = []
    while body and not body[0].lower().startswith("курсовой проект выполняется"):
        intro.append(body.pop(0))

    scope_note = body.pop(0) if body and body[0].lower().startswith("курсовой проект выполняется") else ""

    while body and not body[0].lower().startswith("исходные данные"):
        if body[0].lower().startswith(("2 графическая", "3 информационная", "дата выдачи")):
            break
        body.pop(0)

    input_heading = ""
    input_data: list[str] = []
    if body and body[0].lower().startswith("исходные данные"):
        input_heading = body.pop(0)
        while body and not re.match(r"^[123]\s", body[0]) and not body[0].lower().startswith("дата выдачи"):
            input_data.append(body.pop(0).lstrip("-–—").strip())
        input_data = merge_continuations([item for item in input_data if item])

    parts: list[dict] = []
    footer: list[str] = []
    while body:
        line = body.pop(0)
        if PART_MARK.match(line):
            parts.append({"title": line, "items": []})
        elif line.lower().startswith(("дата выдачи", "дата окончания", "преподаватель")):
            footer.append(line)
        elif parts:
            parts[-1]["items"].append(line)

    return {
        "title": title,
        "lead": lead,
        "intro_center": intro_center,
        "intro": intro,
        "scope_note": scope_note,
        "input_heading": input_heading,
        "input_data": input_data,
        "parts": parts,
        "footer": footer,
    }


def parse_criteria(doc: Document) -> dict | None:
    paras = [norm(p.text) for p in doc.paragraphs]
    start = next((i for i, t in enumerate(paras) if t.lower().startswith(COURSE_END)), None)
    if start is None:
        return None
    title = paras[start]
    items: list[str] = []
    for text in paras[start + 1 :]:
        if text.lower().startswith("студенты группы"):
            break
        if text:
            items.append(text.lstrip("\t"))
    return {"title": title, "items": items}


def parse_control(doc: Document) -> list[dict]:
    table = doc.tables[0]
    rows = []
    for row in table.rows[1:]:
        cells = [c.text.strip() for c in row.cells]
        rows.append(
            {
                "form": cells[1],
                "section": cells[2],
                "lesson": cells[3],
            }
        )
    return rows


def parse_student_lists(doc: Document) -> list[dict]:
    """Группы и подписи преподавателей из блоков «Студенты группы …»."""
    groups: list[dict] = []
    for par in doc.paragraphs:
        text = norm(par.text)
        if text.lower().startswith("студенты группы"):
            match = re.search(r"группы\s+_*([А-ЯЁA-Zа-яёa-z0-9\-]+)_*", text)
            groups.append({"group": match.group(1) if match else "", "teacher_sign": ""})
            continue
        if not groups or not text.lower().startswith("преподаватель"):
            continue
        # подпись — только в строке с подчёркиванием: «Преподаватель ____ Озеров И.Н.»
        if "_" not in text:
            continue
        sign = text.split("_")[-1].strip()
        if sign:
            groups[-1]["teacher_sign"] = sign
    return groups


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--src", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    doc = Document(args.src)
    data = {
        "meta": META,
        "approval": APPROVAL,
        "control": parse_control(doc),
        "requirements": parse_requirements(doc),
        "works": parse_works(doc),
        "course_project": parse_course_project(doc),
        "exam_questions": None,
        "criteria": parse_criteria(doc),
        "student_lists": parse_student_lists(doc),
        "student_list_empty_rows": 25,
    }

    out = Path(args.out)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Сохранено: {out}")
    print(
        "Требования: "
        + ", ".join(f"{k}={len(v) if isinstance(v, list) else '-'}" for k, v in data["requirements"].items() if k != "intro")
    )
    print(f"Тем в таблице работ: {len(data['works']['topics'])}")
    print(f"Секций курсового: {len(data['course_project']['parts']) if data['course_project'] else 0}")
    print(f"Пунктов критериев: {len(data['criteria']['items']) if data['criteria'] else 0}")
    print(f"Групп в листах: {len(data['student_lists'])}")


if __name__ == "__main__":
    main()
