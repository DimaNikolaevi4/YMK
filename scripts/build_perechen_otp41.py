# -*- coding: utf-8 -*-
"""Собрать данные перечней ОТП-41 (МДК 04.02, 04.03) из РП ПМ.04.

Источник: RP/01_Актуальные/15.02.14_РП_ПМ.04_2023.docx (сконвертирован из .doc).
Титул печатается по учебному плану ОТП-41: 15.02.07, 8-й семестр.
Темы, работы и требования берём через parse_rp() — тот же разбор, что и в валидаторе.

Запуск из корня репозитория:
    py scripts/build_perechen_otp41.py
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from validate_perechen import parse_rp

RP = ROOT / "RP" / "01_Актуальные" / "15.02.14_РП_ПМ.04_2023.docx"

PM_TITLE = "Осуществление текущего мониторинга состояния систем автоматизации"
EDU_TITLE = "Автоматизация технологических процессов и производств (по отраслям)"

# Типичные склейки из РП — правим только в формулировках вопросов.
GLUE_FIXES = {
    "Принципысоздания": "Принципы создания",
    "диагностикив": "диагностики в",
    "надежностии": "надежности и",
    "эффективностиавтоматизированного": "эффективности автоматизированного",
    "АСУПовышение": "АСУ. Повышение",
    "Диагностика.Диагностика": "Диагностика. Диагностика",
    "надежностиНадежность": "надежности. Надежность",
}


def clean_req(items: list[str]) -> list[str]:
    out = []
    for item in items:
        item = item.strip()
        if item == "Вариативная часть":
            out.append("Вариативная часть:")
            continue
        out.append(re.sub(r"^-\s*", "", item))
    return out


def theory_heading(text: str) -> str:
    parts = re.split(r"(?<=[а-я])(?=[А-ЯЁ])", text.strip())
    head = parts[0].strip() if parts else text.strip()
    if len(head) < 12:
        head = " ".join(parts[:2]).strip()
    return head[:140]


def readable(text: str) -> str:
    text = re.sub(r"(?<=[а-я])(?=[А-ЯЁ])", " ", text)
    for bad, good in GLUE_FIXES.items():
        text = text.replace(bad, good)
    return text.rstrip(".")


def questions(topic: dict) -> list[str]:
    """До 5 вопросов на тему: по заголовкам теории и практическим работам."""
    name = readable(topic["name"])
    out = [f"Каково содержание темы: {name}?"]
    for theory in topic["theory"]:
        if len(out) >= 5:
            break
        out.append(f"{name}: {readable(theory_heading(theory))}?")
    for work in topic["practical"]:
        if len(out) >= 5:
            break
        out.append(f"{name}: {readable(work['name'])}?")
    return out


def build(code: str) -> dict:
    rp = parse_rp(RP, code, "04")
    mdk_name = rp["mdk_name"]
    topics, exam = [], []
    work_number = 0
    for topic in rp["topics"]:
        practicals = []
        for work in topic["practical"]:
            work_number += 1
            practicals.append(f"Практическое занятие №{work_number} {work['name']}")
        if practicals:
            topics.append({
                "topic": topic["name"],
                "practical": practicals,
                "lab": [],
                "independent": [],
            })
        exam.extend(questions(topic))

    return {
        "meta": {
            "study_year": "2026-2027",
            "doc_kind": "mdk",
            "pm_code": "ПМ 04.",
            "pm_title": PM_TITLE,
            "mdk_code": f"МДК {code}",
            "mdk_title": mdk_name,
            "education_kind": "специальность",
            "education_code": "15.02.07",
            "education_title": EDU_TITLE,
            "groups": ["ОТП-41"],
            "course": 4,
            "semester": 8,
            "teachers": ["Бардаков Д.Н."],
            "year_for_print": "2026",
        },
        "approval": {
            "deputy_director": "Т.В. Якимова",
            "commission_chair": "А.Н. Ткаченко",
            "commission_name": "цикловой комиссии технических дисциплин",
            "protocol_number": "",
            "protocol_date": "",
        },
        "control": [
            {
                "form": "Экзамен по модулю ПМ 04 (ПА.04)",
                "section": f"МДК {code}, темы "
                           + ", ".join(t["name"].split()[1] for t in rp["topics"]),
                "lesson": "ПА.04",
            }
        ],
        "requirements": {
            "heading": None,
            "intro": "С целью овладения указанным видом профессиональной деятельности "
                     "и соответствующими профессиональными компетенциями обучающийся "
                     "в ходе освоения профессионального модуля должен",
            "practice_obligatory": clean_req(rp["requirements"]["practice_obligatory"]),
            "skills_obligatory": clean_req(rp["requirements"]["skills_obligatory"]),
            "knowledge_obligatory": clean_req(rp["requirements"]["knowledge_obligatory"]),
        },
        "works": {
            "mdk_header": f"МДК {code} {mdk_name}",
            "topics": topics,
            "footer": None,
        },
        "course_project": None,
        "exam_questions": {
            "title": "Вопросы к экзамену по модулю ПМ 04 (ПА.04)",
            "items": exam[:32],
        },
        "criteria": "extended",
        "student_lists": [{"group": "ОТП-41", "teacher_sign": "Бардаков Д.Н."}],
        "student_list_empty_rows": 25,
    }


def main() -> None:
    for code in ("04.02", "04.03"):
        data = build(code)
        out = ROOT / "scripts" / f"data_perechen_mdk{code.replace('.', '')}_otp41.json"
        out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            out.name,
            "тем", len(data["works"]["topics"]),
            "работ", sum(len(t["practical"]) for t in data["works"]["topics"]),
            "вопросов", len(data["exam_questions"]["items"]),
        )


if __name__ == "__main__":
    main()
