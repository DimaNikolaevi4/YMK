# -*- coding: utf-8 -*-
"""Проверка «Перечня материалов» на соответствие рабочей программе (РП).

Сверяет собранный .docx с РП и КТП:
  * наименование МДК и профессии на титуле;
  * форму промежуточной аттестации и номер занятия;
  * темы и практические занятия — с РП;
  * требования к результатам (практический опыт / уметь / знать);
  * вопросы к зачёту: количество и привязка к содержанию РП;
  * лист ознакомления: группа и число пустых строк.

Запуск из корня репозитория:
    python scripts/validate_perechen.py \
        --docx "Перечни/перечень ЭТАЛОН.docx" \
        --rp "RP/02_Оригиналы/15.01.31_РП_ПМ.02_2024.docx" \
        --ktp "KTP/КТП МДК 02.01 КИП-31.docx"

Код возврата: 0 — все проверки пройдены, 1 — есть провалы.
"""

from __future__ import annotations

import argparse
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

PASS, FAIL, WARN = "✅", "❌", "⚠️"

STOPWORDS = {
    "который", "которая", "которые", "выполняемых", "выполнения", "выполнение",
    "порядок", "средства", "системы", "систем", "устройств", "устройства",
    "работы", "работ", "работа", "проверки", "проверка", "испытания", "испытаний",
    "наладка", "наладки", "состав", "общие", "сведения", "основные", "основных",
    "перечень", "видов", "виды", "целях", "случае", "место", "период",
}


def norm(text: str) -> str:
    """Привести текст к сравнимому виду: регистр, ё, пунктуация, пробелы."""
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def significant(text: str) -> set[str]:
    return {w for w in norm(text).split() if len(w) >= 5 and w not in STOPWORDS}


def similar(a: str, b: str) -> float:
    return SequenceMatcher(None, norm(a), norm(b)).ratio()


def split_sentences(text: str) -> list[str]:
    """Разбить ячейку требований на пункты и нормализовать концы."""
    parts = [p.strip() for p in re.split(r"(?<=[.;])\s+(?=[а-яёa-z])", text) if p.strip()]
    return [p.strip(" ;.") for p in parts if p.strip(" ;.")]


# ------------------------------------------------------------------ разбор документов


def row_texts(row) -> list[str]:
    """Тексты ячеек строки как есть (по одному tc), без учёта объединений."""
    out = []
    for tc in row._tr.findall(qn("w:tc")):
        text = "".join(t.text or "" for t in tc.iter(qn("w:t")))
        out.append(re.sub(r"\s+", " ", text).strip())
    return out


def cell_lines(cell) -> list[str]:
    lines = [re.sub(r"\s+", " ", p.text).strip() for p in cell.paragraphs]
    return [line for line in lines if line]


def table_rows(table) -> list[list[list[str]]]:
    """Строки таблицы: для каждой ячейки — список строк внутри неё."""
    result = []
    for row in table.rows:
        seen, cells = set(), []
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            cells.append(cell_lines(cell))
        result.append(cells)
    return result


def parse_rp(path: Path) -> dict:
    doc = Document(path)
    paragraphs = [re.sub(r"\s+", " ", p.text).strip() for p in doc.paragraphs]

    pm_name = ""
    for text in paragraphs:
        if re.match(r"^ПМ[.\s]?02\b", text) and len(text) > len(pm_name):
            pm_name = text

    profession = ""
    for text in paragraphs:
        if "професси" in text.lower() and re.search(r"\b\d{2}\.\d{2}\.\d{2}\b", text):
            profession = text
            break

    # «Элементы модуля»: № | МДК.02.01. Наименование | Форма аттестации
    mdk_code, mdk_name, attestation = "", "", ""
    has_module_exam = False
    for table in doc.tables:
        for row in table.rows:
            cells = [re.sub(r"\s+", " ", c.text).strip() for c in row.cells]
            joined = " ".join(cells).lower()
            if "экзамен" in joined and "модул" in joined:
                has_module_exam = True
            if not mdk_code and len(cells) >= 3 and re.match(r"МДК[.\s]*02\.01", cells[1]):
                mdk_code = "МДК 02.01"
                mdk_name = re.sub(r"^МДК[.\s]*02\.01[.\s]*", "", cells[1]).strip()
                attestation = cells[2].strip()
        if mdk_code and has_module_exam:
            break

    requirements = {"practice_obligatory": [], "skills_obligatory": [], "knowledge_obligatory": []}
    current = None
    for text in paragraphs:
        low = text.lower()
        if low.startswith("иметь практический опыт"):
            current = "practice_obligatory"
            continue
        if low == "уметь:":
            current = "skills_obligatory"
            continue
        if low == "знать:":
            current = "knowledge_obligatory"
            continue
        if current and low.startswith(("1.3", "1.4", "общие компетенции", "профессиональные компетенции")):
            current = None
            continue
        if current and text:
            requirements[current].append(text)

    if not any(requirements.values()):
        # Формат 15.01.37: требования в таблице компетенций (уметь / знать / владеть навыками)
        for table in doc.tables:
            header = row_texts(table.rows[0])
            if not header or not header[0].startswith("Код ОК"):
                continue
            for row in table.rows[1:]:
                cells = row_texts(row)
                if not cells or not cells[0].startswith("ПК 2."):
                    continue
                if len(cells) > 1:
                    requirements["skills_obligatory"] += split_sentences(cells[1])
                if len(cells) > 2:
                    requirements["knowledge_obligatory"] += split_sentences(cells[2])
                if len(cells) > 3 and cells[3] != "-":
                    requirements["practice_obligatory"] += split_sentences(cells[3])
            break

    # содержание МДК 02.01
    content = None
    for table in doc.tables:
        if table.rows and row_texts(table.rows[0])[0].startswith("Наименование разделов"):
            content = table
            break
    if content is None:
        raise SystemExit("Не найдена таблица 3.2 (содержание модуля) в РП")

    topics: list[dict] = []
    other_topics: list[dict] = []
    current_topic = None
    in_practice = False
    current_mdk = None
    for row in content.rows:
        texts = [t for t in row_texts(row) if t]
        if not texts:
            continue
        if texts[0].startswith("МДК 02.01"):
            current_mdk = "main"
            current_topic = None
            continue
        if texts[0].startswith(("МДК 02.02", "МДК 02.03", "МДК 02.04")):
            current_mdk = "other"
            current_topic = None
            continue
        if texts[0].startswith("Тема"):
            current_topic = {"name": texts[0], "theory": [], "practical": []}
            (topics if current_mdk == "main" else other_topics).append(current_topic)
            in_practice = False
            continue
        if current_topic is None:
            continue
        joined = " ".join(texts).lower()
        if (
            "практическ" in joined
            and ("заняти" in joined or "работ" in joined)
            and not re.match(r"^\d", texts[0])
        ):
            in_practice = True
            continue
        if re.fullmatch(r"\d+([-–]\d+)?", texts[0] or "") and len(texts) > 1:
            target = current_topic["practical"] if in_practice else current_topic["theory"]
            item = {"num": texts[0], "name": texts[1]} if in_practice else texts[1]
            target.append(item)

    return {
        "pm_name": pm_name,
        "profession": profession,
        "mdk_code": mdk_code,
        "mdk_name": mdk_name,
        "attestation": attestation,
        "requirements": requirements,
        "topics": topics,
        "other_topics": other_topics,
        "has_module_exam": has_module_exam,
    }


def parse_ktp(path: Path) -> dict:
    doc = Document(path)
    lessons, attestation = [], None
    for table in doc.tables:
        if not table.rows or not row_texts(table.rows[0])[0].startswith("№ занятия"):
            continue
        for row in table.rows:
            cells = row_texts(row)
            if len(cells) < 3 or not re.fullmatch(r"\d+(\s*-\s*\d+)?", cells[0] or ""):
                continue
            lessons.append({"num": cells[0], "name": cells[1]})
            if "зачет" in cells[1].lower():
                attestation = {"num": cells[0], "name": cells[1]}
    return {"lessons": lessons, "attestation": attestation}


def parse_docx(path: Path) -> dict:
    doc = Document(path)
    paragraphs = [re.sub(r"\s+", " ", p.text).strip() for p in doc.paragraphs]
    tables = [table_rows(table) for table in doc.tables]
    return {"paragraphs": paragraphs, "tables": tables}


def between(paragraphs: list[str], start: str, stop: str) -> list[str]:
    begin = next((i for i, t in enumerate(paragraphs) if norm(t).startswith(norm(start))), None)
    if begin is None:
        return []
    end = next(
        (i for i, t in enumerate(paragraphs[begin + 1 :], begin + 1) if norm(t).startswith(norm(stop))),
        len(paragraphs),
    )
    return [t for t in paragraphs[begin + 1 : end] if t]


def flat(table_rows_: list) -> str:
    chunks = []
    for row in table_rows_:
        for cell in row:
            chunks.extend(cell)
    return "\n".join(chunks)


# ------------------------------------------------------------------------- проверки


class Report:
    def __init__(self) -> None:
        self.items: list[tuple[str, str, str]] = []

    def add(self, status: str, title: str, detail: str = "") -> None:
        self.items.append((status, title, detail))

    @property
    def failed(self) -> int:
        return sum(1 for status, _, _ in self.items if status == FAIL)

    def print(self) -> None:
        for status, title, detail in self.items:
            print(f"{status} {title}" + (f" — {detail}" if detail else ""))
        ok = sum(1 for status, _, _ in self.items if status == PASS)
        warn = sum(1 for status, _, _ in self.items if status == WARN)
        print("\n" + "-" * 70)
        print(f"Итог: пройдено {ok}, провалов {self.failed}, предупреждений {warn}")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--rp", required=True)
    parser.add_argument("--ktp")
    parser.add_argument("--all-mdk", action="store_true", help="документ по модулю: все МДК РП считаются основными")
    args = parser.parse_args()

    doc = parse_docx(Path(args.docx))
    rp = parse_rp(Path(args.rp))
    if args.all_mdk:
        rp["topics"] = rp["topics"] + rp["other_topics"]
        rp["other_topics"] = []
    ktp = parse_ktp(Path(args.ktp)) if args.ktp else None
    report = Report()

    paragraphs = doc["paragraphs"]
    title_zone = " ".join(paragraphs[:40])

    works = next((t for t in doc["tables"] if t and t[0] and norm(t[0][0][0]).startswith("наименование тем")), [])
    control = next((t for t in doc["tables"] if t and t[0] and norm(t[0][0][0]).startswith("п п")), [])
    student_tables = [t for t in doc["tables"] if t and len(t[0]) >= 2 and norm(t[0][1][0]) == "фио"]

    # 1. титул
    if rp["mdk_name"] and norm(rp["mdk_name"]) in norm(title_zone):
        report.add(PASS, "Титул: наименование МДК совпадает с РП", rp["mdk_name"])
    elif rp["pm_name"] and norm(rp["pm_name"]) in norm(title_zone):
        report.add(PASS, "Титул: наименование ПМ совпадает с РП (документ по модулю)", rp["pm_name"][:80])
    else:
        report.add(FAIL, "Титул: наименование МДК/ПМ не совпадает с РП", f"в РП: {rp['mdk_name']!r}")

    if "пм 02" in norm(title_zone):
        if rp["pm_name"] and norm(rp["pm_name"]) in norm(title_zone):
            report.add(PASS, "Титул: наименование ПМ совпадает с РП")
        else:
            report.add(PASS, "Титул: шифр ПМ 02 указан")
    else:
        report.add(FAIL, "Титул: шифр ПМ 02 не найден")

    prof_code = re.search(r"\b(\d{2}\.\d{2}\.\d{2})\b", rp["profession"] or "")
    if prof_code and prof_code.group(1) in title_zone:
        report.add(PASS, "Титул: профессия совпадает с РП", prof_code.group(1))
    else:
        report.add(FAIL, "Титул: профессия не найдена", rp["profession"][:70])

    if "2026 2027" in norm(title_zone) and "2026" in title_zone:
        report.add(PASS, "Титул: учебный год 2026-2027 и год 2026")
    else:
        report.add(FAIL, "Титул: учебный год / год не совпадают")

    # 2. форма аттестации и номер занятия
    control_flat = flat(control)
    doc_form = " ".join(control[1][1]) if len(control) > 1 and len(control[1]) > 1 else ""
    rp_form = norm(rp["attestation"])
    form_norm = norm(doc_form)
    if rp_form and (rp_form in form_norm or form_norm in rp_form):
        report.add(PASS, "Контроль знаний: форма аттестации совпадает с РП", rp["attestation"])
    elif "экзамен" in form_norm and rp.get("has_module_exam", False):
        report.add(PASS, "Контроль знаний: экзамен по модулю (по РП)", doc_form)
    else:
        report.add(FAIL, "Контроль знаний: форма аттестации не совпадает с РП", f"в РП: {rp['attestation']!r}, в документе: {doc_form!r}")

    if ktp and ktp["attestation"]:
        num = ktp["attestation"]["num"]
        if "экзамен" in form_norm and "модул" in form_norm:
            if re.search(r"ПА\s*\.?\s*0?2", control_flat, re.IGNORECASE):
                report.add(PASS, "Контроль знаний: экзамен по модулю ПА.02 (номера занятия нет)")
            else:
                report.add(WARN, "Контроль знаний: не указан шифр ПА.02")
        elif re.search(rf"(?<![\d-]){re.escape(num)}(?![\d-])", control_flat):
            report.add(PASS, "Контроль знаний: номер занятия совпадает с КТП", f"занятие {num}")
        else:
            report.add(FAIL, "Контроль знаний: номер занятия не совпадает с КТП", f"в КТП {num}")

    # 3. темы и практические занятия
    works_flat = flat(works)
    required_topics = [t["name"] for t in rp["topics"] if t["practical"]]
    skipped_topics = [t["name"] for t in rp["topics"] if not t["practical"]]
    missing_topics = [name for name in required_topics if norm(name) not in norm(works_flat)]
    if missing_topics:
        report.add(
            FAIL,
            f"Перечень работ: не хватает тем с работами ({len(missing_topics)} из {len(required_topics)})",
            "; ".join(missing_topics),
        )
    else:
        report.add(PASS, f"Перечень работ: все темы с работами на месте ({len(required_topics)})")

    if skipped_topics:
        report.add(PASS, f"Перечень работ: темы без работ не печатаются ({len(skipped_topics)})", "; ".join(s[:60] for s in skipped_topics))

    rp_names = [t["name"] for t in rp["topics"]]
    doc_topics = [
        row[0][0] for row in works[1:]
        if row and row[0] and not row[0][0].startswith(("МДК", "Вид контроля"))
    ]
    extra_topics = [d for d in doc_topics if not any(similar(d, name) >= 0.85 for name in rp_names)]
    if extra_topics:
        report.add(WARN, f"Перечень работ: темы, которых нет в РП ({len(extra_topics)})", "; ".join(e[:60] for e in extra_topics))

    doc_practicals = []
    for row in works[1:]:
        if len(row) < 2:
            continue
        for line in row[1]:
            cleaned = re.sub(r"^Практическое занятие\s*(№\s*[\d\s\-–—]+)?\.?\s*", "", line).strip()
            if cleaned:
                doc_practicals.append(cleaned)

    rp_practicals = [
        re.sub(r"^Практическое занятие\.?\s*", "", p["name"]).strip()
        for t in rp["topics"]
        for p in t["practical"]
    ]
    missing = [
        name for name in rp_practicals
        if not any(similar(name, d) >= 0.85 or norm(name) in norm(d) for d in doc_practicals)
    ]
    if missing:
        report.add(
            FAIL,
            f"Перечень работ: не найдены практические занятия РП ({len(missing)} из {len(rp_practicals)})",
            "; ".join(m[:60] for m in missing),
        )
    else:
        report.add(PASS, f"Перечень работ: все практические занятия РП на месте ({len(rp_practicals)})")

    extra = [d for d in doc_practicals if not any(similar(d, r) >= 0.85 or norm(r) in norm(d) for r in rp_practicals)]
    if extra:
        report.add(WARN, f"Перечень работ: пункты, которых нет в РП ({len(extra)})", "; ".join(e[:60] for e in extra))
    else:
        report.add(PASS, "Перечень работ: лишних пунктов нет")

    # 4. требования к результатам
    req_zone = "\n".join(between(paragraphs, "Иметь практический опыт", "Перечень самостоятельных"))
    missing_req = []
    for key in ("practice_obligatory", "skills_obligatory", "knowledge_obligatory"):
        for item in rp["requirements"][key]:
            if norm(item) not in norm(req_zone):
                missing_req.append(item)
    total_req = sum(len(v) for v in rp["requirements"].values())
    if missing_req:
        report.add(FAIL, f"Требования: не хватает {len(missing_req)} из {total_req} пунктов РП", "; ".join(m[:60] for m in missing_req))
    else:
        report.add(PASS, f"Требования: все {total_req} пунктов РП перенесены дословно")

    # 5. вопросы
    questions = between(paragraphs, "Вопросы", "КРИТЕРИИ ОЦЕНКИ")
    questions = [re.sub(r"^\d+\.\s*", "", q) for q in questions]
    if not questions:
        report.add(WARN, "Вопросы к зачёту: блок отсутствует")
    else:
        if 20 <= len(questions) <= 32:
            report.add(PASS, "Вопросы к зачёту: количество в норме", f"{len(questions)} шт. (20–32)")
        else:
            report.add(FAIL, "Вопросы к зачёту: количество вне нормы", f"{len(questions)} шт.")

        chunks = [
            (t["name"], " ".join(t["theory"] + [p["name"] for p in t["practical"]]))
            for t in rp["topics"]
        ]
        other_chunks = [
            (t["name"], " ".join(t["theory"] + [p["name"] for p in t["practical"]]))
            for t in rp["other_topics"]
        ]
        unmatched, foreign, covered = [], [], set()
        for question in questions:
            words = significant(question)
            best_topic, best_hit = None, 0
            for name, text in chunks:
                hit = len(words & significant(f"{text} {name}"))
                if hit > best_hit:
                    best_topic, best_hit = name, hit
            other_hit = max(
                (len(words & significant(f"{text} {name}")) for name, text in other_chunks),
                default=0,
            )
            if best_hit < 2:
                unmatched.append(question)
            elif other_hit > best_hit:
                foreign.append(question)
            else:
                covered.add(best_topic)

        if unmatched:
            report.add(FAIL, f"Вопросы: не привязаны к содержанию МДК 02.01 ({len(unmatched)})", "; ".join(u[:60] for u in unmatched))
        else:
            report.add(PASS, "Вопросы: каждый опирается на содержание МДК 02.01")

        if foreign:
            report.add(
                FAIL,
                f"Вопросы: относятся к другим МДК, не к 02.01 ({len(foreign)})",
                "; ".join(u[:60] for u in foreign),
            )
        else:
            report.add(PASS, "Вопросы: только по МДК 02.01 (не 02.02/02.03)")

        uncovered = [t["name"] for t in rp["topics"] if t["name"] not in covered]
        if uncovered:
            report.add(WARN, f"Вопросы: темы без вопросов ({len(uncovered)})", "; ".join(u[:60] for u in uncovered))
        else:
            report.add(PASS, f"Вопросы: охвачены все темы РП ({len(rp['topics'])})")

    # 6. лист ознакомления
    groups_match = re.search(r"Группа:\s*(.+)", next((t for t in paragraphs if t.startswith("Группа")), ""))
    groups = [g.strip() for g in re.split(r"[,;]", groups_match.group(1))] if groups_match else []
    if student_tables:
        headings = [t for t in paragraphs if t.lower().startswith("студенты группы")]
        missing_groups = [g for g in groups if not any(g in h for h in headings)]
        if missing_groups:
            report.add(FAIL, "Лист ознакомления: нет листа для группы", "; ".join(missing_groups))
        else:
            report.add(PASS, f"Лист ознакомления: листы по всем группам ({len(groups)})", ", ".join(groups))

        table = student_tables[0]
        body = table[1:]
        empty = sum(1 for row in body if all(not cell for cell in row))
        if empty == 25:
            report.add(PASS, "Лист ознакомления: 25 пустых строк")
        else:
            report.add(WARN, "Лист ознакомления: нестандартное число строк", f"{len(body)} строк, пустых {empty}")
    else:
        report.add(FAIL, "Лист ознакомления: таблица не найдена")

    report.print()
    sys.exit(1 if report.failed else 0)


if __name__ == "__main__":
    main()
