#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор практических работ (.docx) — направление «ПРАКТИЧЕСКИЕ».

Методика: ПРАКТИЧЕСКИЕ_INSTRUCTIONS.md (структура §2, оформление §3).
Данные: JSON формата scripts/data_practical_0201_s21.json (см. §6 методики).

CLI:
    python generate_practical_docx.py --data scripts/data_practical_0201_s21.json --work 1
    python generate_practical_docx.py --data ... --all
    python generate_practical_docx.py --data ... --work 1 --outdir "Практические/01_Готовые"

Зависимости: python-docx (stdlib +
python-docx). Хардкодов путей нет.
"""

import argparse
import json
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Pt

# ---------------------------------------------------------------- оформление
FONT = "Times New Roman"
BODY_PT = 12
HEAD_PT = 14


def _set_defaults(doc):
    """Times New Roman 12pt по умолчанию + поля по методике §3."""
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(BODY_PT)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15

    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)


def _para(doc, text, *, center=False, bold=False, size=BODY_PT, indent=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if center:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pf.first_line_indent = Cm(1.25) if indent else None
    run = p.add_run(text)
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)
    return p


def _heading(doc, text):
    """Заголовок блока: полужирный, без отступа, без нумерации."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = None
    run = p.add_run(text)
    run.bold = True
    run.font.name = FONT
    run.font.size = Pt(BODY_PT)
    return p


def build_work(doc, work, meta):
    """Собирает один документ практической работы в doc."""
    _set_defaults(doc)

    # Шапка (без титульного листа)
    _para(doc, "ПРАКТИЧЕСКАЯ РАБОТА № {}".format(work["number"]),
          center=True, bold=True, size=HEAD_PT)
    _para(doc, "МДК {}".format(meta["mdk_code"]), center=True)
    _para(doc, "«{}»".format(meta["mdk_name"]), center=True)
    _para(doc, "ПМ.02 «{}»".format(meta["pm_name"]), center=True)
    _para(doc, "Профессия {}".format(meta["profession"]), center=True)
    _para(doc, "Группы {} · {} курс · {} семестр".format(
        meta["groups"], meta["course"], meta["semester"]), center=True)
    _para(doc, "Количество часов: {}".format(work["hours"]),
          center=True, bold=True)
    doc.add_paragraph()

    blocks = [
        ("Тема:", [work["topic"]], "inline"),
        ("Цель работы:", [work["goal"]], "inline"),
        ("Оборудование:", [work["equipment"]], "inline"),
        ("Краткая теория:", work["theory"], "heading"),
        ("Ход работы:", work["steps"], "heading"),
        ("Содержание отчета:", work["report"], "heading"),
        ("Контрольные вопросы:", work["questions"], "heading"),
        ("Литература:", work["literature"], "heading"),
    ]
    for title, items, kind in blocks:
        if kind == "heading":
            _heading(doc, title)
            for i, item in enumerate(items, start=1):
                if not item:
                    continue
                if title == "Ход работы:":
                    _para(doc, "{}. {}".format(i, item))
                else:
                    _para(doc, item)
        else:  # inline: «Метка: текст» одной строкой, метка полужирная
            text = title + " " + (items[0] if items else "")
            _para(doc, text, indent=False)
            p = doc.paragraphs[-1]
            p.runs[0].bold = True
            p.runs[0].text = title + " "
            rest = p.add_run(items[0] if items else "")
            rest.bold = False
            rest.font.name = FONT
            rest.font.size = Pt(BODY_PT)
        if not any(items):
            raise ValueError(
                "Работа № {}: блок «{}» пуст — заполните data-JSON".format(
                    work["number"], title))
    return doc


def file_name(meta, work):
    return "Практическая работа № {} {} {}.docx".format(
        work["number"], meta.get("mdk_label", "МДК") + " " + meta["mdk_code"],
        meta["filename_suffix"])


def main():
    ap = argparse.ArgumentParser(description="Генератор практических работ")
    ap.add_argument("--data", required=True, help="Путь к data-JSON")
    ap.add_argument("--work", type=int, default=None, help="Номер работы")
    ap.add_argument("--all", action="store_true", help="Сгенерировать все работы")
    ap.add_argument("--outdir", default="Практические/01_Готовые",
                    help="Каталог вывода")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8") as f:
        meta = json.load(f)

    works = meta["works"]
    if args.all:
        targets = works
    elif args.work is not None:
        targets = [w for w in works if w["number"] == args.work]
        if not targets:
            sys.exit("Работа № {} не найдена в {}".format(args.work, args.data))
    else:
        ap.error("Укажите --work N или --all")

    os.makedirs(args.outdir, exist_ok=True)
    for work in targets:
        doc = Document()
        try:
            build_work(doc, work, meta)
        except ValueError as e:
            sys.exit("SKIP {}: {}".format(work["number"], e))
        out = os.path.join(args.outdir, file_name(meta, work))
        doc.save(out)
        print("OK:", out)


if __name__ == "__main__":
    main()
