# -*- coding: utf-8 -*-
"""Компактный генератор практических работ МДК 02.01 (один лист А4 у студента).

Отличия от generate_practical_docx.py:
  * плотная вёрстка (таблицы 9 pt, интервал 1,0 в ячейках) — раздатка 4–5 листов;
  * блоки docx-содержимого берутся из того же data-JSON (type: p | table).

Использование:
    python scripts/generate_practical_compact.py --data scripts/data_practical_0201_s21.json --work 2
"""
import argparse
import json
import os
import sys

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

FONT = "Times New Roman"
BODY = 12
TBL = 9


def setup(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(BODY)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = st.paragraph_format
    pf.line_spacing = 1.15
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin, s.right_margin = Cm(3.0), Cm(1.5)
    s.top_margin, s.bottom_margin = Cm(2.0), Cm(2.0)


def para(doc, text="", *, bold=False, size=BODY, indent=True, space_after=0):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = FONT
    return p


def mixed(doc, bold_part, rest):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    r1 = p.add_run(bold_part)
    r1.bold = True
    r1.font.size = Pt(BODY)
    r1.font.name = FONT
    r2 = p.add_run(rest)
    r2.font.size = Pt(BODY)
    r2.font.name = FONT
    return p


def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(BODY)
    r.font.name = FONT


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(BODY)
    r.font.name = FONT


def _widths(headers, rows, total=15.5):
    """Ширины колонок пропорционально объёму текста (чтобы высокие ячейки не раздували лист)."""
    n = len(headers)
    weights = []
    for j in range(n):
        lens = [len(str(r[j])) for r in rows if j < len(r)] + [len(str(headers[j]))]
        avg = sum(lens) / max(1, len(lens))
        weights.append(max(avg, 4) ** 0.6)
    k = total / sum(weights)
    return [w * k for w in weights]


def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    if widths is None:
        widths = _widths(headers, rows)
    for i, row in enumerate([headers] + rows):
        for j in range(len(headers)):
            val = row[j] if j < len(row) else ""
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(str(val))
            r.bold = i == 0
            r.font.size = Pt(TBL)
            r.font.name = FONT
    para(doc, "", indent=False, space_after=2)
    return t


def add_blocks(doc, blocks):
    for b in blocks:
        if b.get("type") == "table":
            if b.get("caption"):
                caption(doc, b["caption"])
            table(doc, b["headers"], b["rows"])
        elif b.get("text"):
            para(doc, b["text"])


def build_work(work, meta):
    doc = Document()
    setup(doc)

    para(doc, "ПРАКТИЧЕСКАЯ РАБОТА № {}".format(work["number"]), bold=True, size=14, indent=False)
    mixed(doc, meta.get("pm_code", "ПМ.02") + " ", "«{}»".format(meta["pm_name"]))
    mixed(doc, meta.get("mdk_label", "МДК") + " " + meta["mdk_code"] + " ", "«{}»".format(meta["mdk_name"]))
    mixed(doc, "Профессия ", meta["profession"])
    mixed(doc, "Группы {} ".format(meta["groups"]),
          "· {} курс · {} семестр".format(meta["course"], meta["semester"]))
    para(doc, "Количество часов: {}".format(work["hours"]), bold=True, indent=False)

    mixed(doc, "Тема: ", work["topic"])
    mixed(doc, "Цель работы: ", work["goal"])
    mixed(doc, "Оборудование: ", work.get("equipment", "письменные принадлежности, листы А4."))

    heading(doc, "Краткая теория")
    for t in work.get("theory", []):
        para(doc, t)

    heading(doc, "Пример оформления (заполнения) документа")
    add_blocks(doc, work.get("sample", []))

    heading(doc, "Варианты задания (исходные данные)")
    add_blocks(doc, work.get("variants", []))

    heading(doc, "Ход работы")
    for s in work.get("steps", []):
        para(doc, s)

    heading(doc, "Содержание отчёта")
    for s in work.get("report", []):
        para(doc, s)

    heading(doc, "Контрольные вопросы")
    for q in work.get("questions", []):
        para(doc, q)

    heading(doc, "Литература")
    for i, lit in enumerate(work.get("literature", []), start=1):
        para(doc, "{}. {}".format(i, lit))

    return doc


def file_name(meta, work):
    return "Практическая работа № {} {} {}.docx".format(
        work["number"], meta.get("mdk_label", "МДК") + " " + meta["mdk_code"],
        meta["filename_suffix"])


def main():
    ap = argparse.ArgumentParser(description="Компактный генератор практических работ")
    ap.add_argument("--data", required=True)
    ap.add_argument("--work", type=int, default=None)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--outdir", default="Практические/01_Готовые")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8") as f:
        meta = json.load(f)

    works = meta["works"]
    if args.all:
        targets = works
    elif args.work is not None:
        targets = [w for w in works if w["number"] == args.work]
        if not targets:
            sys.exit("Работа № {} не найдена".format(args.work))
    else:
        ap.error("Укажите --work N или --all")

    os.makedirs(args.outdir, exist_ok=True)
    for w in targets:
        if not w.get("sample") and not w.get("variants"):
            sys.exit("SKIP {}: работа не заполнена".format(w["number"]))
        doc = build_work(w, meta)
        out = os.path.join(args.outdir, file_name(meta, w))
        doc.save(out)
        print("OK:", out)


if __name__ == "__main__":
    main()
