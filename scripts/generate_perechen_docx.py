"""Генератор «Перечня материалов для подготовки к промежуточной аттестации».

Собирает .docx из JSON по блокам. На текущем этапе реализован только титул.
Каждый блок — отдельная функция build_<block>, чтобы можно было отлаживать
документ по частям и не пересобирать всё целиком.

Запуск:
    python generate_perechen_docx.py --data data_perechen_etalon.json \
        --out "../Перечни/перечень ЭТАЛОН.docx" --blocks title
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT = "Times New Roman"


def setup_document() -> Document:
    """A4, поля и базовый стиль как в образцах."""
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.3)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.3)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    return doc


def add_par(
    doc: Document,
    text: str = "",
    size: float | None = 12,
    bold: bool = False,
    align: str | None = "left",
    space_before: float = 0,
    space_after: float = 0,
    first_line_indent: float | None = None,
    left_indent: float | None = None,
) -> None:
    par = doc.add_paragraph()
    if align is not None:
        par.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }[align]
    pf = par.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    if text:
        run = par.add_run(text)
        run.font.name = FONT
        if size is not None:
            run.font.size = Pt(size)
        run.bold = bold
    return None


def add_empty(doc: Document, count: int) -> None:
    for _ in range(count):
        add_par(doc, "", None, align=None)


def build_title(doc: Document, data: dict) -> None:
    meta = data["meta"]
    approval = data["approval"]
    doc_kind = meta.get("doc_kind", "kdz")

    # Шапка — 12 pt
    add_par(doc, "МИНИСТЕРСТВО ОБРАЗОВАНИЯ РОСТОВСКОЙ ОБЛАСТИ", 12, align="center")
    add_par(doc, "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ РОСТОВСКОЙ ОБЛАСТИ", 12, align="center")
    add_par(doc, "«САЛЬСКИЙ ИНДУСТРИАЛЬНЫЙ ТЕХНИКУМ»", 12, align="center")
    add_empty(doc, 2)

    # УТВЕРЖДАЮ — сдвинут вправо отступом 11 см
    add_par(doc, "УТВЕРЖДАЮ", 12, align="center", space_after=2, left_indent=11.0)
    add_par(doc, "Зам. директора по УР", 12, align="center", left_indent=11.0)
    add_par(doc, f"____________ {approval['deputy_director']}", 12, align="center", left_indent=11.0)
    add_par(doc, f"«___» ___________ {meta['year_for_print']} г.", 12, align="center", left_indent=11.0)
    add_empty(doc, 4)

    add_par(doc, "ПЕРЕЧЕНЬ МАТЕРИАЛОВ", 16, bold=True, align="center", space_after=2)
    add_par(doc, "для подготовки студентов к промежуточной аттестации", 14, align="center")
    add_par(doc, f"в {meta['study_year']} учебном году", 14, align="center")
    add_empty(doc, 1)

    # ВАЖНО: названия ПМ и профессии берутся буква в букву из РП соответствующего ПМ
    # (`RP/02_Оригиналы/...`), а не из учебного плана: в УП профессия может
    # называться иначе (например, «Монтажник…» вместо «Слесарь-наладчик…»).
    if doc_kind == "diff":
        subject = f"{meta['mdk_code']} «{meta['mdk_title']}»"
    else:
        subject = f"{meta['pm_code']} «{meta['pm_title']}»"
    add_par(doc, f"по профессиональному модулю {subject}", 12, align="justify", first_line_indent=1.25)
    add_par(
        doc,
        f"для профессии: {meta['profession_code']} {meta['profession_title']}",
        12,
        align="justify",
        first_line_indent=1.25,
    )
    add_par(doc, "", 12, align="justify", first_line_indent=1.25)
    add_par(doc, "", 12, align="justify", first_line_indent=1.25)

    add_par(doc, f"Группа: {', '.join(meta['groups'])}", 12, align=None)
    add_par(doc, f"Преподаватель: {', '.join(meta['teachers'])}", 12, align=None)
    add_empty(doc, 4)

    add_par(doc, "Рассмотрено на заседании цикловой комиссии технических дисциплин", 12, align=None)
    add_par(doc, f"протокол № ___ от «___» ___________ {meta['year_for_print']} г.", 12, align=None)
    add_par(doc, f"председатель ______________ /{approval['commission_chair']}/", 12, align=None)
    add_empty(doc, 18)

    add_par(doc, "г. Сальск", 12, align="center")
    add_par(doc, f"{meta['year_for_print']} г.", 12, align="center")


def build_control(doc: Document, data: dict) -> None:
    """Блок «Контроль знаний» — открывает вторую страницу."""
    heading = doc.add_paragraph()
    heading.paragraph_format.page_break_before = True
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run("Контроль знаний")
    run.font.name = FONT
    run.font.size = Pt(12)
    add_empty(doc, 1)

    rows = data.get("control", [])
    table = doc.add_table(rows=1 + len(rows), cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = [Cm(1.7), Cm(7.5), Cm(5.0), Cm(3.0)]
    headers = [
        "№п/п",
        "Контрольные работы, дифференцированный зачет, предусмотренные учебным планом.",
        "Раздел, тема.",
        "Сроки проведения (№ урока).",
    ]
    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = widths[i]
        par = cell.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(text)
        r.font.name = FONT
        r.font.size = Pt(12)

    for n, item in enumerate(rows, start=1):
        cells = table.rows[n].cells
        values = [str(n), item["form"], item["section"], item["lesson"]]
        for i, value in enumerate(values):
            cells[i].width = widths[i]
            par = cells[i].paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 3) else WD_ALIGN_PARAGRAPH.LEFT
            r = par.add_run(value)
            r.font.name = FONT
            r.font.size = Pt(12)

    add_empty(doc, 1)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--blocks", default="title")
    args = parser.parse_args()

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    doc = setup_document()

    blocks = [b.strip() for b in args.blocks.split(",") if b.strip()]
    for block in blocks:
        if block == "title":
            build_title(doc, data)
        elif block == "control":
            build_control(doc, data)
        else:
            raise SystemExit(f"Блок ещё не реализован: {block}")

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"Сохранено: {out}")


if __name__ == "__main__":
    main()
