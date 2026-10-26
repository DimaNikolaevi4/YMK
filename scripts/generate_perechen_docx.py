"""Генератор «Перечня материалов для подготовки к промежуточной аттестации».

Собирает .docx из JSON по блокам. Каждый блок — отдельная функция build_<block>,
чтобы можно было отлаживать документ по частям и не пересобирать всё целиком.

Блоки (порядок в документе):
    title, control, requirements, works, course_project, questions, criteria,
    student_lists

Запуск:
    python scripts/generate_perechen_docx.py \
        --data scripts/data_perechen_etalon.json \
        --out "Перечни/01_Готовые/перечень МДК 02.01 КИП-31.docx" --blocks all

Отдельные блоки:
    --blocks title,control     только титул и таблица «Контроль знаний»
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT = "Times New Roman"

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# Подсекции блока «Требования к результатам освоения»
REQUIREMENTS_ORDER = [
    ("practice_obligatory", "Иметь практический опыт (обязательная часть):"),
    ("skills_obligatory", "Должен уметь (обязательная часть):"),
    ("knowledge_obligatory", "Должен знать (обязательная часть):"),
    ("practice_variative", "Иметь практический опыт (вариативная часть):"),
    ("skills_variative", "Уметь (вариативная часть):"),
    ("knowledge_variative", "Знать (вариативная часть):"),
]

# Готовые шаблоны критериев: можно указать строкой "short"/"extended" вместо текста
CRITERIA_PRESETS = {
    "short": {
        "title": "КРИТЕРИИ ОЦЕНКИ ЗНАНИЙ СТУДЕНТОВ",
        "items": [
            "5 «отлично» - заслуживает студент, обнаруживший всестороннее, систематическое и "
            "глубокое знание учебного программного материала, проявивший творческие способности "
            "и научный подход в понимании и изложении учебного материала, который представлен "
            "последовательно и логично.",
            "4 «хорошо» - заслуживает студент, обнаруживший достаточно полное знание "
            "учебно-программного материала, не допускающий в ответе существенных неточностей, "
            "усвоивший основную литературу, рекомендованную программой.",
            "3 «удовлетворительно» - заслуживает студент, обнаруживший знание основного "
            "учебно-программного материала в объеме, необходимом для дальнейшей учебы, однако "
            "допустивший погрешности при ответе, но обладающий необходимыми знаниями для "
            "устранения наиболее существенных из них.",
            "2 «неудовлетворительно»- выставляется студенту, обнаружившему пробелы в знаниях или "
            "отсутствие знания по значительной части основного учебно-программного материала, "
            "допустившему принципиальные ошибки при решении задач, который не может продолжить "
            "обучение без дополнительных занятий.",
        ],
    },
    "extended": {
        "title": "КРИТЕРИИ ОЦЕНКИ ЗНАНИЙ СТУДЕНТОВ",
        "items": [
            "Оценка – это определение степени усвоения обучаемыми знаний, умений и навыков в "
            "соответствии с требованиями программ обучения и руководящими документами обучения.",
            "«ОТЛИЧНО» - ставится в том случае, если обучаемый глубоко изучил учебный материал и "
            "литературу по проблеме, последовательно и исчерпывающе отвечает на поставленные "
            "вопросы, а при выполнении практической работы – если задание выполнено правильно и "
            "в установленное нормативом время (при отсутствии нормативов – уверенно и быстро).",
            "«ХОРОШО» - ставится тогда, когда обучаемый твердо знает материал и отвечает без "
            "наводящих вопросов, разбирается в литературе по проблеме, а при выполнении "
            "практической работы – если задание выполнено правильно.",
            "«УДОВЛЕТВОРИТЕЛЬНО» - ставится при условии, если обучаемый знает лишь основной "
            "материал, путается в литературе по проблеме, а на заданные вопросы отвечает "
            "недостаточно четко и полно, а при выполнении практической работы – если задание "
            "выполнено, но допускались ошибки, не отразившиеся на качестве выполненной работы.",
            "«НЕУДОВЛЕТВОРИТЕЛЬНО» - ставится в том случае, когда обучаемый не смог достаточно "
            "полно и правильно ответить на поставленные вопросы, не знает литературы по проблеме, "
            "а при выполнении практической работы – если задание не сделано или допущены ошибки, "
            "влияющие на качество выполненной работы.",
            "Требования к оценке:",
            "- оценка должна быть объективной и справедливой, ясной и понятной для обучаемого,",
            "- оценка должна выполнять стимулирующую функцию,",
            "- оценка должна быть всесторонней.",
            "При оценке знаний учитывается:",
            "- объем знаний по учебному предмету (вопросу),",
            "- понимание изученного, самостоятельность суждений, убежденность в излагаемом,",
            "- степень систематизации и глубины знаний,",
            "- действенность знаний, умение применять их с целью решения практических задач.",
            "При оценке навыков и умений учитываются:",
            "- содержание навыков и умений,",
            "- точность, прочность, гибкость навыков и умений,",
            "- возможность применять навыки и умения на практике,",
            "- наличие ошибок, их количество, характер и влияние на работу.",
        ],
    },
}


# ---------------------------------------------------------------- каркас документа


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
    page_break_before: bool = False,
) -> None:
    par = doc.add_paragraph()
    if page_break_before:
        par.paragraph_format.page_break_before = True
    if align is not None:
        par.alignment = ALIGN[align]
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


def set_cell(cell, lines, size=12, bold=False, align="left", valign="top") -> None:
    """Записать в ячейку список строк (каждая — отдельный абзац)."""
    if isinstance(lines, str):
        lines = [lines]
    cell.vertical_alignment = {
        "top": WD_ALIGN_VERTICAL.TOP,
        "center": WD_ALIGN_VERTICAL.CENTER,
    }[valign]

    first = True
    for line in lines:
        par = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        par.alignment = ALIGN[align]
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        run = par.add_run(line)
        run.font.name = FONT
        run.font.size = Pt(size)
        run.bold = bold


def add_table(doc: Document, rows: int, cols: int, widths: list[float]):
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for idx, width in enumerate(widths):
        table.columns[idx].width = Cm(width)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Cm(widths[idx])
    return table


# ------------------------------------------------------------------------ блоки


def build_title(doc: Document, data: dict) -> None:
    meta = data["meta"]
    approval = data["approval"]
    doc_kind = meta.get("doc_kind", "kdz")

    # Шапка — 12 pt
    add_par(doc, "МИНИСТЕРСТВО ОБРАЗОВАНИЯ РОСТОВСКОЙ ОБЛАСТИ", 12, align="center")
    add_par(
        doc,
        "ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ "
        "РОСТОВСКОЙ ОБЛАСТИ",
        12,
        align="center",
    )
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
    if doc_kind == "kdz":
        subject = f"{meta['pm_code']} «{meta['pm_title']}»"
    elif doc_kind == "mdk":
        subject = f"{meta['pm_code']} {meta['mdk_code']} «{meta['mdk_title']}»"
    else:
        subject = f"{meta['mdk_code']} «{meta['mdk_title']}»"
    add_par(doc, f"по профессиональному модулю {subject}", 12, align="justify", first_line_indent=1.25)

    if meta.get("education_kind") == "специальность":
        add_par(
            doc,
            f"специальность (код, название): {meta['education_code']} «{meta['education_title']}»",
            12,
            align="justify",
            first_line_indent=1.25,
        )
    elif "profession_code" in meta:
        add_par(
            doc,
            f"для профессии: {meta['profession_code']} {meta['profession_title']}",
            12,
            align="justify",
            first_line_indent=1.25,
        )

    add_par(doc, "", 12, align="justify", first_line_indent=1.25)
    add_par(doc, "", 12, align="justify", first_line_indent=1.25)

    if meta.get("course"):
        add_par(doc, f"курс: {meta['course']} курс", 12, align=None)
    add_par(doc, f"Группа: {', '.join(meta['groups'])}", 12, align=None)
    add_par(doc, f"Преподаватель: {', '.join(meta['teachers'])}", 12, align=None)
    add_empty(doc, 3)

    add_par(doc, "Рассмотрено на заседании цикловой комиссии технических дисциплин", 12, align=None)
    add_par(doc, f"протокол № ___ от «___» ___________ {meta['year_for_print']} г.", 12, align=None)
    add_par(doc, f"председатель ______________ /{approval['commission_chair']}/", 12, align=None)
    add_empty(doc, 18)

    add_par(doc, "г. Сальск", 12, align="center")
    add_par(doc, f"{meta['year_for_print']} г.", 12, align="center")


def build_control(doc: Document, data: dict) -> None:
    """Блок «Контроль знаний» — открывает вторую страницу."""
    add_par(doc, "Контроль знаний", 12, bold=True, align="center", page_break_before=True)
    add_empty(doc, 1)

    rows = data.get("control", [])
    widths = [1.4, 7.3, 4.3, 4.0]
    table = add_table(doc, 1 + len(rows), 4, widths)

    headers = [
        "№п/п",
        "Контрольные работы, дифференцированный зачет, предусмотренные учебным планом.",
        "Раздел, тема.",
        "Сроки проведения (№ урока).",
    ]
    for i, text in enumerate(headers):
        set_cell(table.rows[0].cells[i], text, bold=True, align="center", valign="center")

    for n, item in enumerate(rows, start=1):
        cells = table.rows[n].cells
        form = item["form"].split("\n") if item.get("form") else [""]
        set_cell(cells[0], f"{n}.", align="center", valign="center")
        set_cell(cells[1], form)
        set_cell(cells[2], item.get("section", ""))
        set_cell(cells[3], item.get("lesson", ""), align="center", valign="center")

    add_empty(doc, 1)


def build_requirements(doc: Document, data: dict) -> None:
    """Блок «Требования к результатам освоения»."""
    req = data.get("requirements")
    if not req:
        return

    if req.get("heading"):
        add_par(doc, req["heading"], 12, bold=True, align="center", space_after=2)
    if req.get("intro"):
        add_par(doc, req["intro"], 12, align="justify", first_line_indent=1.25)

    for key, label in REQUIREMENTS_ORDER:
        items = req.get(key) or []
        if not items:
            continue
        add_par(doc, label, 12, bold=True, align="justify", first_line_indent=1.25)
        for item in items:
            add_par(doc, item, 12, align="justify", first_line_indent=1.25)

    add_empty(doc, 1)


def column_widths(count: int) -> list[float]:
    """Ширины колонок таблицы работ под фактический набор столбцов (см)."""
    if count == 4:
        return [4.3, 4.1, 4.4, 4.7]
    if count == 3:
        return [5.4, 6.0, 5.9]
    if count == 2:
        return [6.4, 10.8]
    return [17.2]


def build_works(doc: Document, data: dict) -> None:
    """Таблица работ. Состав столбцов и заголовок — по фактическим данным:
    нет лабораторных — нет столбца и слова в заголовке; темы без работ не печатаются.
    """
    works = data.get("works")
    if not works:
        return

    topics = works.get("topics") or []
    flags = [
        ("practical", "практических", "Тема практической работы"),
        ("lab", "лабораторных", "Тема лабораторной работы"),
        ("independent", "самостоятельных", "Тема самостоятельной работы"),
    ]
    flags = [f for f in flags if any(t.get(f[0]) for t in topics)]

    columns = ["Наименование тем"] + [f[2] for f in flags]
    available = {f[0] for f in flags}
    heading_order = [
        ("independent", "самостоятельных"),
        ("practical", "практических"),
        ("lab", "лабораторных"),
    ]
    parts = [title for key, title in heading_order if key in available]
    if parts:
        joined = parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " и " + parts[-1]
        heading = f"Перечень {joined} работ"
    else:
        heading = "Перечень работ"

    visible = [
        t for t in topics
        if t.get("header") or any(t.get(f[0]) for f in flags)
    ]

    add_par(doc, heading, 12, bold=True, align="center", page_break_before=True)
    add_empty(doc, 1)

    footer = works.get("footer") or {}
    rows = 1 + (1 if works.get("mdk_header") else 0) + len(visible) + (1 if footer.get("topic") else 0)
    widths = column_widths(len(columns))
    table = add_table(doc, rows, len(columns), widths)

    for i, text in enumerate(columns):
        set_cell(table.rows[0].cells[i], text, bold=True, align="center", valign="center")

    row_index = 1
    if works.get("mdk_header"):
        merged = table.cell(row_index, 0).merge(table.cell(row_index, len(columns) - 1))
        set_cell(merged, works["mdk_header"], align="center")
        row_index += 1

    for topic in visible:
        if topic.get("header"):
            merged = table.cell(row_index, 0).merge(table.cell(row_index, len(columns) - 1))
            set_cell(merged, topic["header"], align="center")
            row_index += 1
            continue
        cells = table.rows[row_index].cells
        set_cell(cells[0], topic.get("topic", ""))
        for idx, (key, _, _) in enumerate(flags, start=1):
            set_cell(cells[idx], topic.get(key) or "")
        row_index += 1

    if footer.get("topic"):
        merged = table.cell(row_index, 0).merge(table.cell(row_index, len(columns) - 1))
        set_cell(merged, [footer["topic"]] + ([footer["text"]] if footer.get("text") else []), align="center")

    add_empty(doc, 1)
    print(
        f"Таблица работ: «{heading}», тем {len(visible)} из {len(topics)}, "
        f"столбцов {len(columns)} (пропущено тем без работ: {len(topics) - len(visible)})"
    )


def build_course_project(doc: Document, data: dict) -> None:
    """Блок «Курсовое проектирование» — только если course_project не null."""
    project = data.get("course_project")
    if not project:
        return

    add_empty(doc, 2)
    add_par(doc, project["title"], 12, bold=True, align="center")
    for line in project.get("lead") or []:
        add_par(doc, line, 12, bold=True, align="center")
    for line in project.get("intro_center") or []:
        add_par(doc, line, 12, align="center")
    for line in project.get("intro") or []:
        add_par(doc, line, 12, align=None)
    if project.get("scope_note"):
        add_par(doc, project["scope_note"], 12, align="justify")
    if project.get("input_heading"):
        add_par(doc, project["input_heading"], 12, bold=True, align="justify")
    for line in project.get("input_data") or []:
        add_par(doc, f"- {line}", 12, align="justify")

    for part in project.get("parts") or []:
        add_par(doc, part["title"], 12, bold=True, align="justify", space_before=2)
        for line in part.get("items") or []:
            add_par(doc, line, 12, align="justify")

    if project.get("footer"):
        add_empty(doc, 2)
        for line in project["footer"]:
            add_par(doc, line, 12, align=None)


def build_questions(doc: Document, data: dict) -> None:
    """Блок «Вопросы к зачёту/экзамену» — только если exam_questions не пуст.

    Единое оформление: номер, табуляция до 1,25 см, текст в одну колонку;
    висячий отступ — переносы строк выровнены по тексту, а не по номеру.
    """
    questions = data.get("exam_questions")
    if not questions:
        return
    if isinstance(questions, dict):
        title = questions.get("title", "Вопросы к дифференцированному зачету")
        items = questions.get("items") or []
    else:
        title = "Вопросы к дифференцированному зачету"
        items = questions

    add_par(doc, title, 12, bold=True, align="center", page_break_before=True)
    add_empty(doc, 1)
    for n, item in enumerate(items, start=1):
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = par.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Cm(1.25)
        pf.first_line_indent = Cm(-1.25)
        pf.tab_stops.add_tab_stop(Cm(1.25))
        run = par.add_run(f"{n}.\t{item}")
        run.font.name = FONT
        run.font.size = Pt(12)
    add_empty(doc, 1)


def resolve_criteria(criteria):
    """Критерии: строка-пресет («short»/«extended») или готовый объект."""
    if isinstance(criteria, str):
        preset = CRITERIA_PRESETS.get(criteria)
        if preset is None:
            raise SystemExit(f"Неизвестный пресет критериев: {criteria}")
        return preset
    return criteria


def build_criteria(doc: Document, data: dict) -> None:
    """Блок «Критерии оценки знаний студентов»."""
    criteria = resolve_criteria(data.get("criteria"))
    if not criteria:
        return

    add_par(doc, criteria.get("title", "КРИТЕРИИ ОЦЕНКИ ЗНАНИЙ СТУДЕНТОВ"),
            12, bold=True, align="center", page_break_before=True)
    add_empty(doc, 1)
    for item in criteria.get("items") or []:
        add_par(doc, item, 12, align="justify", first_line_indent=1.25)
    add_empty(doc, 1)


def build_student_lists(doc: Document, data: dict) -> None:
    """Листы ознакомления — по таблице на группу, всегда пустые строки."""
    groups = data.get("student_lists") or []
    if not groups:
        return

    empty_rows = int(data.get("student_list_empty_rows", 25))
    widths = [1.0, 7.8, 4.1, 5.1]

    for index, group in enumerate(groups):
        add_par(
            doc,
            f"Студенты группы {group['group']} ознакомлены с материалами "
            "промежуточной аттестации:",
            12,
            align="center",
            page_break_before=True,
        )
        add_empty(doc, 1)

        table = add_table(doc, 1 + empty_rows, 4, widths)
        for i, text in enumerate(["№", "ФИО", "Дата", "Подпись"]):
            set_cell(table.rows[0].cells[i], text, bold=True, align="center", valign="center")
        for row in table.rows[1:]:
            for cell in row.cells:
                set_cell(cell, "", align="center")

        add_empty(doc, 1)
        add_par(doc, f"Преподаватель ____________________ {group.get('teacher_sign', '')}",
                12, align="center")
        if index < len(groups) - 1:
            add_empty(doc, 2)


BLOCKS = {
    "title": build_title,
    "control": build_control,
    "requirements": build_requirements,
    "works": build_works,
    "course_project": build_course_project,
    "questions": build_questions,
    "criteria": build_criteria,
    "student_lists": build_student_lists,
}
BLOCK_ORDER = list(BLOCKS)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--blocks", default="all", help="all или список через запятую")
    args = parser.parse_args()

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    doc = setup_document()

    raw = [b.strip() for b in args.blocks.split(",") if b.strip()]
    blocks = BLOCK_ORDER if raw == ["all"] else raw
    unknown = [b for b in blocks if b not in BLOCKS]
    if unknown:
        raise SystemExit(f"Неизвестный блок: {', '.join(unknown)}. Доступны: {', '.join(BLOCK_ORDER)}")

    for block in blocks:
        BLOCKS[block](doc, data)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"Сохранено: {out}")


if __name__ == "__main__":
    main()
