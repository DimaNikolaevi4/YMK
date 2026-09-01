#!/usr/bin/env python3
"""
Скрипт генерации КТП в формате DOCX.

Использование:
    python3 scripts/generate_ktp_docx.py --data data.json --output KTP/МДК_03_01_КТП.docx

Формат data.json (подготавливается ИИ на основе РП и logic.md):
{
  "mdk_code": "03.01",
  "mdk_name": "Технология эксплуатации контрольно-измерительных приборов...",
  "pm_code": "03",
  "pm_name": "Техническое обслуживание и эксплуатация...",
  "profession": "15.01.31 Мастер контрольно-измерительных приборов и автоматики",
  "semester": "__",
  "course": "__",
  "group": "_________",
  "year": "20__-20__",
  "total_hours": 129,
  "practice_hours": 107,
  "with_teacher_hours": 129,
  "theory_hours": 22,
  "practical_hours": 107,
  "lab_hours": 0,
  "coursework_hours": 0,
  "independent_hours": 0,
  "exam_form": "Дифференцированный зачет",
  "commission_chair": "Ткаченко А.Н.",
  "director": "Т.В. Якимова",
  "topics": [
    {
      "name": "Тема 1.1 Основы метрологии и поверка средств измерений",
      "total_hours": 4,
      "practice_hours": 4,
      "lessons": [
        {"name": "...", "hours": 2, "is_practice": false, "is_lab": false, "practice_prep": 0,
         "equipment": "ОИ1, ДИ1", "task": "Стр.5-10", "note": "Вх.контр"},
        ...
      ]
    },
    ...
  ],
  "equipment": ["электроизмерительные приборы", ...],
  "sources_basic": [
    {"code": "ОИ 1", "name": "...", "author": "...", "publisher": "..."},
    ...
  ],
  "sources_additional": [
    {"code": "ДИ 1", "name": "...", "author": "...", "publisher": "..."},
    ...
  ]
}
"""

import argparse
import json
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def set_cell_shading(cell, color):
    """Установить цвет заливки ячейки."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_text(cell, text, bold=False, size=9, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """Установить текст в ячейке с заданным форматированием."""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = alignment
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    # Убираем отступы
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)


def add_paragraph(doc, text, bold=False, size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=6):
    """Добавить абзац с заданным форматированием."""
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    return p


def generate_ktp(data, output_path):
    doc = Document()

    # ========== МАРГИНЫ ==========
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(1.5)

    # ========== ТИТУЛЬНЫЙ ЛИСТ ==========
    add_paragraph(doc, 'МИНИСТЕРСТВО ОБЩЕГО И ПРОФЕССИОНАЛЬНОГО ОБРАЗОВАНИЯ РОСТОВСКОЙ ОБЛАСТИ', bold=True, size=11)
    add_paragraph(doc, '')
    add_paragraph(doc, 'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ РОСТОВСКОЙ ОБЛАСТИ', bold=True, size=11)
    add_paragraph(doc, '«САЛЬСКИЙ ИНДУСТРИАЛЬНЫЙ ТЕХНИКУМ»', bold=True, size=12)
    add_paragraph(doc, '(ГБПОУ РО «СИТ»)', size=11)
    add_paragraph(doc, '')
    add_paragraph(doc, '')

    # Таблица утверждения
    t = doc.add_table(rows=3, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    set_cell_text(t.cell(0, 1), 'УТВЕРЖДАЮ', bold=True, size=10)
    set_cell_text(t.cell(1, 1), 'Зам. директора ГБПОУ РО «СИТ»', size=10)
    set_cell_text(t.cell(1, 2), f'____________ {data.get("director", "Т.В. Якимова")}', size=10)
    set_cell_text(t.cell(2, 1), '«____» ______________ 20 ____ г.', size=10)
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)

    add_paragraph(doc, '')
    add_paragraph(doc, 'КАЛЕНДАРНО-ТЕМАТИЧЕСКИЙ ПЛАН', bold=True, size=14)
    add_paragraph(doc, '')

    semester = data.get('semester', '__')
    course = data.get('course', '__')
    year = data.get('year', '20__-20__')
    group = data.get('group', '_________')

    add_paragraph(doc, f'на ____ семестр(ы) {year} учебного года ____ курс', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, f'учебной группы (учебных групп) {group}', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, f'Профессиональный модуль: ПМ.{data["pm_code"]} {data["pm_name"]}', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, f'Междисциплинарные курсы: МДК {data["mdk_code"]} {data["mdk_name"]}', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, f'по профессии: {data["profession"]}', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')

    total = data['total_hours']
    practice_total = data['practice_hours']
    with_teacher = data.get('with_teacher_hours', total)
    theory = data.get('theory_hours', 0)
    practical = data.get('practical_hours', 0)
    lab = data.get('lab_hours', 0)
    coursework = data.get('coursework_hours', 0)
    independent = data.get('independent_hours', 0)
    exam = data.get('exam_form', '')

    add_paragraph(doc, f'Объем образовательной программы: ______{total}______________ (часов);', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'в том числе в форме практической подготовки ___{practice_total}______________(часа);', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, f'Учебная нагрузка во взаимодействии с преподавателем________{with_teacher}_______________ (часов):', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'из нее:', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'теоретическое обучение __{theory}___ (часов);                    практические занятия __{practical} __ (часов);', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'лабораторные занятия ___{lab if lab else "_____"}____ (часов);                   курсовая работа/проект ___{coursework if coursework else "_____"}____ (часов);', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'самостоятельная работа ___{independent if independent else "_____"}____ (часов);', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'посредняя аттестация в форме __ {exam}____', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '                                                                               (указать форму)', size=9, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')

    chair = data.get('commission_chair', 'Ткаченко А.Н.')
    add_paragraph(doc, f'Составлен в соответствии с рабочей программой ПМ {data["pm_code"]}, утверждённой __________________', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '              (дата утверждения)', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, 'Рассмотрен на заседании цикловой комиссии _____________________________ дисциплин', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, 'Протокол  №_______от________________20_____ года', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'Председатель цикловой комиссии ___________________________/{chair}/', size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')
    add_paragraph(doc, '')
    add_paragraph(doc, 'г. Сальск', size=11, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph(doc, f'{year} уч. год', size=11, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # ========== ТАБЛИЦА 1 ==========
    doc.add_page_break()
    add_paragraph(doc, 'Распределение часов по профессиональному модулю', bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, 'Таблица 1', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')

    t1_headers = ['Междисциплинарный курс (индекс МДК)', 'Курс', 'Семестр',
                   'Объём образовательной программы', 'Всего, часов',
                   'Теоретические занятия', 'Лабораторные работы, часов',
                   'Практические занятия, часов', 'Курсовые работы (проекты), часов',
                   'Самостоятельная работа обучающегося, часов',
                   'Учебная практика, часов', 'Производственная практика, часов']

    t1 = doc.add_table(rows=5, cols=len(t1_headers))
    t1.style = 'Table Grid'

    # Заголовок
    for j, h in enumerate(t1_headers):
        set_cell_text(t1.cell(0, j), h, bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(t1.cell(0, j), 'D9E2F3')

    # Данные МДК
    row = t1.rows[1]
    set_cell_text(row.cells[0], f'МДК {data["mdk_code"]}', size=8)
    set_cell_text(row.cells[3], str(total), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[4], str(with_teacher), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[5], str(theory), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[6], str(lab) if lab else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[7], str(practical) if practical else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[8], str(coursework) if coursework else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[9], str(independent) if independent else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[10], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[11], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Практика
    row_p = t1.rows[3]
    set_cell_text(row_p.cells[0], 'Практика', size=8)
    set_cell_text(row_p.cells[10], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row_p.cells[11], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Всего
    row_t = t1.rows[4]
    set_cell_text(row_t.cells[0], 'Всего', size=8, bold=True)
    set_cell_text(row_t.cells[3], str(total), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[4], str(with_teacher), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[5], str(theory), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[6], str(lab) if lab else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[7], str(practical) if practical else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[8], str(coursework) if coursework else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[9], str(independent) if independent else '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[10], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    set_cell_text(row_t.cells[11], '-', size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_paragraph(doc, '')
    add_paragraph(doc, f'Форма промежуточной аттестации обучающихся за семестр по междисциплинарному курсу', size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, f'(МДК {data["mdk_code"]} {data["mdk_name"]}) – ({exam}).', size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # ========== ТАБЛИЦА 2 ==========
    add_paragraph(doc, '')
    add_paragraph(doc, 'Содержание обучения по профессиональному модулю', bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '                                                                                                                                        Таблица 2', size=10, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    add_paragraph(doc, '')

    t2_headers = ['№ п/п', 'Наименование разделов, тем, тематики самостоятельной работы',
                   'Кол-во часов', 'В форме практической подготовки',
                   'Вид учебного занятия', 'Оснащение занятий',
                   'Задание для студентов', 'Примечание']
    t2_sub = ['1', '2', '3', '5', '', '6', '7', '8']

    # Считаем общее количество строк
    num_lessons = sum(len(topic['lessons']) for topic in data['topics'])
    num_topic_rows = len(data['topics'])
    total_rows = 3 + 1 + num_topic_rows + num_lessons + 1  # sub-headers + MDK header + topics + lessons + Итого

    t2 = doc.add_table(rows=total_rows, cols=8)
    t2.style = 'Table Grid'

    # Подзаголовки (номера граф)
    for j, h in enumerate(t2_sub):
        set_cell_text(t2.cell(0, j), h, bold=True, size=7, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(t2.cell(0, j), 'D9E2F3')

    # Заголовки колонок
    for j, h in enumerate(t2_headers):
        set_cell_text(t2.cell(1, j), h, bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(t2.cell(1, j), 'D9E2F3')

    # Заголовочная строка МДК
    row_idx = 2
    set_cell_text(t2.cell(row_idx, 1), f'МДК {data["mdk_code"]}   {data["mdk_name"]}', bold=True, size=8)
    set_cell_text(t2.cell(row_idx, 2), str(total), bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(t2.cell(row_idx, 3), str(practice_total), bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_shading(t2.cell(row_idx, 0), 'F2F2F2')
    for j in range(8):
        set_cell_shading(t2.cell(row_idx, j), 'F2F2F2')

    row_idx = 3
    lesson_num = 1

    for topic in data['topics']:
        # Заголовочная строка темы
        set_cell_text(t2.cell(row_idx, 1), topic['name'], bold=True, size=8)
        set_cell_text(t2.cell(row_idx, 2), str(topic['total_hours']), bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(t2.cell(row_idx, 3), str(topic['practice_hours']) if topic['practice_hours'] else '', bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        for j in range(8):
            set_cell_shading(t2.cell(row_idx, j), 'F2F2F2')
        row_idx += 1

        # Занятия
        for lesson in topic['lessons']:
            set_cell_text(t2.cell(row_idx, 0), str(lesson_num), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text(t2.cell(row_idx, 1), lesson['name'], size=8)
            set_cell_text(t2.cell(row_idx, 2), str(lesson['hours']), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)

            if lesson.get('practice_prep'):
                set_cell_text(t2.cell(row_idx, 3), str(lesson['practice_prep']), size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)

            # Вид занятия
            if lesson.get('is_lab'):
                vid = 'Лаб. работа'
            elif lesson.get('is_practice'):
                vid = 'Практ. занятие'
            else:
                vid = 'Урок'
            set_cell_text(t2.cell(row_idx, 4), vid, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)

            set_cell_text(t2.cell(row_idx, 5), lesson.get('equipment', ''), size=8)
            set_cell_text(t2.cell(row_idx, 6), lesson.get('task', ''), size=8)
            set_cell_text(t2.cell(row_idx, 7), lesson.get('note', ''), size=8)

            lesson_num += 1
            row_idx += 1

    # Итого
    set_cell_text(t2.cell(row_idx, 1), 'Итого:', bold=True, size=8)
    set_cell_text(t2.cell(row_idx, 2), str(total), bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(t2.cell(row_idx, 3), str(practice_total), bold=True, size=8, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    for j in range(8):
        set_cell_shading(t2.cell(row_idx, j), 'D9E2F3')

    # ========== ТАБЛИЦА 2а ==========
    doc.add_page_break()
    add_paragraph(doc, 'Материально-техническое обеспечение занятий', bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, 'Таблица 2а', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    add_paragraph(doc, '')

    equipment = data.get('equipment', [])
    t3 = doc.add_table(rows=max(len(equipment), 1) + 1, cols=2)
    t3.style = 'Table Grid'
    set_cell_text(t3.cell(0, 0), '№ п/п', bold=True, size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(t3.cell(0, 1), 'Материально-техническое обеспечение занятий', bold=True, size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    for j in range(2):
        set_cell_shading(t3.cell(0, j), 'D9E2F3')
    for i, eq in enumerate(equipment):
        set_cell_text(t3.cell(i + 1, 0), str(i + 1), size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(t3.cell(i + 1, 1), eq, size=9)

    # ========== ТАБЛИЦЫ 2б и 2в ==========
    add_paragraph(doc, '')
    add_paragraph(doc, 'Информационное обеспечение обучения', bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    sources_basic = data.get('sources_basic', [])
    if sources_basic:
        add_paragraph(doc, 'Основные источники (ОИ):', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        add_paragraph(doc, 'Таблица 2б', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        add_paragraph(doc, '')

        t4 = doc.add_table(rows=len(sources_basic) + 1, cols=4)
        t4.style = 'Table Grid'
        for j, h in enumerate(['№ п/п', 'Наименование', 'Автор', 'Издательство, год издания']):
            set_cell_text(t4.cell(0, j), h, bold=True, size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(t4.cell(0, j), 'D9E2F3')
        for i, src in enumerate(sources_basic):
            set_cell_text(t4.cell(i + 1, 0), src.get('code', f'ОИ {i+1}'), size=9)
            set_cell_text(t4.cell(i + 1, 1), src.get('name', ''), size=9)
            set_cell_text(t4.cell(i + 1, 2), src.get('author', ''), size=9)
            set_cell_text(t4.cell(i + 1, 3), src.get('publisher', ''), size=9)

    sources_add = data.get('sources_additional', [])
    if sources_add:
        add_paragraph(doc, '')
        add_paragraph(doc, 'Дополнительные источники (ДИ):', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        add_paragraph(doc, 'Таблица 2в', bold=True, size=10, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        add_paragraph(doc, '')

        t5 = doc.add_table(rows=len(sources_add) + 1, cols=4)
        t5.style = 'Table Grid'
        for j, h in enumerate(['№ п/п', 'Наименование', 'Автор', 'Издательство, год издания']):
            set_cell_text(t5.cell(0, j), h, bold=True, size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(t5.cell(0, j), 'D9E2F3')
        for i, src in enumerate(sources_add):
            set_cell_text(t5.cell(i + 1, 0), src.get('code', f'ДИ {i+1}'), size=9)
            set_cell_text(t5.cell(i + 1, 1), src.get('name', ''), size=9)
            set_cell_text(t5.cell(i + 1, 2), src.get('author', ''), size=9)
            set_cell_text(t5.cell(i + 1, 3), src.get('publisher', ''), size=9)

    doc.save(output_path)
    print(f'КТП сохранён: {output_path}')


def main():
    parser = argparse.ArgumentParser(description='Генерация КТП в формате DOCX')
    parser.add_argument('--data', required=True, help='Путь к JSON-файлу с данными КТП')
    parser.add_argument('--output', required=True, help='Путь к выходному файлу .docx')
    args = parser.parse_args()

    with open(args.data, 'r', encoding='utf-8') as f:
        data = json.load(f)

    generate_ktp(data, args.output)


if __name__ == '__main__':
    main()