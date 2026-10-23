# -*- coding: utf-8 -*-
"""Команда «ищи»: разбор загруженных готовых КТП (KTP/гот/) —
поиск страниц в столбце «Задания для студентов», литературы, отличий от наших версий."""
import re, sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

GOT_0102 = '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx'
GOT_0502 = '/home/z/my-project/YMK/KTP/гот/КТП МДК 05.02 М 2 курс.docx'
OURS_0102 = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'
OURS_0502 = '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx'

PAGE_RE = re.compile(r'(стр\.?\s*[\d\-–\s,]+|ОИ\s*\d[^;]{0,40})', re.IGNORECASE)


def find_table(tables, *patterns):
    for idx, rows in enumerate(tables):
        head = ' '.join(c['text'] for r in rows[:6] for c in r)
        if all(p.lower() in head.lower() for p in patterns):
            return idx, rows
    return None, None


def analyze(path, label):
    t, tables = doc_text_and_tables(path)
    print('=' * 100)
    print(f'### {label}: {path.split("/")[-1]}  (абзацев {len(t)}, таблиц {len(tables)})')
    # титул: год
    tb = ' | '.join(t[:40])
    m = re.search(r'20\d\d\s*[-–—/]\s*20\d\d', tb)
    print('Уч. год в титуле:', m.group(0) if m else '?')
    # шапки всех таблиц — чтобы понять структуру
    for idx, rows in enumerate(tables):
        first = ' '.join(c['text'] for c in rows[0])[:90]
        print(f'  Таблица #{idx}: {len(rows)} строк | 1-я: {first}')
    return t, tables


def t2_pages(path, label):
    """Ищем в Т2 (№ занятия) столбец «Задания для студентов» и строки со страницами."""
    t, tables = doc_text_and_tables(path)
    idx, t2 = find_table(tables, '№ занятия')
    if t2 is None:
        print(f'{label}: Т2 не найдена!')
        return
    # заголовок колонок
    hdr = [c['text'] for c in t2[2]] if len(t2) > 2 else []
    print(f'\n--- {label}: Т2 #{idx}, {len(t2)} строк; шапка строка 2: {hdr}')
    with_pages, total_body = [], 0
    for i, row in enumerate(t2[3:], start=3):
        cells = [c['text'] for c in row]
        # столбец 9 = «Задания для студентов» (последняя широкая)
        body = cells[-1] if cells else ''
        num = cells[0].strip() if cells else ''
        if num.isdigit():
            total_body += 1
        m = PAGE_RE.search(body)
        if m and (num.isdigit() or not num):
            with_pages.append((i, num, m.group(0)[:50], body[:70]))
    print(f'    занятий с нумерацией: {total_body}; строк со ссылками на страницы/ОИ: {len(with_pages)}')
    for i, num, pg, body in with_pages[:8]:
        print(f'    строка {i} №{num}: «{pg}»  | {body}')


def lit_tables(path, label):
    """Таблицы литературы (2б/2в/2г) — ищем строки с ОИ/ДИ/ЭИ."""
    t, tables = doc_text_and_tables(path)
    print(f'\n--- {label}: литература (абзацы и таблицы с «Основные источники») ---')
    for j, par in enumerate(t):
        if re.search(r'(Основные источники|Дополнительные|Электронные|Интернет-ресурсы)', par):
            print(f'    абзац[{j}]: {par[:100]}')
    for idx, rows in enumerate(tables):
        txt = ' '.join(c['text'] for r in rows[:3] for c in r)
        if 'ОИ' in txt or 'основные источники' in txt.lower():
            print(f'    Таблица #{idx} ({len(rows)} строк): {txt[:120]}')


if __name__ == '__main__':
    for path, label in [(GOT_0102, 'ГОТ 01.02'), (OURS_0102, 'НАША 01.02'),
                        (GOT_0502, 'ГОТ 05.02'), (OURS_0502, 'НАША 05.02 (эталон)')]:
        t, tables = analyze(path, label)
        t2_pages(path, label)
        lit_tables(path, label)
        print()
