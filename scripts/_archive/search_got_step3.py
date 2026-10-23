# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 3: страницы в столбце 9 (ячейка 8) Т2 всех четырёх файлов
+ полный diff Т1/Т2 ГОТ vs НАША."""
import sys, re
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

FILES = {
    'ГОТ 01.02': '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx',
    'НАША 01.02': '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx',
    'ГОТ 05.02': '/home/z/my-project/YMK/KTP/гот/КТП МДК 05.02 М 2 курс.docx',
    'ЭТАЛОН 05.02': '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx',
    'UPLOAD (старый) 01.02': '/home/z/my-project/upload/КТП МДК 01.02 М 2 курс.docx',
}

PAGE_RE = re.compile(r'стр\.?\s*\d|ОИ\s*\d|метод', re.IGNORECASE)


def find_t2(tables):
    for rows in tables:
        head = ' '.join(c['text'] for r in rows[:4] for c in r)
        if '№ занятия' in head or (len(rows) > 10 and rows[0][0]['text'] == '№ занятия'):
            return rows
    return None


def pages_report(path, label):
    _, tables = doc_text_and_tables(path)
    t2 = find_t2(tables)
    with_p = []
    for i, row in enumerate(t2[3:], 3):
        if len(row) > 8:
            c8 = row[8]['text'].strip()
            if c8 and PAGE_RE.search(c8):
                with_p.append((i, row[0]['text'], c8[:45]))
    total = sum(1 for r in t2[3:] if r[0]['text'].strip().isdigit())
    print(f'--- {label}: занятий {total}; строк где кол.9 непусто/со «стр|ОИ|метод»: {len(with_p)}')
    samples = with_p[:6]
    for i, num, c in samples:
        print(f'    №{num}: «{c}»')
    if not with_p:
        # показать, что вообще в кол.9
        vals = [r[8]['text'][:40] for r in t2[3:] if len(r) > 8 and r[8]['text'].strip()][:6]
        print('    примеры содержимого кол.9:', vals)


for lbl, p in FILES.items():
    try:
        pages_report(p, lbl)
    except Exception as e:
        print(f'--- {lbl}: ОШИБКА {e}')
    print()

# полный diff Т1 и Т2: ГОТ 01.02 vs НАША 01.02 и ГОТ 05.02 vs ЭТАЛОН
def diff_tables(p1, l1, p2, l2, tidx, name):
    _, t_a = doc_text_and_tables(p1)
    _, t_b = doc_text_and_tables(p2)
    a, b = t_a[tidx], t_b[tidx]
    print(f'=== {name}: {l1} ({len(a)} строк) vs {l2} ({len(b)} строк) ===')
    ndiff = 0
    for i in range(max(len(a), len(b))):
        ra = [c['text'] for c in a[i]] if i < len(a) else ['<нет>']
        rb = [c['text'] for c in b[i]] if i < len(b) else ['<нет>']
        if ra != rb:
            ndiff += 1
            if ndiff <= 12:
                print(f'  строка {i}:')
                print(f'    {l1}: {[x[:42] for x in ra]}')
                print(f'    {l2}: {[x[:42] for x in rb]}')
    print(f'  всего различающихся строк: {ndiff}\n')

diff_tables(FILES['ГОТ 01.02'], 'ГОТ', FILES['НАША 01.02'], 'НАША', 1, 'Т1 01.02')
diff_tables(FILES['ГОТ 01.02'], 'ГОТ', FILES['НАША 01.02'], 'НАША', 2, 'Т2 01.02')
diff_tables(FILES['ГОТ 01.02'], 'ГОТ', FILES['UPLOAD (старый) 01.02'], 'UPLOAD-старый', 2, 'Т2 01.02 ГОТ vs старый загруженный')
diff_tables(FILES['ГОТ 05.02'], 'ГОТ', FILES['ЭТАЛОН 05.02'], 'ЭТАЛОН', 2, 'Т2 05.02')
