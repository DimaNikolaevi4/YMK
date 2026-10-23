# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 2: дамп строк Т2 эталона и ГОТ-файлов + diff литературы."""
import sys, re, hashlib
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

FILES = {
    'ГОТ 01.02': '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx',
    'НАША 01.02': '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx',
    'ГОТ 05.02': '/home/z/my-project/YMK/KTP/гот/КТП МДК 05.02 М 2 курс.docx',
    'ЭТАЛОН 05.02': '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx',
}

def md5(p):
    return hashlib.md5(open(p, 'rb').read()).hexdigest()[:10]

# 1) контрольные суммы: может, ГОТ = ранее загруженные файлы из /upload?
import os
up_dir = '/home/z/my-project/upload'
if os.path.isdir(up_dir):
    for f in sorted(os.listdir(up_dir)):
        p = os.path.join(up_dir, f)
        print(f'UPLOAD: {f}  md5={md5(p)}  size={os.path.getsize(p)}')
for lbl, p in FILES.items():
    print(f'{lbl}: md5={md5(p)} size={os.path.getsize(p)}')

# 2) полная строка Т2 эталона (теоретическое занятие) — где столбец «Задания»?
_, t_et = doc_text_and_tables(FILES['ЭТАЛОН 05.02'])
t2 = t_et[2]
print('\n=== ЭТАЛОН 05.02, Т2 строка 4 и 5 (полностью) ===')
for i in (3, 4):
    for j, c in enumerate(t2[i]):
        print(f'  [{j}] span={c["span"]} «{c["text"][:60]}»')

# 3) где в Т2 страниц-подобные строки у эталона?
PAGE_RE = re.compile(r'стр\.?\s*\d', re.IGNORECASE)
cnt = 0
print('\n=== ЭТАЛОН: строки Т2, где ЛЮБАЯ ячейка содержит «стр. N» ===')
for i, row in enumerate(t2):
    for j, c in enumerate(row):
        if PAGE_RE.search(c['text']):
            cnt += 1
            if cnt <= 6:
                print(f'  строка {i}, ячейка [{j}]: «{c["text"][:60]}»')
print('всего:', cnt)

# 4) diff литературы 01.02: таблицы 4/5/6
print('\n=== ЛИТЕРАТУРА 01.02: ГОТ vs НАША ===')
_, t_got = doc_text_and_tables(FILES['ГОТ 01.02'])
_, t_our = doc_text_and_tables(FILES['НАША 01.02'])
for ti in (4, 5, 6):
    print(f'--- Таблица #{ti} ---')
    g = [' | '.join(c['text'] for c in r) for r in t_got[ti]]
    o = [' | '.join(c['text'] for c in r) for r in t_our[ti]]
    for line in g:
        mark = 'ГОТ ' if line in o else '>>> только в ГОТ'
        print(f'  {mark}: {line[:110]}')
    for line in o:
        if line not in g:
            print(f'  >>> только в НАША: {line[:110]}')
