# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 4: полная Т1 ГОТ/НАША 01.02 + ЭИ старого загруза."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

GOT = '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx'
OUR = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'
OLD = '/home/z/my-project/upload/КТП МДК 01.02 М 2 курс.docx'

for lbl, path in (('ГОТ 01.02', GOT), ('НАША 01.02', OUR)):
    _, tables = doc_text_and_tables(path)
    print(f'=== Т1 {lbl} (12 строк, полный дамп) ===')
    for i, row in enumerate(tables[1]):
        print(f'  {i}: {[c["text"][:28] for c in row]}')
    print()

print('=== ЭИ (таблица 6) старого загруза UPLOAD ===')
_, t_old = doc_text_and_tables(OLD)
for row in t_old[6]:
    print('  ', ' | '.join(c['text'][:45] for c in row))
print()
print('=== ЭИ (таблица 6) ГОТ (новый загруз) ===')
_, t_got = doc_text_and_tables(GOT)
for row in t_got[6]:
    print('  ', ' | '.join(c['text'][:60] for c in row))
