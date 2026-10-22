# -*- coding: utf-8 -*-
"""Все таблицы РП 15.01.31 ПМ.02 кратко + поиск Т3.4 и литературы (ОИ/ДИ/ЭИ)."""
import re, sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

RP = '/home/z/my-project/YMK/RP/15.01.31_РП_ПМ.02_2024.docx'

def celltext(v):
    if isinstance(v, dict):
        return str(v.get('text', '')).strip()
    return str(v).strip()

def rows_of(table):
    out = []
    for r in table:
        if isinstance(r, dict):
            keys = sorted(r.keys())
            out.append([celltext(r[k]).replace('\n', ' ') for k in keys])
        else:
            out.append([celltext(v).replace('\n', ' ') for v in r])
    return out

texts, tables = doc_text_and_tables(RP)
print(f'Таблиц: {len(tables)}')
for ti, t in enumerate(tables):
    rows = rows_of(t)
    print(f'\n===== Таблица #{ti}: {len(rows)} строк =====')
    for i, row in enumerate(rows[:14]):
        nz = [f'{j}:"{v[:80]}"' for j, v in enumerate(row) if v]
        print(f'r{i}: ' + (' | '.join(nz) if nz else '(пусто)'))
    if len(rows) > 14:
        print(f'   ... ещё {len(rows)-14} строк')

print('\n===== Тексты: ОИ/ДИ/ЭИ/литература/аттестация =====')
for s in texts:
    t = s.strip()
    if re.search(r'(ОИ|ДИ|ЭИ)\s*[-–]?\s*\d|литерат|аттеста|Основные источники|Дополнительные|Электронные', t, re.I):
        print('  • ' + t[:190].replace('\n', ' '))
