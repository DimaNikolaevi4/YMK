# -*- coding: utf-8 -*-
"""Краткий дамп всех таблиц двух КТП: первые ячейки строк + Т1 целиком."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

YMK = '/home/z/my-project/YMK'

def celltext(v):
    if isinstance(v, dict):
        return str(v.get('text', '')).strip()
    return str(v).strip()

def rows_of(table):
    out = []
    for r in table:
        if isinstance(r, dict):
            keys = sorted(r.keys())
            out.append([celltext(r[k]).replace('\n', '⏎') for k in keys])
        else:
            out.append([celltext(v).replace('\n', '⏎') for v in r])
    return out

for label in ['02.01', '02.02']:
    path = f'{YMK}/KTP/КТП МДК {label} С-21.docx'
    texts, tables = doc_text_and_tables(path)
    print(f'\n########## КТП МДК {label} С-21: {len(tables)} таблиц ##########')
    for ti, t in enumerate(tables):
        rows = rows_of(t)
        first = rows[0][0][:60] if rows and rows[0] else '?'
        print(f'\n--- таблица #{ti}: {len(rows)} строк; r0c0="{first}" ---')
        if rows and rows[0] and ('Междисциплинарный' in rows[0][0] or 'Календарно' in first or 'тематический' in first.lower()):
            for i, row in enumerate(rows):
                nz = [f'{j}:"{v}"' for j, v in enumerate(row) if v]
                print(f'r{i}: ' + (' | '.join(nz) if nz else '(все пусты)'))
