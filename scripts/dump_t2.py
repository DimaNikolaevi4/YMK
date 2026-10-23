# -*- coding: utf-8 -*-
"""Полный дамп Т2 (занятия) обоих КТП: №, раздел/тема, текст, кол.4, кол.8, кол.9 + суммы."""
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
            out.append([celltext(r[k]).replace('\n', ' ') for k in keys])
        else:
            out.append([celltext(v).replace('\n', ' ') for v in r])
    return out

for label in ['02.01', '02.02']:
    path = f'{YMK}/KTP/КТП МДК {label} С-21.docx'
    texts, tables = doc_text_and_tables(path)
    for t in tables:
        rows = rows_of(t)
        if not rows or '№ занятия' not in (rows[0][0] if rows[0] else ''):
            continue
        print(f'\n########## Т2 КТП МДК {label} С-21: {len(rows)} строк ##########')
        sums = 0
        for i, row in enumerate(rows):
            joined = ' '.join(row)
            if i < 3 or not any(row):
                continue
            # Т2 колонки: 0=№, 1=Раздел, 2=Тема, 3=час(кол.4?), ..., кол.8/кол.9 около 8/9
            n = row[0]
            hrs = row[3] if len(row) > 3 else ''
            c8 = row[8] if len(row) > 8 else ''
            c9 = row[9] if len(row) > 9 else ''
            topic = (row[2] or row[1])[:75]
            print(f'{n:>4} | ч:{hrs:>3} | {topic:75s} | к8:"{c8}" | к9:"{c9}"')
            if hrs.isdigit():
                sums += int(hrs)
        print(f'>>> СУММА кол.4 (числовые строки): {sums}')
