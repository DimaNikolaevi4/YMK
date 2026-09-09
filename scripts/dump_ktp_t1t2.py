# -*- coding: utf-8 -*-
"""Полный дамп Т1 + Т2 КТП 02.01/02.02 С-21."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables, flat

BASE = '/home/z/my-project/YMK'
for f in ['KTP/КТП МДК 02.01 С-21.docx', 'KTP/КТП МДК 02.02 С-21.docx']:
    p = f'{BASE}/{f}'
    print('=' * 100)
    print(f)
    texts, tables = doc_text_and_tables(p)
    for ti, rows in enumerate(tables):
        head = ' '.join(c['text'] for rr in rows[:3] for c in rr)
        ncols = max(sum(c['span'] for c in rr) for rr in rows)
        if 'Междисциплинарный курс' in head and ti < 3:
            print(f'--- Таблица #{ti} (Т1), {len(rows)} строк ---')
            for ri, rr in enumerate(rows):
                f2 = (flat(rr) + [''] * ncols)[:ncols]
                print(f'  r{ri}: {" | ".join(x[:26] for x in f2)}')
        elif '№ занятия' in head:
            print(f'--- Таблица #{ti} (Т2), {len(rows)} строк ---')
            for ri, rr in enumerate(rows):
                f2 = (flat(rr) + [''] * ncols)[:ncols]
                # печать: №, тема(обрез), кол3, кол4, вид, ОК, ПК, кол8, кол9, кол10
                no, tema, k3, k4, vid = f2[0], f2[1], f2[2], f2[3], f2[4]
                print(f'  r{ri}: №={no} | {tema[:55]} | {k3} | {k4} | {vid[:14]} | {f2[7][:12]} | {f2[8][:12]} | {f2[9][:16]}')
