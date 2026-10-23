# -*- coding: utf-8 -*-
"""Сравнение текущих КТП 02.01/02.02 С-21/С-22 с новым РП ПМ.02 (15.01.37)."""
import re, sys, zipfile
from xml.etree import ElementTree as ET
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import analyze, doc_text_and_tables, flat

BASE = '/home/z/my-project/YMK'
FILES = [
    'KTP/КТП МДК 02.01 С-21.docx',
    'KTP/КТП МДК 02.01 С-22.docx',
    'KTP/КТП МДК 02.02 С-21.docx',
    'KTP/КТП МДК 02.02 С-22.docx',
]

for f in FILES:
    p = f'{BASE}/{f}'
    print('=' * 90)
    print(f)
    r = analyze(p)
    for k, v in r.items():
        if k not in ('file', 'dir'):
            print(f'   {k} = {v}')
    # Т2 построчно: все строки с номером занятия
    texts, tables = doc_text_and_tables(p)
    for ti, rows in enumerate(tables):
        head = ' '.join(c['text'] for rr in rows[:3] for c in rr)
        if '№ занятия' in head or ('Наименование' in head and 'Раздел' in head):
            print(f'   --- Таблица #{ti} (Т2): {len(rows)} строк ---')
            ncols = max(sum(c["span"] for c in rr) for rr in rows)
            for ri, rr in enumerate(rows[:8]):
                f2 = (flat(rr) + [''] * ncols)[:ncols]
                print(f'     r{ri}: {" | ".join(x[:40] for x in f2)}')
