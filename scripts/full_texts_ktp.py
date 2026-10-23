# -*- coding: utf-8 -*-
"""Полные тексты: Т2 занятия КТП 02.01/02.02 С-21 vs пункты РП 3.2."""
import sys, zipfile
from xml.etree import ElementTree as ET
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables, flat

BASE = '/home/z/my-project/YMK'

def ktp_rows(f):
    texts, tables = doc_text_and_tables(f'{BASE}/{f}')
    for ti, rows in enumerate(tables):
        head = ' '.join(c['text'] for rr in rows[:3] for c in rr)
        if '№ занятия' in head:
            ncols = max(sum(c['span'] for c in rr) for rr in rows)
            out = []
            for rr in rows:
                f2 = (flat(rr) + [''] * ncols)[:ncols]
                out.append(f2)
            return out
    return None

for f in ['KTP/КТП МДК 02.01 С-21.docx', 'KTP/КТП МДК 02.02 С-21.docx']:
    print('#' * 100)
    print(f)
    rows = ktp_rows(f)
    for ri, f2 in enumerate(rows):
        if ri < 4:
            continue
        no, tema, k3, k4 = f2[0], f2[1], f2[2], f2[3]
        print(f'--- r{ri}: №[{no}] часы[{k3}|{k4}] {tema}')
