# -*- coding: utf-8 -*-
"""Дамп всех таблиц нового РП 15.01.37_РП_ПМ.02_2025.docx + текст между ними."""
import re, sys, zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
PATH = '/home/z/my-project/YMK/RP/15.01.37_РП_ПМ.02_2025.docx'

with zipfile.ZipFile(PATH) as z:
    xml = z.read('word/document.xml').decode('utf-8')
root = ET.fromstring(xml)
body = root.find('w:body', NS)

tbl_i = 0
for el in body:
    tag = el.tag.split('}')[1]
    if tag == 'p':
        t = ''.join(n.text or '' for n in el.iter(f'{{{W}}}t')).strip()
        if t:
            print(f'P: {t[:180]}')
    elif tag == 'tbl':
        rows = []
        for tr in el.findall('w:tr', NS):
            cells = []
            for tc in tr.findall('w:tc', NS):
                txt = ''.join(n.text or '' for n in tc.iter(f'{{{W}}}t')).strip()
                span = 1
                tcpr = tc.find('w:tcPr', NS)
                if tcpr is not None:
                    gs = tcpr.find('w:gridSpan', NS)
                    if gs is not None:
                        span = int(gs.get(f'{{{W}}}val'))
                cells.append((txt, span))
            rows.append(cells)
        ncols = max(sum(s for _, s in r) for r in rows) if rows else 0
        print(f'\n===== Таблица #{tbl_i} ({len(rows)} строк x {ncols} колонок) =====')
        for ri, r in enumerate(rows):
            # разворачиваем gridSpan в колонки
            f = []
            for txt, span in r:
                f.append(txt)
                if span > 1:
                    f.extend([''] * (span - 1))
            line = ' | '.join(x for x in f)
            line = re.sub(r'\s+', ' ', line)
            print(f'  r{ri}: {line[:300]}')
        tbl_i += 1
