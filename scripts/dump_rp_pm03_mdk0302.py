# -*- coding: utf-8 -*-
"""Дамп тематического плана МДК 03.02 из РП ПМ.03 (15.01.37) для построения КТП."""
import re, os, zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'RP', '15.01.37_РП_ПМ.03_2025.docx')

def docx_tables(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    tables = []
    for tbl in root.iter(f'{{{W}}}tbl'):
        rows = []
        for tr in tbl.findall('w:tr', NS):
            cells = []
            for tc in tr.findall('w:tc', NS):
                txt = ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t'))
                tcpr = tc.find('w:tcPr', NS)
                span = 1
                if tcpr is not None:
                    gs = tcpr.find('w:gridSpan', NS)
                    if gs is not None:
                        span = int(gs.get(f'{{{W}}}val'))
                vm = tcpr is not None and tcpr.find('w:vMerge', NS) is not None
                cont = bool(vm) and (tcpr.find('w:vMerge', NS).get(f'{{{W}}}val') != 'restart')
                cells.append({'text': txt.strip(), 'span': span, 'vmc': cont})
            rows.append(cells)
        tables.append(rows)
    return tables

tables = docx_tables(os.path.abspath(PATH))
# Таблица #7 — тематический план (140 строк)
for ti, rows in enumerate(tables):
    joined = ' '.join(c['text'] for r in rows[:3] for c in r)
    if '3.2' in joined and ('тематическ' in joined.lower() or 'Содержание' in joined):
        print(f'### Таблица #{ti+1} — заголовок: {joined[:150]}')
for ti, rows in enumerate(tables):
    if len(rows) == 140:
        print(f'### Таблица #{ti+1} (140 строк) — блок МДК 03.02 (стр.38..76):')
        for ri in range(36, min(78, len(rows))):
            cells = []
            for c in rows[ri]:
                if c['vmc']:
                    cells.append('<v>')
                else:
                    t = c['text'].replace('\n', ' ')[:110]
                    cells.append(t if t else '·' if c['span'] == 1 else f'[x{c["span"]}]')
            print(f'{ri:3d} | ' + ' || '.join(cells))
        break
