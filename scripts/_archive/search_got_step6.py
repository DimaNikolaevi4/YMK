# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 6: trHeight шапки Т2 в ГОТ vs НАША vs эталон + план синхронизации."""
import sys, zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def t2_trheights(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    body = root.find('w:body', NS)
    tables = body.findall('w:tbl', NS)
    t2 = tables[2]
    out = []
    for tr in t2.findall('w:tr', NS)[:8]:
        trpr = tr.find('w:trPr', NS)
        h = None
        if trpr is not None:
            e = trpr.find('w:trHeight', NS)
            if e is not None:
                h = e.get(f'{{{W}}}val')
        texts = ''.join(n.text or '' for n in tr.iter(f'{{{W}}}t'))[:30]
        out.append((h, texts))
    return out


FILES = {
    'ГОТ 01.02 (новый)': '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx',
    'НАША 01.02': '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx',
    'ЭТАЛОН 05.02': '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx',
    'ГОТ 05.02': '/home/z/my-project/YMK/KTP/гот/КТП МДК 05.02 М 2 курс.docx',
}
for lbl, p in FILES.items():
    print(f'=== {lbl} ===')
    for h, t in t2_trheights(p):
        print(f'  trHeight={h}: {t}')
