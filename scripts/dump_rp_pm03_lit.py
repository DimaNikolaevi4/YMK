# -*- coding: utf-8 -*-
"""Извлечение литературы/оборудования/ОК-ПК из РП ПМ.03 (15.01.37) + проверка 4-часовых практ. в КТП 01.02."""
import re, os, zipfile, sys
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')

def docx_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    paras = []
    for p in root.iter(f'{{{W}}}p'):
        t = ''.join(n.text or '' for n in p.iter(f'{{{W}}}t')).strip()
        if t:
            paras.append(t)
    return paras

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
                txt = ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()
                cells.append(txt)
            rows.append(cells)
        tables.append(rows)
    return tables

rp = os.path.join(BASE, 'RP', '15.01.37_РП_ПМ.03_2025.docx')
paras = docx_text(rp)
print('=== ПОИСК ЛИТЕРАТУРЫ/ОБОРУДОВАНИЯ (текст РП ПМ.03) ===')
capture = None
for i, t in enumerate(paras):
    tl = t.lower()
    if 'основная литература' in tl or 'основные источники' in tl:
        capture = 'ОИ'
    elif 'дополнительная литература' in tl or 'дополнительные источники' in tl:
        capture = 'ДИ'
    elif 'электронные изд' in tl or 'интернет-ресурс' in tl or 'электронные ресурс' in tl:
        capture = 'ЭИ'
    elif 'условия реализации' in tl and 'обеспечение' in tl:
        capture = 'ОБОРУД'
    elif re.match(r'\d+\.\d+', t) and capture and 'литератур' not in tl:
        pass
    if capture and (re.match(r'(ОИ|ДИ|ЭИ)\s*\d', t) or 'кабинет' in tl or 'лаборатория' in tl or 'мастерская' in tl or capture == 'ОБОРУД'):
        if len(t) > 3:
            print(f'[{capture}] {t[:160]}')
    if capture == 'ОБОРУД' and 'технического обслуживания' in tl and 'средств' in tl:
        pass

print()
print('=== Таблицы РП с литературой (содержат «Академия»/«urait»/http) ===')
tables = docx_tables(rp)
for ti, rows in enumerate(tables):
    j = ' '.join(' '.join(r) for r in rows)
    if re.search(r'Академия|Урайт|urait|http|издател', j, re.I):
        print(f'--- Таблица #{ti+1} ({len(rows)} строк)')
        for r in rows:
            line = ' || '.join(x.replace('\n', ' ')[:90] for x in r if x)
            if line:
                print('   ', line[:250])

print()
print('=== КТП МДК 01.02 М-21: как оформлены 4-часовые практические ===')
sys.path.insert(0, os.path.join(BASE, 'scripts'))
from extract_ktp_hours import doc_text_and_tables
t, ktables = doc_text_and_tables(os.path.join(BASE, 'KTP', 'КТП МДК 01.02 М-21.docx'))
for rows in ktables:
    j = ' '.join(c['text'] for r in rows[:4] for c in r)
    if re.search(r'Тематический план|Содержание учебного материала', j, re.I):
        for r in rows:
            f = ([c['text'] for c in r] + [''] * 11)[:11]
            name = f[1]
            if 'Практ' in name or 'Лаб' in name or 'практ' in name.lower():
                print(' | '.join(x.replace(chr(10), ' ')[:55] for x in f[:6]))
        break
