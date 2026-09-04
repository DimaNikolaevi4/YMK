# -*- coding: utf-8 -*-
"""Правка исправленного КТП МДК 01.02 (от владельца):
1) Т1: теория с.3 36 → 38 (КДЗ 2 ч перенесён в с.4 по образцу эталона 05.02;
   иначе «Всего» теория 56 и объём с.3 = 38 не бьются);
2) Т2: trHeight всех строк по эталону (шапка: 20/230/1695, тело: 20).
Результат → KTP/КТП МДК 01.02 М-21.docx + копия в download/."""
import zipfile, shutil, os
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
SRC = '/home/z/my-project/upload/КТП МДК 01.02 М 2 курс.docx'
DST = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'

ET.register_namespace('w', W)
for pfx, uri in [
    ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
    ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('v', 'urn:schemas-microsoft-com:vml'),
    ('o', 'urn:schemas-microsoft-com:office:office'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
]:
    ET.register_namespace(pfx, uri)

with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    data = {n: z.read(n) for n in names}

root = ET.fromstring(data['word/document.xml'].decode('utf-8'))
body = root.find('w:body', NS)
tbls = body.findall('w:tbl', NS)


def cell_text(tc):
    return ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()


# ---------- 1. Т1: строка МДК 01.02 с.3 — теория 36 → 38 ----------
t1 = None
for tbl in tbls:
    head = ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:150]
    if 'Междисциплинарный курс' in head:
        t1 = tbl
        break
assert t1 is not None, 'Т1 не найдена'
fixed_t1 = 0
for tr in t1.findall('w:tr', NS):
    tcs = tr.findall('w:tc', NS)
    texts = [cell_text(tc) for tc in tcs]
    if texts and texts[0] == 'МДК 01.02' and len(texts) > 5 and texts[2] == '3' and texts[5] == '36':
        # ячейка с "36" (теория с.3): заменить текст первого w:t в первом w:p
        target = tcs[5]
        ts = target.findall('.//w:t', NS)
        for t in ts:
            if t.text and t.text.strip() == '36':
                t.text = '38'
                fixed_t1 += 1
                break
        # xml:space preserve не нужен — чистое число
assert fixed_t1 == 1, f'Т1: ожидалась 1 правка, сделано {fixed_t1}'
print('Т1: теория с.3 36 → 38: OK')

# ---------- 2. Т2: trHeight по эталону ----------
t2 = None
for tbl in tbls:
    head = ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:150]
    if '№ занятия' in head:
        t2 = tbl
        break
assert t2 is not None, 'Т2 не найдена'
HEIGHTS = {0: '20', 1: '230', 2: '1695'}  # остальным строкам — 20
trs = t2.findall('w:tr', NS)
for i, tr in enumerate(trs):
    val = HEIGHTS.get(i, '20')
    trpr = tr.find('w:trPr', NS)
    if trpr is None:
        trpr = ET.Element(f'{{{W}}}trPr')
        tr.insert(0, trpr)
    # убрать существующие trHeight
    for th in trpr.findall('w:trHeight', NS):
        trpr.remove(th)
    th = ET.SubElement(trpr, f'{{{W}}}trHeight')
    th.set(f'{{{W}}}val', val)
print(f'Т2: trHeight проставлены на {len(trs)} строк (0=20, 1=230, 2=1695, тело=20)')

data['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.writestr(n, data[n])
print('Сохранено:', DST)

# копия в download
os.makedirs('/home/z/my-project/download', exist_ok=True)
shutil.copy(DST, '/home/z/my-project/download/КТП МДК 01.02 М-21.docx')
print('Копия: /home/z/my-project/download/КТП МДК 01.02 М-21.docx')
