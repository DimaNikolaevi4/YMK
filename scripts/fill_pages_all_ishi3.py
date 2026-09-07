# -*- coding: utf-8 -*-
"""Команда владельца «делай, редактируй все (полностью все из нагрузки) ктп… если нет
данных пиши приблизительно страницы» (ищи-3, 2026-09-07).

Заполняет кол. 9 Т2 во ВСЕХ 8 КТП нагрузки по плану scripts/pages_plan_ishi3.json:
- КТП 05.02 (эталон): 7 занятий (были вымышленные «Стр. …» из исходника);
- КТП 01.02 М-21: 12 занятий тем 2.3/2.4 («Конспект лекций» → страницы);
- КТП 02.01 С-21/С-22: 18 теоретических занятий;
- КТП 02.02 С-21/С-22: 11 теоретических занятий;
- КТП 03.02 С-21/С-22: 12 теоретических занятий.
Практические/лабораторные остаются «Метод. указания» (эталон); диф.зачёт — пусто.
Повторный запуск безопасен: assert на исходные значения; при уже заполненной ячейке
(совпадает с новым значением) — пропуск без ошибки.
"""
import zipfile, json, sys
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

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


def cell_text(tc):
    return ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()


def set_cell_text(tc, new_text):
    ts = tc.findall('.//w:t', NS)
    assert ts, 'нет w:t в ячейке'
    first = True
    for t in ts:
        if first:
            t.text = new_text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            first = False
        else:
            t.text = ''


def fill_col9(path, pages, label):
    """pages: {№ занятия: {'old':…, 'new':…}}."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    root = ET.fromstring(data['word/document.xml'].decode('utf-8'))
    body = root.find('w:body', NS)
    t2 = None
    for tbl in body.findall('w:tbl', NS):
        head = ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:200]
        if '№ занятия' in head:
            t2 = tbl
            break
    assert t2 is not None, f'{label}: Т2 не найдена'
    filled, skipped, mism = [], [], []
    for tr in t2.findall('w:tr', NS):
        tcs = tr.findall('w:tc', NS)
        if len(tcs) < 9:
            continue
        c0 = cell_text(tcs[0])
        if c0.isdigit() and int(c0) in pages:
            old = cell_text(tcs[8])
            if old.isdigit():
                # номерная строка шапки Т2 («1|2|…|11») — не занятие
                continue
            want_old = pages[int(c0)]['old']
            new = pages[int(c0)]['new']
            if old == new:
                skipped.append(int(c0))
            elif old == want_old:
                set_cell_text(tcs[8], new)
                filled.append(int(c0))
            else:
                mism.append(f"№{c0}: в файле «{old}», ожидалось «{want_old}»")
    assert not mism, f'{label}: расхождения:\n  ' + '\n  '.join(mism)
    assert sorted(filled + skipped) == sorted(pages), \
        f'{label}: затронуто {sorted(filled + skipped)}, ожидалось {sorted(pages)}'
    data['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    print(f'{label}: заполнено {len(filled)}, уже было {len(skipped)} (из {len(pages)})')
    return filled


plan = json.load(open('/home/z/my-project/YMK/scripts/pages_plan_ishi3.json', encoding='utf-8'))

TOTAL = {}
for key in ['ktp_05_02', 'ktp_01_02', 'ktp_02_01', 'ktp_02_02', 'ktp_03_02']:
    blk = plan[key]
    pages = {int(k): v for k, v in blk['lessons'].items()}
    files = [blk['file']] if 'file' in blk else blk['files']
    for f in files:
        label = f.split('/')[-1].replace('.docx', '')
        TOTAL.setdefault(key, []).append((label, fill_col9(f, pages, label)))

print('\nИТОГ:')
n = 0
for key, items in TOTAL.items():
    for label, filled in items:
        n += len(filled)
        print(f'  {label}: {len(filled)} ячеек')
print(f'ВСЕГО заполнено/изменено ячеек кол.9: {n}')
