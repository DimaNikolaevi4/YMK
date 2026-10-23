# -*- coding: utf-8 -*-
"""Разделение кол.8/кол.9 Т2 по правилу владельца (команда от 08.09.2026):

  «в столбце 8 мы указываем обозначение нужной литературы, а в 9 — только
   страницы из этой литературы, соответствующей теме занятия».

Пример владельца (КТП 01.02, занятие 22):
  БЫЛО:      кол.8 «ОИ1», кол.9 «ОИ 2, с. 136-146»
  ДОЛЖНО:    кол.8 «ОИ2», кол.9 «с. 136-146»

Правила обработки строки занятия Т2:
  • кол.9 = «<ВИД> <N>, с. X-Y» (мои внесения ищи-2/ищи-3) →
      кол.8 = «<ВИД><N>» (без пробела — формат кол.8 эталона 05.02 и примера
      владельца), кол.9 = «с. X-Y».
  • кол.9 «Метод. указания», «Стр. X-Y» (эталон), «9» (вводное), пусто — не трогаются.
  • Обозначение проверяется по спискам литературы Т4/Т5/Т6 («ОИ 1» → «ОИ1»).
Механика — как в fill_pages_all_ishi3.py: zipfile + ElementTree, только w:t
первого run'а (форматирование сохраняется), повторный запуск безопасен.
"""
import json
import re
import zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
for pfx, uri in [
    ('w', W), ('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    ('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
    ('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('w14', 'urn:schemas-microsoft-com:office:word'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('v', 'urn:schemas-microsoft-com:vml'),
    ('o', 'urn:schemas-microsoft-com:office:office'),
]:
    ET.register_namespace(pfx, uri)

FILES = [
    '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 02.01 С-21.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 02.01 С-22.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 02.02 С-21.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 02.02 С-22.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 03.02 С-21.docx',
    '/home/z/my-project/YMK/KTP/КТП МДК 03.02 С-22.docx',
]

PAT = re.compile(r'^(ОИ|ДИ|ЭИ)\s*(\d+)(?:,\s*|\s+)(с\.\s*.+)$')


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


def lit_labels(root):
    """Обозначения литературы из таблиц ОИ/ДИ/ЭИ (T4/T5/T6) → {'ОИ1','ДИ7',…}."""
    body = root.find('w:body', NS)
    labels = set()
    for tbl in body.findall('w:tbl', NS):
        trs = tbl.findall('w:tr', NS)
        if len(trs) < 2:
            continue
        head = cell_text(trs[0])
        if 'Наименование' in head and ('Автор' in head or 'URL' in head):
            for tr in trs[1:]:
                tcs = tr.findall('w:tc', NS)
                if tcs:
                    lbl = cell_text(tcs[0]).replace(' ', '')
                    if re.match(r'^(ОИ|ДИ|ЭИ)\d+$', lbl):
                        labels.add(lbl)
    return labels


def fix_file(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    root = ET.fromstring(data['word/document.xml'].decode('utf-8'))
    body = root.find('w:body', NS)
    t2 = None
    for tbl in body.findall('w:tbl', NS):
        if '№ занятия' in ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:200]:
            t2 = tbl
            break
    assert t2 is not None, f'{path}: Т2 не найдена'

    labels = lit_labels(root)
    assert labels, f'{path}: списки литературы не найдены'
    changed, mism = [], []
    for tr in t2.findall('w:tr', NS):
        tcs = tr.findall('w:tc', NS)
        if len(tcs) < 9:
            continue
        c0 = cell_text(tcs[0])
        if not c0.isdigit():
            continue
        old9 = cell_text(tcs[8])
        if old9.isdigit():          # номерная строка шапки («…8|9|…») или вводное «9»
            continue
        m = PAT.match(old9)
        if not m:
            continue                # «Метод. указания», «Стр. X-Y», пусто — не трогаем
        kind, num, pages = m.group(1), m.group(2), m.group(3).strip()
        new8, new9 = kind + num, pages
        if new8 not in labels:
            mism.append(f'№{c0}: обозначения {new8!r} нет в литературе {path.split("/")[-1]}')
            continue
        old8 = cell_text(tcs[7])
        set_cell_text(tcs[7], new8)
        set_cell_text(tcs[8], new9)
        changed.append((int(c0), old8, old9, new8, new9))
    assert not mism, ';\n  '.join(mism)
    data['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    return changed


TOTAL = 0
for f in FILES:
    ch = fix_file(f)
    TOTAL += len(ch)
    print(f'{f.split("/")[-1]}: исправлено {len(ch)} строк')
    for c0, old8, old9, new8, new9 in ch[:4]:
        print(f'   №{c0}: кол.8 {old8!r}→{new8!r}, кол.9 {old9!r}→{new9!r}')
print(f'\nВСЕГО исправлено строк: {TOTAL}')

# Пометка правила в реестрах/плане (кратко, в _meta)
NOTE = ('col8_col9_rule_2026-09-08: кол.8 = обозначение литературы без пробела (ОИ2/ДИ7/ЭИ1), '
        'кол.9 = только страницы «с. X-Y»; значения реестра вида «ОИ 2, с. X-Y» в docx разделены '
        '(скрипт fix_col8_col9_split.py).')
for jf in ['pages_plan_ishi3', 'lit_pages_05_02', 'lit_pages_01_02',
           'lit_pages_15_01_37', 'lit_pages_03_02']:
    p = f'/home/z/my-project/YMK/scripts/{jf}.json'
    j = json.load(open(p, encoding='utf-8'))
    meta = j.setdefault('_meta', {})
    notes = meta.setdefault('notes', [])
    if NOTE not in notes:
        notes.append(NOTE)
        json.dump(j, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print(f'{jf}.json: _meta дополнен')
