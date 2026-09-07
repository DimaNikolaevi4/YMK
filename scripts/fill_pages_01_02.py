# -*- coding: utf-8 -*-
"""Команда «ищи», финал: заполнение кол.9 Т2 КТП МДК 01.02 реальными страницами
по найденным оглавлениям (Конюхова ОИ 1, Москаленко ОИ 2) + исправление
перепутанных URL Юрайт в таблице ЭИ.

Источники страниц (реальные оглавления):
- Конюхова Е.А. «Электроснабжение объектов» (полный текст издания 320 с.,
  kkgtk.kg; Академия 2021 — стереотипное переиздание той же вёрстки);
- Москаленко В.В. «Системы автоматизированного управления электропривода»
  (ИНФРА-М, 208 с., PDF с elprivod.nmu.org.ua; 2023 — переиздание).

Темы 2.3/2.4 (Попов, Бычков) — оглавления в открытых источниках не найдены,
остаются «Конспект лекций» до решения владельца."""
import zipfile, shutil, os, sys
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
PATH = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'

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

with zipfile.ZipFile(PATH) as z:
    names = z.namelist()
    data = {n: z.read(n) for n in names}

root = ET.fromstring(data['word/document.xml'].decode('utf-8'))
body = root.find('w:body', NS)
tbls = body.findall('w:tbl', NS)

# ---------- 1. Кол.9 Т2: страницы ----------
PAGES = {
    1: 'ОИ 1, с. 285-292',   # Этапы развития АСУ ТП — Конюхова Гл.20.1-20.4
    2: 'ОИ 1, с. 149-153',   # Структура АСКУЭ — Гл.11 расход и потери э/э
    3: 'ОИ 1, с. 46-55',     # Характеристики устройств — Гл.4 оборудование подстанций
    4: 'ОИ 1, с. 10-16',     # Функции АСДУ — Гл.1.4-1.7 управление ЭЭС
    5: 'ОИ 1, с. 293-304',   # Противоаварийная защита — Гл.20.5-20.11
    6: 'ОИ 1, с. 117-129',   # Технический учёт — Гл.8 графики нагрузок
    17: 'ОИ 2, с. 87-101',   # АПВ — Москаленко Гл.3.1-3.2
    18: 'ОИ 2, с. 101-114',  # АВР — Гл.3.3
    19: 'ОИ 2, с. 114-119',  # АПВ двустороннее питание — Гл.3.4-4.1
    20: 'ОИ 2, с. 9-14',     # Регулирование напряжения — Гл.1.1-1.3
    21: 'ОИ 2, с. 119-136',  # Методы регулирования — Гл.4.2-4.3
    22: 'ОИ 2, с. 136-146',  # Автоматика фидеров — Гл.4.3-4.4
    23: 'ОИ 2, с. 179-186',  # Управление освещением — Гл.5.1-5.2
    24: 'ОИ 2, с. 186-193',  # Дизель-генератор — Гл.5.2-5.3
}


def cell_text(tc):
    return ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()


def set_cell_text(tc, new_text):
    """Пишет new_text в первый w:t ячейки, очищает остальные прогоны."""
    ts = tc.findall('.//w:t', NS)
    assert ts, 'нет w:t в ячейке'
    first = True
    for t in ts:
        if first:
            t.text = new_text
            # пробел-сохранение для составных значений
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            first = False
        else:
            t.text = ''


t2 = None
for tbl in tbls:
    head = ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:150]
    if '№ занятия' in head:
        t2 = tbl
        break
assert t2 is not None, 'Т2 не найдена'

filled = []
for tr in t2.findall('w:tr', NS):
    tcs = tr.findall('w:tc', NS)
    if len(tcs) < 9:
        continue
    c0 = cell_text(tcs[0])
    if c0.isdigit() and int(c0) in PAGES and cell_text(tcs[8]) == 'Конспект лекций':
        set_cell_text(tcs[8], PAGES[int(c0)])
        filled.append(int(c0))

assert sorted(filled) == sorted(PAGES), f'заполнено {sorted(filled)}, ожидалось {sorted(PAGES)}'
print(f'Кол.9 Т2: заполнено {len(filled)} уроков реальными страницами: {sorted(filled)}')

# ---------- 2. ЭИ: обмен URL (в файле владельца ссылки перепутаны) ----------
t6 = tbls[6]  # таблица ЭИ (последняя)
head6 = ''.join(t.text or '' for t in t6.iter(f'{{{W}}}t'))[:200]
assert 'URL' in head6, 'таблица 6 не похожа на ЭИ'
url_map = {'https://urait.ru/bcode/513864': 'https://urait.ru/bcode/517783',
           'https://urait.ru/bcode/517783': 'https://urait.ru/bcode/513864'}
swapped = 0
for tr in t6.findall('w:tr', NS):
    for tc in tr.findall('w:tc', NS):
        txt = cell_text(tc)
        if txt in url_map:
            set_cell_text(tc, url_map[txt])
            swapped += 1
assert swapped == 2, f'ожидались 2 замены URL, сделано {swapped}'
print('ЭИ: URL Юрайт обменяны местами (513864 <-> 517783) — Климова/Бредихин исправлены')

# ---------- Сохранение ----------
data['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')
with zipfile.ZipFile(PATH, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.writestr(n, data[n])
print('Сохранено:', PATH)
shutil.copy(PATH, '/home/z/my-project/download/КТП МДК 01.02 М-21.docx')
print('Копия: /home/z/my-project/download/КТП МДК 01.02 М-21.docx')

# ---------- Контроль ----------
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables
_, tb = doc_text_and_tables(PATH)
t2b = tb[2]
print('\nКонтроль кол.9 (первые заполненные):')
shown = 0
for row in t2b[3:]:
    if len(row) > 8 and row[0]['text'].strip().isdigit() and 'ОИ' in row[8]['text']:
        print(f"  №{row[0]['text']:>3}: {row[8]['text']}")
        shown += 1
        if shown >= 6:
            break
print('ЭИ строки:')
for row in tb[6]:
    print('  ', ' | '.join(c['text'][:55] for c in row)[:160])
