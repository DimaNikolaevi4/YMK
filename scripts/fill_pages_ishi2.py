# -*- coding: utf-8 -*-
"""Команда «ищи-2», финал: заполнение кол.9 Т2 реальными страницами по ОГЛАВЛЕНИЯМ,
ПОДТВЕРЖДЁННЫМ официальными источниками:

1) КТП МДК 05.02 М2 курс (эталон) — 18 теоретических занятий по реальному
   оглавлению Нестеренко В.М., Мысьянов А.М. «Технология электромонтажных работ»
   (Академия, 592 с., PDF dist.bsut.by; оглавление разобрано командой «ищи»).
   Формат «Стр. X-Y» — как в эталоне. Темы 2.1, 2.6, 2.8 НЕ трогаем (в Нестеренко
   не покрыты — нужны оглавления Сидоровой/Ярочкиной).

2) КТП МДК 01.02 М-21 — занятие 35 «КИП инженерных систем МКД» по ПОЛНОМУ
   оглавлению Полуянович Н.К., Дубяго М.Н. «Эксплуатация электротехнических
   систем объектов ЖКХ» (Феникс, 2020, 158 с.), снятому с официальной карточки
   издательства phoenixbooks.ru (раздел «Методы, средства измерения и принцип
   действия контрольно-измерительных приборов», с. 97–104).

Всё остальное (ОИ 3 Попов, ОИ 4 Бычков, ОИ 15.01.37 и т.д.) — оглавлений
в открытых источниках нет, остаётся как есть до сканов от владельца.
Повторный запуск безопасен (assert на исходные значения).
"""
import zipfile, shutil, os, sys
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


def fill_col9(path, pages, expect_old, label):
    """pages: {№ занятия: новый текст}; expect_old: {№: ожидаемый текущий текст}."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    root = ET.fromstring(data['word/document.xml'].decode('utf-8'))
    body = root.find('w:body', NS)
    tbls = body.findall('w:tbl', NS)
    t2 = None
    for tbl in tbls:
        head = ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:200]
        if '№ занятия' in head:
            t2 = tbl
            break
    assert t2 is not None, f'{label}: Т2 не найдена'
    filled = []
    for tr in t2.findall('w:tr', NS):
        tcs = tr.findall('w:tc', NS)
        if len(tcs) < 9:
            continue
        c0 = cell_text(tcs[0])
        if c0.isdigit() and int(c0) in pages:
            old = cell_text(tcs[8])
            assert old == expect_old[int(c0)], f'{label} №{c0}: текущее «{old}» != ожидаемое «{expect_old[int(c0)]}»'
            set_cell_text(tcs[8], pages[int(c0)])
            filled.append(int(c0))
    assert sorted(filled) == sorted(pages), f'{label}: заполнено {sorted(filled)}, ожидалось {sorted(pages)}'
    data['word/document.xml'] = ET.tostring(root, xml_declaration=True, encoding='UTF-8')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    print(f'{label}: заполнено {len(filled)} ячеек кол.9: {sorted(filled)}')
    return sorted(filled)


# ---------- 1. КТП МДК 05.02 (эталон) — Нестеренко, реальное оглавление ----------
P_0502 = 'Нестеренко (реальное оглавление, «ищи»):'
PAGES_0502 = {
    # Тема 2.2 Сборка и монтаж осветительных электроустановок, аппаратов защиты и ПРА (6 зан.)
    3:  'Стр. 100-115',   # Гл.4 «Основные сведения об электрическом освещении» (100-140)
    4:  'Стр. 115-125',
    5:  'Стр. 125-140',
    6:  'Стр. 166-183',   # Гл.6 «Монтаж светильников…» (166-193)
    7:  'Стр. 183-188',
    8:  'Стр. 188-193',
    # Тема 2.3 Монтаж кабельных линий, шинопроводов, троллейных линий (4 зан.)
    18: 'Стр. 316-330',   # Гл.10 «Кабельные линии до 1 кВ» (316-363)
    19: 'Стр. 330-343',
    20: 'Стр. 355-363',
    21: 'Стр. 413-424',   # Гл.12 «Шинопроводы и троллейные линии» (413-443)
    # Тема 2.4 Монтаж защитного заземления и зануления (1 зан.)
    25: 'Стр. 141-152',   # Гл.5 «Монтаж устройств защитного заземления» (141-165)
    # Тема 2.5 Монтаж электрических машин и силовых трансформаторов (4 зан.)
    27: 'Стр. 444-468',   # Гл.13 «Трансформаторы» (444-528)
    28: 'Стр. 468-503',
    29: 'Стр. 503-518',
    30: 'Стр. 518-529',   # Гл.14 «Монтаж КТП» (529-540)
    # Тема 2.7 Ремонт воздушных и кабельных линий (3 зан.)
    40: 'Стр. 364-375',   # Гл.11 «Воздушные линии до 1 кВ» (364-412)
    41: 'Стр. 375-391',
    42: 'Стр. 391-412',
}
OLD_0502 = {
    3: 'Стр. 79-93', 4: 'Стр. 94-108', 5: 'Стр. 109-123', 6: 'Стр. 124-138',
    7: 'Стр. 139-153', 8: 'Стр. 154-167',
    18: 'Стр. 228-241', 19: 'Стр. 242-255', 20: 'Стр. 256-269', 21: 'Стр. 270-283',
    25: 'Стр. 124-167',
    27: 'Стр. 203-218', 28: 'Стр. 219-234', 29: 'Стр. 263-273', 30: 'Стр. 274-284',
    40: 'Стр. 285-296', 41: 'Стр. 297-308', 42: 'Стр. 309-320',
}
F_0502 = '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx'
fill_col9(F_0502, PAGES_0502, OLD_0502, 'КТП 05.02 (эталон)')
shutil.copy(F_0502, '/home/z/my-project/download/КТП МДК 05.02 М2 курс.docx')
print('Копия: /home/z/my-project/download/КТП МДК 05.02 М2 курс.docx')

# ---------- 2. КТП МДК 01.02 М-21 — занятие 35 по Полуянович (Феникс, verified) ----------
PAGES_0102 = {35: 'ДИ 1, с. 97-104'}
OLD_0102 = {35: 'Конспект лекций'}
F_0102 = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'
fill_col9(F_0102, PAGES_0102, OLD_0102, 'КТП 01.02 М-21')
shutil.copy(F_0102, '/home/z/my-project/download/КТП МДК 01.02 М-21.docx')
print('Копия: /home/z/my-project/download/КТП МДК 01.02 М-21.docx')

# ---------- Контроль ----------
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

for path, idx_t2 in [(F_0502, None), (F_0102, 2)]:
    _, tb = doc_text_and_tables(path)
    # ищем Т2 заново
    for i, rows in enumerate(tb):
        head = ' '.join(c['text'] for c in rows[0])
        if '№ занятия' in head:
            idx_t2 = i
            break
    print(f'\nКонтроль {path.split("/")[-1]}:')
    for row in tb[idx_t2][3:]:
        c0 = row[0]['text'].strip()
        if c0.isdigit() and int(c0) in {**PAGES_0502, **PAGES_0102}:
            print(f"  №{c0:>3}: {row[8]['text']}")
