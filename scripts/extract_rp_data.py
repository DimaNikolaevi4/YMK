# -*- coding: utf-8 -*-
"""Универсальный экстрактор РП в промежуточный JSON:
   - РП ПМ.02 (15.01.37): T7, часть 1 (МДК 02.01 Наладка) и часть 2 (Пусконаладка) — темы/занятия/практ.
   - РП ПМ.01 (08.02.09): T6, МДК 01.02 — темы 2.1-2.4, уроки, лаб./практ. работы
   - Литература 4.2.1/4.2.2/4.2.3 обоих РП (полные строки)
   - Заголовки лабораторий раздела 4.1 (для Таблицы 2а)
"""
from docx import Document
from docx.oxml.ns import qn
import json
import re

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
OUT = '/home/z/my-project/scripts/rp_extract.json'


def grid_row(tr):
    cells, gp = [], 0
    for tc in tr.findall(qn('w:tc')):
        el = tc.find(qn('w:tcPr') + '/' + qn('w:gridSpan'))
        span = int(el.get(qn('w:val'))) if el is not None else 1
        ts = tc.findall('.//' + W + 't')
        cells.append((gp, ''.join(x.text or '' for x in ts).strip()))
        gp += span
    return cells


def clean(s):
    return ' '.join((s or '').split())


def parse_lessons(rows, start):
    """Разбирает строки тематического плана начиная с start.
    Возвращает список: {'kind': 'mdk'|'topic'|'block'|'lesson'|'work'|'att', ...}"""
    out = []
    for i in range(start, len(rows)):
        cells = grid_row(rows[i])
        d = {gp: t for gp, t in cells}
        c0 = clean(d.get(0, ''))
        joined = ' '.join(t for gp, t in cells)
        if (c0.startswith('Учебная практика') or c0.startswith('ПМ') or c0.startswith('Производственная практика')) \
                and any(o['kind'] == 'mdk' for o in out):
            break
        nums = [(gp, t) for gp, t in cells if t.isdigit() and gp >= 2]
        # МДК — строка строго начинается с «МДК»
        if re.match(r'^МДК', c0):
            out.append({'kind': 'mdk', 'name': c0,
                        'hours': int(nums[0][1]) if nums else 0,
                        'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        # Тема
        m = re.match(r'^Тема\s*(\d+\.\d+)\.?\s*(.*)', c0)
        if m:
            out.append({'kind': 'topic', 'num': m.group(1), 'name': clean(m.group(2)),
                        'hours': int(nums[0][1]) if nums else 0,
                        'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        # Блок практ./лаб.
        if 'Практические занятия' in joined or 'Лабораторные работы' in joined:
            out.append({'kind': 'block',
                        'hours': int(nums[0][1]) if nums else 0,
                        'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        # Промежуточная аттестация
        if 'Промежуточная аттестация' in c0:
            out.append({'kind': 'att', 'name': c0,
                        'hours': int(nums[0][1]) if nums else 0})
            continue
        # Урок теории: col1 = "N. Название"
        c1 = clean(d.get(1, ''))
        m1 = re.match(r'^(\d+)\.\s*(.+)', c1)
        if m1 and nums:
            out.append({'kind': 'lesson', 'num': int(m1.group(1)), 'name': clean(m1.group(2)),
                        'hours': int(nums[0][1]) if nums else 0,
                        'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        # РП-раскладка типа ПМ.02: col1 = "N" (номер), название — в col2/col3
        if c1.isdigit() and nums:
            name = ''
            for gp, t in cells:
                if 2 <= gp <= 3 and len(t) > len(name):
                    name = clean(t)
            mm = re.match(r'^Практическое занятие\.?\s*(.*)', name, re.I)
            if mm:
                out.append({'kind': 'work', 'type': 'pract', 'num': int(c1),
                            'name': clean(mm.group(1)) or name,
                            'hours': int(nums[0][1]) if nums else 0,
                            'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            else:
                out.append({'kind': 'lesson', 'num': int(c1), 'name': name,
                            'hours': int(nums[0][1]) if nums else 0,
                            'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        # Лаб./практ. работа: col1 = "Лабораторная работа № N..." / "Практическая работа № N..."
        m2 = re.match(r'^((?:Лабораторная|Практическая) работа\s*№\s*(\d+)[.\s]*(.*))', c1, re.I)
        if m2 and nums:
            out.append({'kind': 'work', 'type': 'lab' if 'абораторная' in m2.group(1) else 'pract',
                        'num': int(m2.group(2)), 'name': clean(m2.group(3)) or clean(m2.group(1)),
                        'hours': int(nums[0][1]) if nums else 0,
                        'prep': int(nums[1][1]) if len(nums) > 1 else 0})
            continue
        if c0 or c1:
            out.append({'kind': 'other', 'c0': c0[:60], 'c1': c1[:60], 'nums': [t for _, t in nums]})
    return out


data = {}

# ============ РП ПМ.02 (15.01.37) ============
doc = Document('/home/z/my-project/YMK/RP/02_Оригиналы/15.01.37_РП_ПМ.02_2025 (2).docx')
tbl = doc.tables[7]
rows = tbl._tbl.findall('.//' + W + 'tr')
items = parse_lessons(rows, 0)
# разбить на части по строкам МДК
parts, cur = [], None
for it in items:
    if it['kind'] == 'mdk':
        cur = {'mdk_row': it, 'items': []}
        parts.append(cur)
    elif cur is not None:
        cur['items'].append(it)
data['pm02_150137'] = {'parts': parts}

# Литература ПМ.02
doc_p = [' '.join(p.text.split()) for p in doc.paragraphs]
doc_p = [p for p in doc_p if p]
i1 = next((i for i, t in enumerate(doc_p) if t.startswith('4.2.1')), None)
i2 = next((i for i, t in enumerate(doc_p) if t.startswith('5. Контроль')), None)
lit = doc_p[i1:i2 if i2 else i1 + 30]
data['pm02_150137']['literature'] = lit

# Заголовки лабораторий (строки 4.1 без тире, до «4.2»)
i41 = next((i for i, t in enumerate(doc_p) if t.startswith('4.1.')), None)
labs = []
for t in doc_p[i41:i1 if i1 else i41 + 40]:
    if t.startswith(('Лаборатория', 'Мастерская', 'Электромонтажная', 'Слесарная', 'Кабинет', 'Учебные плакаты')):
        labs.append(t)
data['pm02_150137']['labs'] = labs

# ============ РП ПМ.01 (08.02.09) ============
doc1 = Document('/tmp/conv/08.02.09_ПМ.01_РП_2024.docx')
tbl1 = doc1.tables[6]
rows1 = tbl1._tbl.findall('.//' + W + 'tr')
items1 = parse_lessons(rows1, 0)
# только часть МДК 01.02 (вторая)
parts1, cur = [], None
for it in items1:
    if it['kind'] == 'mdk':
        cur = {'mdk_row': it, 'items': []}
        parts1.append(cur)
    elif cur is not None:
        cur['items'].append(it)
data['pm01_080209'] = {'parts': parts1}

# Литература ПМ.01 (4.2.1 / 4.2.2 / 4.2.3)
p1 = [' '.join(p.text.split()) for p in doc1.paragraphs]
p1 = [p for p in p1 if p]
ia = next((i for i, t in enumerate(p1) if t.startswith('4.2.1.')), None)
ib = next((i for i, t in enumerate(p1) if t.startswith('5. КОНТРОЛЬ') or t.startswith('5. Контроль')), None)
data['pm01_080209']['literature'] = p1[ia:ib if ib else ia + 40]

# Лаборатории 4.1 ПМ.01
ia1 = next((i for i, t in enumerate(p1) if t.startswith('4.1.')), None)
labs1 = [t for t in p1[ia1:ia] if t.startswith(('Лаборатория', 'Мастерская', 'Кабинет', 'Электромонтажная'))]
data['pm01_080209']['labs'] = labs1

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

# Контрольные сводки
for key in ('pm02_150137', 'pm01_080209'):
    print('=' * 70)
    print(key)
    for pi, part in enumerate(data[key]['parts']):
        r = part['mdk_row']
        th = sum(x['hours'] for x in part['items'] if x['kind'] == 'lesson')
        tt = sum(x['hours'] for x in part['items'] if x['kind'] == 'topic')
        pw = sum(x['hours'] for x in part['items'] if x['kind'] == 'work')
        bl = sum(x['hours'] for x in part['items'] if x['kind'] == 'block')
        at = sum(x['hours'] for x in part['items'] if x['kind'] == 'att')
        n_lessons = sum(1 for x in part['items'] if x['kind'] == 'lesson')
        n_works = sum(1 for x in part['items'] if x['kind'] == 'work')
        print(f"  Часть {pi}: '{r['name'][:50]}' объём={r['hours']} практ={r['prep']}")
        print(f"    тем Σ={tt}, уроков {n_lessons} Σ={th}, работ {n_works} Σ={pw}, блоков Σ={bl}, атт Σ={at}")
        others = [x for x in part['items'] if x['kind'] == 'other']
        for o in others:
            print(f"    OTHER: {o}")
print('\nСохранено:', OUT)
