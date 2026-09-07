# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 5: посеместровая структура Т2 (семестры, теория/практ, КДЗ)."""
import sys, re
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

PATH = '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx'
_, tables = doc_text_and_tables(PATH)
t2 = tables[2]

cur_sem = None
tot = {'3': [0, 0, 0], '4': [0, 0, 0]}  # [часы теор, часы практ, занятия]
print('=== Т2: семестровые маркеры и типы занятий ===')
for i, row in enumerate(t2[3:], 3):
    c0 = row[0]['text'].strip()
    name = row[1]['text'].strip() if len(row) > 1 else ''
    low = name.lower()
    if 'семестр' in low:
        m = re.search(r'(\d)', name)
        cur_sem = m.group(1) if m else '?'
        print(f'строка {i}: >>> {name[:70]}')
        continue
    if not c0.isdigit():
        if name:
            print(f'строка {i}: [служебная] {name[:70]}')
        continue
    # часы: кол.3 (обязательная нагрузка?) — структура: 0=№, 1=наименование, 2..? = часы
    nums = [c['text'].strip() for c in row[2:] if c['text'].strip().isdigit()]
    hours = nums[0] if nums else '?'
    is_kdz = 'зачет' in low or 'зачёт' in low
    is_prakt = 'практ' in low or 'лаборатор' in low
    kind = 'КДЗ' if is_kdz else ('ПРАКТ' if is_prakt else 'ТЕОР')
    if cur_sem in tot:
        h = int(hours) if hours.isdigit() else 0
        tot[cur_sem][0 if kind == 'ТЕОР' else (1 if kind in ('ПРАКТ',) else 2)] += 0  # placeholder
        if kind == 'ТЕОР':
            tot[cur_sem][0] += h
        elif kind == 'ПРАКТ':
            tot[cur_sem][1] += h
        else:
            tot[cur_sem][2] += h
    if c0.isdigit() and (int(c0) <= 3 or int(c0) in (11, 12, 16, 17, 18, 19, 20, 24, 25, 51, 52)):
        print(f'  №{c0} [{kind}] часы={hours}: {name[:65]}')

print()
for s, (t, p, k) in tot.items():
    print(f'Семестр {s}: теория {t} ч, практ {p} ч, КДЗ {k} ч, итого {t+p+k} ч')
print(f'Всего: теория {sum(tot[s][0] for s in tot)} ч + КДЗ {sum(tot[s][2] for s in tot)} ч, практ {sum(tot[s][1] for s in tot)} ч')

# Похожий дамп для эталона 05.02 (для сравнения логики Т1/Т2)
print('\n=== ЭТАЛОН 05.02: семестры Т1 (строки 5-11) ===')
_, tables_et = doc_text_and_tables('/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx')
for i, row in enumerate(tables_et[1]):
    vals = [c['text'][:24] for c in row]
    if any(v for v in vals):
        print(f'  {i}: {vals}')
