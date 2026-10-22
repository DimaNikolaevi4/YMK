# -*- coding: utf-8 -*-
"""Верификация КТП 02.01/02.02 С-21/С-22 против нового РП ПМ.02 (15.01.37)
и ответа владельца «46+2 делаем по рп».
Выводит: Т1 обоих КТП, титульные часы, Т3.4 РП (полностью), сводку РП 3.2 (кол.4).
"""
import re, sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

YMK = '/home/z/my-project/YMK'
KTP1 = f'{YMK}/KTP/КТП МДК 02.01 С-21.docx'
KTP2 = f'{YMK}/KTP/КТП МДК 02.02 С-21.docx'
RP   = f'{YMK}/RP/15.01.37_РП_ПМ.02_2025.docx'

def rows_of(table):
    """table = список строк; строка = словарь {col: {'text','span'}} или {col: str}."""
    out = []
    for r in table:
        if isinstance(r, dict):
            keys = sorted(r.keys())
            vals = []
            for k in keys:
                v = r[k]
                if isinstance(v, dict):
                    vals.append(str(v.get('text', '')).strip().replace('\n', '⏎'))
                else:
                    vals.append(str(v).strip().replace('\n', '⏎'))
            out.append(vals)
        else:
            out.append([str(v).strip().replace('\n', '⏎') for v in r])
    return out

def show_t1(path, label):
    _, tables = doc_text_and_tables(path)
    for t in tables:
        rows = rows_of(t)
        if rows and rows[0] and rows[0][0].startswith('Междисциплинарный курс'):
            print(f'\n=== Т1 {label} ({len(rows)} строк) ===')
            for i, row in enumerate(rows):
                nz = [f'{j}:"{v}"' for j, v in enumerate(row) if v]
                print(f'r{i}: ' + (' | '.join(nz) if nz else '(все пусты)'))
            break

def show_title_hours(path, label):
    texts, _ = doc_text_and_tables(path)
    print(f'\n=== Титул {label}: строки с «час» ===')
    for s in texts:
        if 'час' in s.lower() and len(s) < 300:
            print('  ', s.strip())

print('#' * 70)
print('КТП Т1 и титулы')
show_t1(KTP1, '02.01 С-21')
show_t1(KTP2, '02.02 С-21')
show_title_hours(KTP1, '02.01')
show_title_hours(KTP2, '02.02')

print('\n' + '#' * 70)
print('РП ПМ.02 (15.01.37): таблицы с часами')
texts_rp, tables_rp = doc_text_and_tables(RP)
for ti, t in enumerate(tables_rp):
    rows = rows_of(t)
    if not rows:
        continue
    flat2 = ' '.join(' '.join(r) for r in rows[:3])
    flatall = ' '.join(' '.join(r) for r in rows)
    if 'Наименование разделов' in flat2 or ('МДК' in flatall and 'КДЗ' in flatall) or \
       ('Раздел' in flat2 and 'час' in flatall.lower()):
        print(f'\n=== РП таблица #{ti} ({len(rows)} строк) — кандидат Т3.4 ===')
        for i, row in enumerate(rows):
            nz = [f'{j}:"{v}"' for j, v in enumerate(row) if v]
            print(f'r{i}: ' + (' | '.join(nz) if nz else '(пусто)'))

# РП 3.2: темы 1.x.x
print('\n=== РП 3.2: темы и кол.4 ===')
for ti, t in enumerate(tables_rp):
    rows = rows_of(t)
    cnt = sum(1 for r in rows if r and re.match(r'^1\.[1-4]\.\d+', r[0]))
    if cnt >= 3:
        print(f'\n-- РП таблица #{ti}: {cnt} тем --')
        for r in rows:
            if r and re.match(r'^1\.[1-4]\.\d+', r[0]):
                print(f'   {r[0][:70]:70s} | ' + ' | '.join(f'"{x}"' for x in r[1:6]))
            elif r and re.match(r'^1\.[1-4]\b', r[0]):
                print(f'   ### ТЕМА {r[0][:70]:70s} | ' + ' | '.join(f'"{x}"' for x in r[1:6]))
