# -*- coding: utf-8 -*-
"""Сравнение загруженного исправленного КТП с нашей генерацией и эталоном 05.02."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

UP = '/home/z/my-project/upload/КТП МДК 01.02 М 2 курс.docx'
OURS = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'
ET = '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx'


def tables_of(path):
    return doc_text_and_tables(path)[1]


def find_table(tables, *patterns):
    """Найти таблицу, чьи первые строки содержат все паттерны."""
    for rows in tables:
        head = ' '.join(c['text'] for r in rows[:5] for c in r)
        if all(p.lower() in head.lower() for p in patterns):
            return rows
    return None


def row_texts(rows, i):
    return [c['text'] for c in rows[i]] if i < len(rows) else []


up_t, ours_t, et_t = tables_of(UP), tables_of(OURS), tables_of(ET)
t2_up = find_table(up_t, '№ занятия')
t2_our = find_table(ours_t, '№ занятия')
t2_et = find_table(et_t, '№ занятия')

print('=== ШАПКА Т2 ЭТАЛОН 05.02 (%d строк) ===' % len(t2_et))
for i in range(min(5, len(t2_et))):
    print(i, [c['text'][:45] for c in t2_et[i]])
print()
print('=== ШАПКА Т2 НАША 01.02 (%d строк) ===' % len(t2_our))
for i in range(min(5, len(t2_our))):
    print(i, [c['text'][:45] for c in t2_our[i]])
print()
print('=== ШАПКА Т2 ЗАГРУЖЕННАЯ (%d строк) ===' % len(t2_up))
for i in range(min(5, len(t2_up))):
    print(i, [c['text'][:45] for c in t2_up[i]])
print()

print('=== ШИРИНА КОЛОНОК (gridSpan) шапки Т2: эталон / наша / загруж. ===')
for name, t in (('эталон', t2_et), ('наша', t2_our), ('загруж', t2_up)):
    spans = [sum(c['span'] for c in r) for r in t[:5]]
    print(f'{name}: строк в первых 5: {spans}, всего строк {len(t)}')
print()

print('=== diff Т2 тело: наша vs загруженная ===')
n = max(len(t2_our), len(t2_up))
diffs = 0
for i in range(5, n):
    if row_texts(t2_our, i) != row_texts(t2_up, i):
        diffs += 1
        if diffs <= 15:
            print(f'строка {i}:')
            print('  НАША :', [x[:55] for x in row_texts(t2_our, i)])
            print('  НОВАЯ:', [x[:55] for x in row_texts(t2_up, i)])
print('всего различающихся строк тела:', diffs)

# семестровые маркеры и КДЗ в загруженной Т2
print()
print('=== Строки с "семестр" (без номера) и КДЗ/зачет в загруженной Т2 ===')
for i, r in enumerate(t2_up):
    txt = ' '.join(c['text'] for c in r)
    low = txt.lower()
    if ('семестр' in low and not txt.strip().isdigit()) or 'зачет' in low or 'зачёт' in low:
        print(i, [c['text'][:60] for c in r][:4])
