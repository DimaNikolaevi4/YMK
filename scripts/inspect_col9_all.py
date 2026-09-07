# -*- coding: utf-8 -*-
"""Инспекция: кол.9 (и кол.8 для контекста) во всех 8 КТП."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

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

for path in FILES:
    _, tb = doc_text_and_tables(path)
    print('=' * 100)
    print(path.split('/')[-1])
    t2 = None
    for i, rows in enumerate(tb):
        head = ' '.join(c['text'] for c in rows[0])
        if '№ занятия' in head:
            t2 = rows
            break
    if t2 is None:
        print('  !! Т2 не найдена')
        continue
    print(f'  строк в Т2: {len(t2)}')
    for row in t2:
        c0 = row[0]['text'].strip()
        c3 = row[2]['text'].strip()[:60] if len(row) > 3 else ''
        c8 = row[8]['text'].strip() if len(row) > 8 else ''
        if c0.isdigit():
            print(f'  №{c0:>3} | {c3:<60} | кол9: {c8}')
        elif 'Тема' in c0 or 'Итого' in c0 or 'Дифф' in c0 or 'зач' in c0.lower() or 'контроль' in c0.lower():
            print(f'  [{c0[:30]:<30} | {c3:<60} | кол9: {c8[:50]}')
