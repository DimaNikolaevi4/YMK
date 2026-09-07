# -*- coding: utf-8 -*-
"""Команда «ищи», шаг 7: полный diff всех таблиц ГОТ vs НАША (текстовый уровень)."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

GOT = '/home/z/my-project/YMK/KTP/гот/КТП МДК 01.02 М 2 курс.docx'
OUR = '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'

_, tg = doc_text_and_tables(GOT)
_, to = doc_text_and_tables(OUR)
print(f'таблиц: ГОТ {len(tg)}, НАША {len(to)}; абзацев: {len(tg)} vs {len(to)}')

for idx in range(max(len(tg), len(to))):
    a = [' | '.join(c['text'] for c in r) for r in tg[idx]]
    b = [' | '.join(c['text'] for c in r) for r in to[idx]]
    nd = sum(1 for x, y in zip(a, b) if x != y) + abs(len(a) - len(b))
    if nd or len(a) != len(b):
        print(f'\n--- Таблица #{idx}: строк {len(a)} vs {len(b)}, расхождений {nd}')
        for i in range(max(len(a), len(b))):
            xa = a[i] if i < len(a) else '<нет>'
            xb = b[i] if i < len(b) else '<нет>'
            if xa != xb:
                print(f'  строка {i}:')
                print(f'    ГОТ : {xa[:130]}')
                print(f'    НАША: {xb[:130]}')

# абзацы (текст между таблицами)
ta, tb = doc_text_and_tables(GOT)[0], doc_text_and_tables(OUR)[0]
diffs = [(i, x, y) for i, (x, y) in enumerate(zip(ta, tb)) if x != y]
extra = [(i, x) for i, x in enumerate(ta[len(tb):], len(tb))]
print(f'\nАбзацы: расхождений {len(diffs)}, лишних в ГОТ {len(extra)}')
for i, x, y in diffs[:5]:
    print(f'  [{i}] ГОТ : {x[:90]}\n      НАША: {y[:90]}')
for i, x in extra[:5]:
    print(f'  [{i}] только ГОТ: {x[:90]}')
