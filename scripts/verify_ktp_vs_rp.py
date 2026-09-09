# -*- coding: utf-8 -*-
"""Финальная сверка: каждый урок КТП Т2 = пункту «Содержание» РП 3.2 (нового)."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables, flat

BASE = '/home/z/my-project/YMK'
RP = f'{BASE}/RP/15.01.37_РП_ПМ.02_2025.docx'


def rp_content_items():
    texts, tables = doc_text_and_tables(RP)
    for rows in tables:
        j = ' '.join(c['text'] for rr in rows[:2] for c in rr)
        if 'Содержание учебного материала' in j:
            items, cur_theme = [], None
            for rr in rows:
                f2 = (flat(rr) + [''] * 8)[:8]
                c0, rest = f2[0].strip(), [x.strip() for x in f2[1:]]
                if c0.startswith('МДК') or c0.startswith('Тема'):
                    cur_theme = c0
                txts = [x for x in rest if len(x) > 40]
                nums = [x for x in rest if x.isdigit()]
                if txts and nums:
                    items.append((cur_theme, txts[0]))
            return items
    return []


norm = lambda s: ' '.join(s.replace('\u00a0', ' ').split()).lower().rstrip('.')

ok = bad = 0
rp_items = rp_content_items()
for f in ['KTP/КТП МДК 02.01 С-21.docx', 'KTP/КТП МДК 02.02 С-21.docx']:
    texts, tables = doc_text_and_tables(f'{BASE}/{f}')
    print('=' * 95)
    print(f)
    ktp_lessons = []
    for rows in tables:
        head = ' '.join(c['text'] for rr in rows[:3] for c in rr)
        if '№ занятия' in head:
            ncols = max(sum(c['span'] for c in rr) for rr in rows)
            for rr in rows:
                f2 = (flat(rr) + [''] * ncols)[:ncols]
                if f2[0].strip().isdigit() and not f2[1].strip().isdigit():
                    name = f2[1].strip()
                    if 'зачет' not in name.lower():
                        ktp_lessons.append((f2[0].strip(), name, f2[2].strip()))
    # уроки (не ПЗ) сравниваем с пунктами содержания РП
    rp_pool = list(rp_items)
    for num, name, hrs in ktp_lessons:
        if name.startswith('Практическое занятие'):
            continue
        hit = None
        for it in rp_pool:
            if norm(it[1]) == norm(name):
                hit = it
                break
        if hit:
            ok += 1
            rp_pool.remove(hit)
            print(f'  ✅ №{num:>2} ({hrs} ч) = РП [{hit[0][:28]}] {name[:52]}...')
        else:
            # частичное совпадение?
            part = [it for it in rp_pool if norm(name)[:60] in norm(it[1])]
            bad += 1
            print(f'  ❌ №{num:>2} ({hrs} ч) НЕ совпадает: {name[:70]}...' +
                  (f' ← частично [{part[0][0][:20]}]' if part else ''))
print('=' * 95)
print(f'Уроков сверено: ✅ {ok}, ❌ {bad}')
