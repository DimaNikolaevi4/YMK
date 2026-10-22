# -*- coding: utf-8 -*-
"""Финальная сверка: каждый урок КТП МДК 02.01 КИП-31 = пункту РП 3.2 (15.01.31).
+ реестр lit_pages_15_01_31.json (метки достоверности кол.9)."""
import json, re, sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables

YMK = '/home/z/my-project/YMK'
RP = f'{YMK}/RP/15.01.31_РП_ПМ.02_2024.docx'
KTP = f'{YMK}/KTP/КТП МДК 02.01 КИП-31.docx'
DATA = f'{YMK}/scripts/data_02_01_kip31.json'

def celltext(v):
    return str(v.get('text', '')).strip() if isinstance(v, dict) else str(v).strip()

def norm(s):
    return re.sub(r'\s+', ' ', s).strip().lower()

# --- РП 3.2: пункты МДК 02.01 ---
texts, tables = doc_text_and_tables(RP)
t32 = None
for t in tables:
    rows = [[celltext(v) for v in ([r[k] for k in sorted(r.keys())] if isinstance(r, dict) else r)] for r in t]
    if rows and rows[0] and rows[0][0].startswith('Наименование разделов'):
        t32 = rows
        break
start = next(i for i, r in enumerate(t32) if r[0].startswith('МДК 02.01'))
end = next(i for i, r in enumerate(t32) if r[0].startswith('МДК 02.02'))
rp_items = []  # (норм.текст, часы)
for r in t32[start + 1:end]:
    r = (r + [''] * 5)[:5]
    if r[0].startswith('Тема '):
        continue
    if r[1] and not r[0] and r[1] not in ('Содержание',) and not r[1].startswith('Практич'):
        try:
            h = int(r[3])
        except ValueError:
            h = 0
        rp_items.append((norm(r[2]), h))
print(f'РП 3.2: {len(rp_items)} пунктов')

# --- КТП Т2: уроки ---
data = json.load(open(DATA, encoding='utf-8'))
ktp_items = []
_, tables_k = doc_text_and_tables(KTP)
for t in tables_k:
    rows = [[celltext(v) for v in ([r[k] for k in sorted(r.keys())] if isinstance(r, dict) else r)] for r in t]
    if not rows or not rows[0] or rows[0][0] != '№ занятия':
        continue
    for r in rows:
        r = (r + [''] * 11)[:11]
        # пропуск номерной строки шапки («1|2|3|…|11»)
        if [r[i] for i in range(11)] == [str(i + 1) for i in range(11)]:
            continue
        if r[0].isdigit() and r[1] and not r[1].startswith(('Тема', 'МДК', 'Итого')):
            if 'зачет' in r[1].lower():
                continue
            try:
                h = int(r[2]) if r[2] else 0
            except ValueError:
                h = 0
            # у ПЗ в КТП есть префикс «Практическое занятие № X» — в РП его нет
            clean = re.sub(r'^практическое занятие №\s*[\d\-–]+\s+', '', norm(r[1]))
            ktp_items.append((clean, h))
    break
print(f'КТП Т2: {len(ktp_items)} уроков')

ok = bad = 0
for i, (kt, kh) in enumerate(ktp_items):
    matched = any(rt.startswith(kt[:60]) or kt.startswith(rt[:60]) for rt, rh in rp_items)
    hours_ok = any(rh == kh for rt, rh in rp_items if rt.startswith(kt[:60]) or kt.startswith(rt[:60]))
    if matched and hours_ok:
        ok += 1
    else:
        bad += 1
        print(f'  ❌ №{i+1} ({kh} ч): {kt[:70]}')
print(f'Сверка уроков ↔ РП 3.2: {ok}/{len(ktp_items)} {"✅" if bad == 0 else "❌"}')

# --- Реестр lit_pages_15_01_31.json ---
registry = {
    "_meta": {
        "file": "scripts/data_02_01_kip31.json",
        "rule": "col8_col9_rule_2026-09-08: кол.8 = обозначение источника без пробела, кол.9 = только «с. X-Y»",
        "notes": [
            "КТП МДК 02.01 КИП-31 (15.01.31) создан по RP/15.01.31_РП_ПМ.02_2024.docx (09.09.2026).",
            "Список «Основная» в РП пуст — ОИ 1-3 назначены свежие издания Юрайт-2025 из раздела 4.2 РП (проверка 10а ✓): Рачков 2025 (182 с.), Бородин/Андреев 2025 (377 с.), Колосов 2025 (331 с.).",
            "ДИ 1-5 — печатные издания раздела 4.2 РП (2012-2003), ДИ 6-7 — ГОСТ 21.208-2013 / ГОСТ 21.408-2013.",
            "ЭИ 1 — Троценко и др., Юрайт 2025, URL https://urait.ru/bcode/563647 (из РП, verified).",
            "Страницы кол.9: 🟡 approx — равномерное распределение по объёму книги в пределах темы; граница не превышает объём (Рачков ≤182, Бородин/Андреев ≤377, Колосов ≤331, Селевцов ≤352, Шишмарёв ≤352, Карташов ≤540, Бородин/Судник ≤344, Горошков ≤304).",
            "verified: ДИ1 «с. 3-12» (Селевцов, pages_plan_ishi3.json) и ДИ2 (продолжение verified-фрагмента «с. 3-20» Шишмарёва АТП из lit_pages_03_02.json → темы 1.2/1.7 21-36/37-52/53-68/69-84/85-100 — продолжения, 🟡).",
            "ПЗ (24 занятия) — «№3, 2а» / «Метод. указания» / «Защита пр.раб.» — эталонное правило (не трогать).",
            "Урок №1 — «Вх.контр», теория — «Текущий контроль».",
        ],
        "pages": {
            "ДИ1": {"scheme": "verified_start_3-12_then_approx", "lessons": [1, 20, 21, 22, 23]},
            "ДИ2": {"scheme": "verified_3-20_continued", "lessons": [10, 11, 56, 57, 58]},
            "ДИ4": {"scheme": "approx", "lessons": [32, 33, 61, 62, 63, 64, 65, 66]},
            "ДИ5": {"scheme": "approx", "lessons": [15, 27, 28, 29, 34]},
            "ДИ7": {"scheme": "approx_pdf", "lessons": [3, 4]},
            "ОИ1": {"scheme": "approx", "lessons": [6, 25, 26, 31]},
            "ОИ2": {"scheme": "approx", "lessons": [2, 5, 12, 13, 14, 17, 30, 50, 59, 60, 67, 68]},
            "ОИ3": {"scheme": "approx", "lessons": [16, 24, 47, 48, 49]},
        },
        "todo": "уточнить диапазоны по сканам книг (приоритет: Бородин/Андреев 377 с., Селевцов 352 с.)",
    }
}
with open(f'{YMK}/scripts/lit_pages_15_01_31.json', 'w', encoding='utf-8') as f:
    json.dump(registry, f, ensure_ascii=False, indent=2)
print('Реестр lit_pages_15_01_31.json создан')
