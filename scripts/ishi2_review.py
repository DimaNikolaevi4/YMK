#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ищи-2, шаг 2: сводка результатов поиска — топ-хиты по каждому запросу."""
import json, os, re

BASE = '/home/z/my-project/YMK/docs/search/ishi2'

# Домены, наиболее перспективные для оглавлений
GOOD = re.compile(r'(pdf|studfile|insappia|bookvoed|labirint|chitai|books\.ru|rusneb|elibrary|'
                  r'lanbook|urait|znanium|academia-moscow|infra-m|phoenixrosto|catalog|biblio|'
                  r'docplayer|vdoc|helpiks|studme|electric|elec\.ru|twirpx|nashaucheba|uchebnik|'
                  r'free|library|lib\.)', re.I)

for f in sorted(os.listdir(BASE)):
    if not f.endswith('.json'):
        continue
    try:
        data = json.load(open(os.path.join(BASE, f)))
    except Exception as e:
        print(f'{f}: ОШИБКА {e}')
        continue
    print(f'\n{"="*90}\n### {f}  ({len(data)} результатов)')
    for it in data:
        mark = '★' if GOOD.search(it.get('url','')) else ' '
        name = (it.get('name') or '')[:80]
        url = (it.get('url') or '')[:100]
        snip = re.sub(r'\s+', ' ', (it.get('snippet') or ''))[:120]
        print(f' {mark} {name}\n    {url}\n    {snip}')
