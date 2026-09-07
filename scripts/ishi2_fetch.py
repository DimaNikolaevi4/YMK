#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ищи-2, шаг 3: вытягивание контента перспективных страниц через page_reader."""
import json, subprocess, os, time, sys

BASE = '/home/z/my-project/YMK'
OUT = os.path.join(BASE, 'docs/search/ishi2/pages')
os.makedirs(OUT, exist_ok=True)

URLS = {
    # Карточки Академии (могут содержать «Содержание»)
    'acad_sidorova':      'https://academia-moscow.ru/catalogue/info.php?id=883969',
    'acad_yarochkina_nal':'https://academia-moscow.ru/catalogue/info.php?id=586863',
    'acad_nesterenko_pod':'https://academia-moscow.ru/catalogue/info.php?id=520734',
    'acad_bychkov_er':    'https://academia-moscow.ru/catalogue/info.php?id=520714',
    'acad_shashkova_ch2': 'https://academia-moscow.ru/catalogue/info.php?id=502641',
    'acad_grigorieva':    'https://academia-moscow.ru/catalogue/info.php?id=511376',
    'acad_andreev':       'https://academia-moscow.ru/catalogue/info.php?id=181952',
    # Лань: читалки ebsReader
    'lan_apollonsky':     'https://lanbook.com/ebsReader.php?id=1995',
    'lan_poluyanovich_mn':'https://lanbook.com/ebsReader.php?id=2169',
    # Феникс: карточка Полуянович ЖКХ (в выдаче упоминалось «Содержание»)
    'phx_poluyanovich':   'https://www.phoenixbooks.ru/books/book/O0105481/ekspluataciya-elektrotehnicheskih-sistem-ob-ektov-zhkh-uchebnoe-posobie.html',
    # Райкова: документ с «Содержанием»
    'dok_raikova':        'https://dokumen.pub/17982479a9de9ac89276074a8fd3f86c.html',
    # Пост. №170: структура разделов на КТНД
    'cntd_170':           'https://docs.cntd.ru/document/901869896/titles',
    #rusneb: Бычков Ч.1
    'rusneb_bychkov_ch1': 'https://rusneb.ru/catalog/000199_000009_008997994',
    # Лабиринт: Попов (иногда есть «Фрагмент»/содержание)
    'lab_popov2':         'https://www.labirint.ru/books/704357/',
    # Ермолаев: карточка читай-города
    'cg_ermolaev':        'https://www.chitai-gorod.ru/product/tehnicheskoe-obsluzhivanie-i-ekspluataciya-priborov-i-sistem-avtomatiki-1',
}

todo = sys.argv[1:] if len(sys.argv) > 1 else list(URLS)
for key in todo:
    out_path = os.path.join(OUT, f'{key}.json')
    if os.path.exists(out_path) and os.path.getsize(out_path) > 500:
        print(f'== {key}: есть, пропуск')
        continue
    args = json.dumps({'url': URLS[key]}, ensure_ascii=False)
    print(f'== {key}: {URLS[key][:80]}')
    try:
        r = subprocess.run(['z-ai', 'function', '-n', 'page_reader', '-a', args, '-o', out_path],
                           capture_output=True, text=True, timeout=120)
        tail = (r.stdout + r.stderr).strip().splitlines()
        print('   ', tail[-1] if tail else '(нет вывода)', '| size:', os.path.getsize(out_path) if os.path.exists(out_path) else 0)
    except Exception as e:
        print('   ОШИБКА:', e)
    time.sleep(1)
print('DONE')
