#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ищи-2, шаг 4: волна 2 — целевые страницы с оглавлениями/полными текстами."""
import json, subprocess, os, time, sys

BASE = '/home/z/my-project/YMK'
OUT = os.path.join(BASE, 'docs/search/ishi2/pages')
os.makedirs(OUT, exist_ok=True)

URLS = {
    # Попов: aldebaran (обещает «скачать книгу бесплатно»)
    'alb_popov_spo':   'https://aldebaran.one/author/n-m-popov/book-kniga-izmereniya-v-elektricheskih-setyah-0-4-10-kv-2719449-1051872',
    # Попов: карточка УРСС (2-е изд., 228 с.) — у urss часто есть «Содержание»
    'urss_popov':      'https://urss.ru/cgi-bin/db.pl?lang=Ru&blang=ru&page=Book&id=291711',
    # Бычков «Эксплуатация и ремонт»: библиотечные карточки
    'mpei_bychkov':    'https://opac.mpei.ru/OpacUnicode/index.php?url=/notices/index/IdNotice:302405/Source:default',
    'arbat_bychkov_er':'https://mdk-arbat.ru/book/3374740',
    # Сидорова: books.ru / читай-город
    'cg_sidorova':     'https://www.chitai-gorod.ru/product/sborka-montazh-regulirovka-i-remont-uzlov-i-meh-oborudovaniya-uchebnik-2907461',
    # Ярочкина «Проверка и наладка»: читай-город
    'cg_yarochkina':   'https://www.chitai-gorod.ru/product/proverka-i-naladka-elektrooborudovaniya-uchebnik-2871104',
    # Ермолаев: bookvoed
    'bv_ermolaev':     'https://www.bookvoed.ru/product/tekhnicheskoe-obsluzhivanie-i-ekspluatatsiya-priborov-i-sistem-avtomatiki-7796228',
    # Григорьева: карточка ЦНСХБ
    'dm_grigorieva':   'https://lib.dm-centre.ru/lib/document/gpntb/ESVODT/d62e7a56f092bb38d678b0e1b042c4cc',
    # Пост. 170: консультант+ (структура разделов)
    'cons_170':        'https://www.consultant.ru/document/cons_doc_LAW_44772/',
    # Полуянович «Монтаж, наладка...» (ОИ 4 05.02): карточка Лани
    'lan_poluyanovich_cat': 'https://lanbook.com/catalog/energetika/montazh-naladka-ekspluatatsiya-i-remont-sistem-elektrosnabzheniya-promyshlennyh-predpriyat',
    # Райкова: карточка Юрайт-документа ДВФУ
    'dvfu_raikova':    'https://library.dvfu.ru/lib/document/EBSUrait/9BCDC527-3189-4C25-B585-DB4D1EA7056B',
    # Полуянович ЖКХ (Феникс 2020, 158 с.): PDF ТТИ ЮФУ с упоминанием — там может быть библиография+структура
    'tti_poluyanovich':'https://ntb.tti.sfedu.ru/wp-content/uploads/2023/07/U_%E2%84%96-2.pdf',
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
