# -*- coding: utf-8 -*-
"""ищи-3: извлечение оглавлений из скачанных полных PDF."""
import subprocess, re, os

F = '/home/z/my-project/YMK/docs/search/ishi3/fetch'
BOOKS = {
    'poluyanovich_lan':  (1, 300),   # Полуянович Таганрог 2007, 280 с. (PDF 283 с.)
    'akimova_full':      (1, 310),   # Акимова Академия (PDF 303 с.)
    'romanovich':        (1, 320),   # Романович (PDF 316 с.)
    'shishmarev_avtomatika_trial': (1, 40),  # Шишмарёв Автоматика trial Литрес
}

for name, (a, b) in BOOKS.items():
    pdf = f'{F}/{name}.pdf'
    txt_path = f'{F}/{name}_fulltext.txt'
    if not os.path.exists(txt_path) or os.path.getsize(txt_path) < 1000:
        r = subprocess.run(['pdftotext', '-f', str(a), '-l', str(b), pdf, txt_path],
                           capture_output=True, text=True)
        if r.returncode != 0:
            print(name, 'ERR', r.stderr[:200]); continue
    t = open(txt_path, encoding='utf-8', errors='ignore').read()
    print('=' * 100)
    print(f'{name}: {len(t)} символов')
    # ищем «Содержание»/«Оглавление»
    for m in re.finditer(r'(СОДЕРЖАНИЕ|Содержание|ОГЛАВЛЕНИЕ|Оглавление)', t):
        seg = t[m.start():m.start() + 4200]
        # не берём, если это оглавление в конце большого текста без номеров страниц
        print(f'--- вхождение на позиции {m.start()} ---')
        print(seg[:4200])
        print()
        break
