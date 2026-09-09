# -*- coding: utf-8 -*-
"""Доводка титула: пересогласование слова «час» после числа теории (34→часа, 20→часов)."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from sync_ktp_0201_0202_new_rp import replace_para_number, BASE
from docx import Document

JOBS = [
    ('KTP/КТП МДК 02.01 С-21.docx', '36', '34', 'часа'),
    ('KTP/КТП МДК 02.01 С-22.docx', '36', '34', 'часа'),
    ('KTP/КТП МДК 02.02 С-21.docx', '22', '20', 'часов'),
    ('KTP/КТП МДК 02.02 С-22.docx', '22', '20', 'часов'),
]
# файлы УЖЕ содержат новые числа (34/20) — слово меняем относительно них;
# replace_para_number ищет old-число: подсовываем текущее число как «old», новое = то же
for path, _old, cur, word in JOBS:
    p = f'{BASE}/{path}'
    doc = Document(p)
    replace_para_number(doc, 'теоретическое обучение', cur, cur, word=word)
    doc.save(p)
    print(f'✔ {path}: теория {cur} ({word})')
