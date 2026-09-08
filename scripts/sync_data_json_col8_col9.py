# -*- coding: utf-8 -*-
"""Синхронизация equipment (кол.8) и task (кол.9) в data_*.json с фактическим
состоянием docx после fix_col8_col9_split.py (+ ищи-2/ищи-3).

Без этого перегенерация docx из JSON откатила бы кол.8/кол.9 к «Конспект
лекций»/«ОИ1». Пары JSON ← docx:
  data_05_02_example  ← КТП МДК 05.02 М2 курс
  data_01_02_m21      ← КТП МДК 01.02 М-21
  data_02_01_s21      ← КТП МДК 02.01 С-21 (С-22 идентичен)
  data_02_02_s21      ← КТП МДК 02.02 С-21 (С-22 идентичен)
  data_03_02_s21      ← КТП МДК 03.02 С-21
  data_03_02_s22      ← КТП МДК 03.02 С-22
"""
import json
import sys
import zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

PAIRS = [
    ('/home/z/my-project/YMK/scripts/data_05_02_example.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx'),
    ('/home/z/my-project/YMK/scripts/data_01_02_m21.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 01.02 М-21.docx'),
    ('/home/z/my-project/YMK/scripts/data_02_01_s21.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 02.01 С-21.docx'),
    ('/home/z/my-project/YMK/scripts/data_02_02_s21.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 02.02 С-21.docx'),
    ('/home/z/my-project/YMK/scripts/data_03_02_s21.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 03.02 С-21.docx'),
    ('/home/z/my-project/YMK/scripts/data_03_02_s22.json',
     '/home/z/my-project/YMK/KTP/КТП МДК 03.02 С-22.docx'),
]


def cell_text(tc):
    return ''.join(t.text or '' for t in tc.iter(f'{{{W}}}t')).strip()


def t2_rows(docx):
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read('word/document.xml').decode('utf-8'))
    for tbl in root.find('w:body', NS).findall('w:tbl', NS):
        if '№ занятия' in ''.join(t.text or '' for t in tbl.iter(f'{{{W}}}t'))[:200]:
            out = {}
            for tr in tbl.findall('w:tr', NS):
                tcs = tr.findall('w:tc', NS)
                if len(tcs) < 9:
                    continue
                c0 = cell_text(tcs[0])
                if not c0.isdigit():
                    continue
                c8 = cell_text(tcs[8])
                if c8.isdigit() and c0 == '1':
                    continue  # номерная строка шапки
                out[int(c0)] = (cell_text(tcs[7]), c8)
            return out
    raise RuntimeError(f'{docx}: Т2 не найдена')


for jpath, dpath in PAIRS:
    rows = t2_rows(dpath)
    j = json.load(open(jpath, encoding='utf-8'))
    n_upd, mism = 0, []
    for topic in j['topics']:
        for les in topic.get('lessons', []):
            num = les['num']
            if num not in rows:
                mism.append(f'№{num} нет в docx')
                continue
            new_eq, new_task = rows[num]
            if les.get('equipment') != new_eq or les.get('task') != new_task:
                les['equipment'], les['task'] = new_eq, new_task
                n_upd += 1
    assert not mism, f'{jpath}: {mism}'
    json.dump(j, open(jpath, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'{jpath.split("/")[-1]}: обновлено {n_upd} занятий (из {len(rows)} строк docx)')
print('\nГотово: data-JSON синхронизированы с docx.')
