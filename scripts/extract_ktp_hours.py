# -*- coding: utf-8 -*-
"""Извлечение часов из КТП (титул + Таблица 2) для сводки по часам."""
import re, os, sys, zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

def doc_text_and_tables(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(xml)
    body = root.find('w:body', NS)
    title_text, tables = [], []
    for el in body:
        tag = el.tag.split('}')[1]
        if tag == 'p':
            t = ''.join(n.text or '' for n in el.iter(f'{{{W}}}t')).strip()
            if t:
                title_text.append(t)
        elif tag == 'tbl':
            rows = []
            for tr in el.findall('w:tr', NS):
                cells = []
                for tc in tr.findall('w:tc', NS):
                    txt = ''.join(n.text or '' for n in tc.iter(f'{{{W}}}t')).strip()
                    span = 1
                    tcpr = tc.find('w:tcPr', NS)
                    if tcpr is not None:
                        gs = tcpr.find('w:gridSpan', NS)
                        if gs is not None:
                            span = int(gs.get(f'{{{W}}}val'))
                    cells.append({'text': txt, 'span': span})
                rows.append(cells)
            tables.append(rows)
    return title_text, tables

def flat(row):
    return [c['text'] for c in row]

def is_num(s):
    return bool(re.fullmatch(r'\d{1,3}(,\d)?', s.strip()))

def analyze(path):
    t, tables = doc_text_and_tables(path)
    res = {'file': os.path.basename(path), 'dir': os.path.basename(os.path.dirname(path))}
    # --- титул: год и «в т.ч. в форме практической подготовки»
    title_block = ' | '.join(t[:45])
    m = re.search(r'20\d\d\s*[-–—/]\s*20\d\d', title_block)
    res['уч_год_в_титуле'] = m.group(0) if m else ''
    def num_after(pattern):
        mm = re.search(pattern + r'[\s_]*(\d+)', title_block)
        return int(mm.group(1)) if mm else None
    res['титул_объём_прогр'] = num_after(r'Объем образовательной программы')
    res['титул_практ_подг'] = num_after(r'практической подготовки')
    res['титул_нагрузка_препод'] = num_after(r'нагрузка во взаимодействии с преподавателем')
    res['титул_теория'] = num_after(r'теоретическое обучение')
    res['титул_практ_зан'] = num_after(r'практические занятия')
    m_att = re.search(r'аттестация в форме\s*(.*?)\(час', title_block)
    res['титул_аттестация'] = m_att.group(1).strip(' _') if m_att else None
    # --- Таблица 2: тематический план (МДК-строка, дифзачёт, Итого, суммы кол.3/4)
    t2 = None
    for rows in tables:
        j = ' '.join(c['text'] for r in rows[:4] for c in r)
        if re.search(r'(Тематический план|Содержание учебного материала|Наименование разделов)', j, re.I):
            t2 = rows
            break
    if t2:
        ncols = max(sum(c['span'] for c in r) for r in t2)
        sum3 = sum4 = 0
        for r in t2:
            f = (flat(r) + [''] * ncols)[:ncols]
            rowtxt = ' '.join(f)
            if re.search(r'МДК\s*\d+\.\d+', rowtxt) and not re.search(r'№ занятия|наименование', rowtxt, re.I):
                nums = [x for x in f[2:4] if is_num(x)]
                if len(nums) >= 2:
                    res['т2_строка_МДК'] = [int(float(x)) for x in nums[:2]]
            if (f[0] or '').strip().lower().startswith('итог'):
                nums = [x for x in f[1:4] if is_num(x)]
                if len(nums) >= 2:
                    res['т2_Итого'] = [int(float(x)) for x in nums[:2]]
            if re.fullmatch(r'\d{1,2}', f[0] or '') or re.match(r'^Тема\b', f[1] or ''):
                nums = [x for x in f[2:4] if is_num(x)]
                if len(nums) >= 2:
                    sum3 += int(float(nums[0])); sum4 += int(float(nums[1]))
            if (f[1] or '').strip().lower() == 'дифференцированный зачет' or (f[0] or '').strip().lower() == 'дифференцированный зачет':
                nums = [x for x in f[2:4] if is_num(x)]
                if nums:
                    res['т2_дифзачет_часы'] = int(float(nums[0]))
        res['т2_сумма_кол3'] = sum3
        res['т2_сумма_кол4'] = sum4
    # --- Таблица 1: по семестрам (строки МДК + Всего)
    t1 = None
    for rows in tables:
        j = ' '.join(c['text'] for r in rows[:2] for c in r)
        if re.search(r'Распределение|Объем образовательной прогр', j, re.I) or ('Семестр' in j and 'Курс' in j):
            t1 = rows
            break
    if t1:
        ncols = max(sum(c['span'] for c in r) for r in t1)
        sem_rows, total_row = [], None
        for r in t1:
            f = (flat(r) + [''] * ncols)[:ncols]
            if re.match(r'^\s*МДК\s*\d+\.\d+\s*$', f[0] or ''):
                nums = [x for x in f if is_num(x)]
                if len(nums) >= 4:
                    sem_rows.append({'курс': f[1], 'семестр': f[2], 'числа': [int(float(x)) for x in nums[:8]]})
            if (f[0] or '').strip().lower() == 'всего':
                nums = [x for x in f[1:] if is_num(x)]
                total_row = [int(float(x)) for x in nums[:6]]
        res['т1_семестры'] = sem_rows
        res['т1_Всего'] = total_row
    return res

if __name__ == '__main__':
    base = '/home/z/my-project/YMK'
    files = [
        'KTP/КТП МДК 05.02 М2 курс.docx',
        'KTP/КТП МДК 01.02 М-21.docx',
        'KTP/КТП МДК 02.01 С-21.docx',
        'KTP/КТП МДК 02.01 С-22.docx',
        'KTP/КТП МДК 02.02 С-21.docx',
        'KTP/КТП МДК 02.02 С-22.docx',
        'КТП МДК 02.01 КИП3 курс.docx',
        'КТП МДК 02.01 ОТП 4 курс.docx',
        'КТП МДК 03.01 ЭМ 3 курс.docx',
    ]
    for f in files:
        p = os.path.join(base, f)
        if not os.path.exists(p):
            print(f'!!! НЕ НАЙДЕН: {f}')
            continue
        r = analyze(p)
        print('===', r['dir'] + '/' + r['file'])
        for k, v in r.items():
            if k not in ('file', 'dir'):
                print(f'   {k} = {v}')
