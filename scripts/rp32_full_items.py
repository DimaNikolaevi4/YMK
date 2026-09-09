# -*- coding: utf-8 -*-
"""Полные тексты пунктов содержания РП 3.2 (МДК 02.01/02.02)."""
import sys
sys.path.insert(0, '/home/z/my-project/YMK/scripts')
from extract_ktp_hours import doc_text_and_tables, flat

PATH = '/home/z/my-project/YMK/RP/15.01.37_РП_ПМ.02_2025.docx'
texts, tables = doc_text_and_tables(PATH)
for ti, rows in enumerate(tables):
    j = ' '.join(c['text'] for rr in rows[:2] for c in rr)
    if 'Содержание учебного материала' in j or 'Наименование разделов' in j:
        ncols = max(sum(c['span'] for c in rr) for rr in rows)
        for ri, rr in enumerate(rows):
            f2 = (flat(rr) + [''] * ncols)[:ncols]
            tema = f2[0].strip()
            num = f2[1].strip()
            txt = f2[2].strip()
            k3 = f2[4].strip()
            k4 = f2[5].strip() if len(f2) > 5 else ''
            if tema.startswith('МДК') or tema.startswith('Тема') or tema.startswith('Всего') or 'аттестация' in tema.lower():
                print(f'\n### r{ri}: [{tema[:90]}] col3={k3} col4={k4}')
            elif num and txt:
                print(f'  r{ri} [{num}] col3={k3} col4={k4}: {txt}')
            elif num and not txt:
                print(f'  r{ri} [{num}] col3={k3} col4={k4}: <пусто>')
