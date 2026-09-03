# -*- coding: utf-8 -*-
"""Пробник: извлечение строк МДК с часами из всех РП (docx — grid-парсинг, doc — antiword)."""
import re, subprocess, sys, zipfile, os

BASE = os.path.dirname(os.path.abspath(__file__)).replace('/scripts', '')

def docx_tables(path):
    """Все таблицы docx как списки строк (grid-парсинг с учётом gridSpan/vMerge)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    from xml.etree import ElementTree as ET
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    root = ET.fromstring(xml)
    tables = []
    for tbl in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'):
        rows = []
        for tr in tbl.findall('w:tr', ns):
            cells = []
            for tc in tr.findall('w:tc', ns):
                txt = ''.join(t.text or '' for t in tc.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
                tcpr = tc.find('w:tcPr', ns)
                span = 1
                if tcpr is not None:
                    gs = tcpr.find('w:gridSpan', ns)
                    if gs is not None:
                        span = int(gs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                vmerge = tcpr is not None and tcpr.find('w:vMerge', ns) is not None
                cont = vmerge and tcpr.find('w:vMerge', ns).get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') != 'restart'
                cells.append({'text': txt.strip(), 'span': span, 'vmerge_cont': cont})
            rows.append(cells)
        tables.append(rows)
    return tables

def flat(row, ncols):
    """Развёртка строки в ncols с обработкой vMerge-продолжений."""
    out = []
    for c in row:
        if c['vmerge_cont']:
            out.append('')
        else:
            out.append(c['text'])
            out.extend([''] * (c['span'] - 1))
    return (out + [''] * ncols)[:ncols]

def probe_docx(path, mdk_pat=r'МДК\s*\d+\.\d+'):
    print(f'\n########## {os.path.basename(path)} ##########')
    tables = docx_tables(path)
    for ti, rows in enumerate(tables):
        ncols = max(sum(c['span'] for c in r) for r in rows[:20]) if rows else 0
        hits = []
        for ri, r in enumerate(rows):
            joined = ' '.join(c['text'] for c in r if c['text'])
            if re.search(mdk_pat, joined) or ('Итого' in joined and 'ПМ' in joined):
                hits.append((ri, r))
        if not hits:
            continue
        # строка-заголовок: ближайшая строка выше с цифрами 1 2 3 4...
        print(f'--- Таблица #{ti+1} (строк {len(rows)}, ~{ncols} кол.) — {len(hits)} совпадений')
        for ri, r in hits[:12]:
            f = flat(r, ncols)
            cells = [c for c in f]
            print(f'  стр.{ri}: {cells}')
            if len(hits) > 12: break

def probe_doc(path):
    print(f'\n########## {os.path.basename(path)} (antiword) ##########')
    txt = subprocess.run(['antiword', '-w', '200', path], capture_output=True, text=True).stdout
    for line in txt.splitlines():
        if re.search(r'МДК\s*\d+\.\d+', line) or re.match(r'\s*\d+\.\d+\s', line) or 'Итого' in line:
            print(' ', line.strip()[:180])

if __name__ == '__main__':
    probe_docx(os.path.join(BASE, 'RP', 'РП_ПМ_05_M21.docx'))
    probe_docx(os.path.join(BASE, 'RP', '15.01.37_РП_ПМ.02_2025 (2).docx'))
    probe_docx(os.path.join(BASE, 'RP', '15.01.37_РП_ПМ.03_2025.docx'))
    probe_doc(os.path.join(BASE, 'RP', '08.02.09_ПМ.01_РП_2024.doc'))
    probe_doc(os.path.join(BASE, 'RP', 'РП ПМ 03 КИП.doc'))
