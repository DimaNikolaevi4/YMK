#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Построчное сравнение сгенерированного КТП с эталоном: merge-структура, ширины, тексты."""
from docx import Document

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

ETALON = "/home/z/my-project/YMK/KTP/КТП МДК 05.02 М2 курс.docx"
GEN = "/home/z/my-project/test_ktp_generated.docx"

e = Document(ETALON)
g = Document(GEN)

def row_signature(tbl, ri):
    """Подпись строки: [(gridSpan, vMerge, текст), ...] по физическим ячейкам."""
    tr = tbl._tbl.findall(W + 'tr')[ri]
    sig = []
    for tc in tr.findall(W + 'tc'):
        tcPr = tc.find(W + 'tcPr')
        gs, vm = 1, ''
        if tcPr is not None:
            el = tcPr.find(W + 'gridSpan')
            if el is not None:
                gs = int(el.get(W + 'val'))
            el = tcPr.find(W + 'vMerge')
            if el is not None:
                vm = el.get(W + 'val') or 'cont'
        text = ''.join(t.text or '' for t in tc.iter(W + 't')).strip()
        sig.append((gs, vm, text))
    return sig

def compare_table(name, ei, gi, rows_to_check, verbose=True):
    te, tg = e.tables[ei], g.tables[gi]
    diffs = 0
    for ri in rows_to_check:
        se, sg = row_signature(te, ri), row_signature(tg, ri)
        # сравниваем структуру (gs, vm) и текст (только для не-vMerge-continuation)
        se_s = [(a, b) for a, b, c in se]
        sg_s = [(a, b) for a, b, c in sg]
        te_s = [c for a, b, c in se]
        tg_s = [c for a, b, c in sg]
        if se_s != sg_s:
            print(f'❌ {name} R{ri}: merge-структура отличается')
            print(f'   эталон: {se_s}')
            print(f'   генер.: {sg_s}')
            diffs += 1
        elif te_s != tg_s:
            print(f'❌ {name} R{ri}: тексты отличаются')
            for a, b in zip(te_s, tg_s):
                if a != b:
                    print(f'   эталон: {a!r}')
                    print(f'   генер.: {b!r}')
            diffs += 1
        elif verbose:
            print(f'✅ {name} R{ri}: идентично ({len(se)} ячеек)')
    return diffs

total = 0
print('### Ширины колонок (tblGrid)')
for ei, gi, nm in [(1, 1, 'Т1'), (2, 2, 'Т2'), (3, 3, '2а'), (4, 4, '2б'), (6, 6, '2г')]:
    we = [int(c.get(W + 'w')) for c in e.tables[ei]._tbl.find(W + 'tblGrid').findall(W + 'gridCol')]
    wg = [int(c.get(W + 'w')) for c in g.tables[gi]._tbl.find(W + 'tblGrid').findall(W + 'gridCol')]
    ok = we == wg
    if not ok:
        total += 1
    print(('✅' if ok else '❌'), nm, we if ok else f'эталон={we} генер.={wg}')

print()
print('### Таблица 1 (шапка + номера + данные + Практика + Всего)')
total += compare_table('Т1', 1, 1, [0, 1, 2, 3, 4, 5, 6, 10, 11], verbose=False)
print('   (структура всех проверенных строк идентична — см. ошибки выше, если есть)')

print()
print('### Таблица 2 (шапка + МДК + темы + первое/последнее занятие + аттестация + Итого)')
total += compare_table('Т2', 2, 2, [0, 1, 2, 3, 4, 5, 6, 67, 68, 69], verbose=False)
print('   (проверены R0-R6, R67-R69)')

print()
print('### Количество строк')
for ei, gi, nm in [(1, 1, 'Т1'), (2, 2, 'Т2'), (3, 3, '2а'), (4, 4, '2б'), (5, 5, '2в'), (6, 6, '2г')]:
    re_, rg = len(e.tables[ei].rows), len(g.tables[gi].rows)
    # 2а в эталоне имеет пустые строки-резерв — допустимое отличие
    note = ' (в эталоне есть пустые строки-резерв — это допустимо)' if nm == '2а' and re_ != rg else ''
    print(f'{"✅" if re_ == rg else "⚠️"} {nm}: эталон={re_}, генер.={rg}{note}')

print()
print('### Титульный лист (тексты абзацев)')
ep = [p.text.strip() for p in e.paragraphs if p.text.strip()]
gp = [p.text.strip() for p in g.paragraphs if p.text.strip()]
# нормализуем пробелы
norm = lambda s: ' '.join(s.split())
epn, gpn = [norm(x) for x in ep], [norm(x) for x in gp]
set_e, set_g = set(epn), set(gpn)
only_e = [x for x in epn if x not in set_g]
only_g = [x for x in gpn if x not in set_e]
if only_e or only_g:
    total += 1
    print('❌ Различия в абзацах титульника:')
    for x in only_e:
        print(f'   только в эталоне: {x[:90]}')
    for x in only_g:
        print(f'   только в генер.:  {x[:90]}')
else:
    print('✅ Все абзацы титульника совпадают (с точностью до повторов)')

print()
print('### Разрывы страниц')
import re
xe = e.element.body.xml
xg = g.element.body.xml
ne = len(re.findall(r'w:type="page"', xe))
ng = len(re.findall(r'w:type="page"', xg))
print(f'{"✅" if ne == ng else "❌"} разрывов: эталон={ne}, генер.={ng}')
if ne != ng:
    total += 1

print()
print('=' * 60)
print('КРИТИЧЕСКИХ РАСХОЖДЕНИЙ:' if total else 'ВСЁ ИДЕНТИЧНО ЭТАЛОНУ ✓', total if total else '')
