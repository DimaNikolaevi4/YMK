# -*- coding: utf-8 -*-
"""
Синхронизация КТП МДК 02.01 и 02.02 (С-21/С-22) с НОВЫМ РП 15.01.37_РП_ПМ.02_2025.docx
(владелец заменил файл 09.09.2026, команда: «изменил РП — измени КТП»).

Новый РП: МДК 02.01 = 48 (34 теор + 12 ПЗ + 2 КДЗ), МДК 02.02 = 48 (20 + 26 + 2).
Правки Т2 строго построчно из РП 3.2:
  02.01: тексты 1.1.5/1.1.7/1.2.2/1.2.7, вставка 1.1.6, объединения (1.2.2=11+12, 1.2.7=17+18)
  02.02: тексты 1.3.3/1.3.4/1.4.4/1.4.6, вставки 1.4.1/1.4.5, объединения (1.3.3=3+4, 1.3.4=5+6)
Плюс: перенумерация, часы тем/МДК/Итого, Т1, титул, пересборка data JSON.
"""
import copy
import json
import re
import sys
from docx import Document

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NSW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': NSW}
BASE = '/home/z/my-project/YMK'
RP_PATH = f'{BASE}/RP/15.01.37_РП_ПМ.02_2025.docx'


# ---------------- RP: точные тексты пунктов 3.2 ----------------
def load_rp_texts():
    from xml.etree import ElementTree as ET
    with zipfile_open(RP_PATH) as f:
        xml = f.read()
    root = ET.fromstring(xml)
    body = root.find('w:body', NS)
    texts = {}
    tbl_i = 0
    for el in body:
        if el.tag.split('}')[1] != 'tbl':
            continue
        if tbl_i == 8:  # таблица 3.2
            rows = el.findall('w:tr', NS)
            # (индекс строки, ключ)
            wanted = {
                9: '1.1.5', 10: '1.1.6', 11: '1.1.7',
                17: '1.2.2', 22: '1.2.7',
                37: '1.3.3', 38: '1.3.4',
                47: '1.4.1', 50: '1.4.4', 51: '1.4.5', 52: '1.4.6',
            }
            for ri, tr in enumerate(rows):
                if ri not in wanted:
                    continue
                tcs = tr.findall('w:tc', NS)
                # текст = самая длинная ячейка (содержание), часы = ячейка-число
                cand = [(''.join(n.text or '' for n in tc.iter(f'{{{NSW}}}t'))).strip()
                        for tc in tcs]
                longest = max(cand, key=len)
                digits = [c for c in cand if re.fullmatch(r'\d{1,2}', c)]
                assert '2' in digits, (wanted[ri], cand)
                texts[wanted[ri]] = longest
            break
        tbl_i += 1
    assert len(texts) == 11, texts.keys()
    return texts


from contextlib import contextmanager


@contextmanager
def zipfile_open(path):
    import zipfile
    with zipfile.ZipFile(path) as z:
        yield z.open('word/document.xml')


# ---------------- низкоуровневые утилиты ----------------
def set_tc_text(tc, new_text):
    """Первый w:t ячейки = new_text, остальные очищаются (форматирование run'ов сохраняется).
    Пустая ячейка без w:t — создаётся новый run в первом абзаце."""
    from lxml import etree
    ts = tc.findall('.//' + W + 't')
    if not ts:
        p = tc.find(W + 'p')
        assert p is not None, 'в ячейке нет абзаца'
        r = etree.SubElement(p, W + 'r')
        ts = [etree.SubElement(r, W + 't')]
    ts[0].text = new_text
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for t in ts[1:]:
        t.text = ''


def tc_text(tc):
    return ''.join(n.text or '' for n in tc.iter(W + 't')).strip()


def tr_cells(tr):
    return tr.findall(W + 'tc')


def row_c0(tr):
    return tc_text(tr_cells(tr)[0])


def row_c1(tr):
    return tc_text(tr_cells(tr)[1])


def replace_para_number(doc, marker, old, new, word=None):
    """Заменяет число old→new в абзаце титуле (в одном run; форматирование сохраняется).
    word — если задано, первое «часов»/«часа» после числа пересогласовывается (7е)."""
    for p in doc.paragraphs:
        if marker in p.text:
            for idx, r in enumerate(p.runs):
                if old in r.text:
                    r.text = r.text.replace(old, new, 1)
                    if word:
                        for r2 in p.runs[idx:]:
                            if 'часов' in r2.text:
                                r2.text = r2.text.replace('часов', word, 1)
                                break
                            if 'часа' in r2.text:
                                r2.text = r2.text.replace('часа', word, 1)
                                break
                    return
            raise RuntimeError(f'{marker}: {old} не найдено в одном run; runs={[r.text for r in p.runs]}')
    raise RuntimeError(f'абзац-маркер {marker!r} не найден')


# ---------------- работа с Т2 ----------------
def data_rows(t2):
    """Строки занятий (c0 — число, c1 — ТЕКСТ; шапка r3 «1|2|3…» исключается)."""
    return [tr for tr in t2._tbl.findall(W + 'tr')
            if row_c0(tr).isdigit() and not row_c1(tr).strip().isdigit()]


def find_row(t2, prefix):
    for tr in data_rows(t2):
        if row_c1(tr).startswith(prefix):
            return tr
    raise RuntimeError(f'строка {prefix!r} не найдена')


def fill_lesson_row(tr, name, eq, task, control='Текущий контроль'):
    cs = tr_cells(tr)
    assert len(cs) == 11, len(cs)
    vals = ['0', name, '2', '', 'Урок', 'ОК 1 - ОК 7, ОК 9', 'ПК 2.1 - ПК 2.2',
            eq, task, control, '']
    for tc, v in zip(cs, vals):
        set_tc_text(tc, v)


def update_theme_hours(t2, prefix, hours):
    for tr in t2._tbl.findall(W + 'tr'):
        c0, c1 = row_c0(tr), row_c1(tr)
        if not c0 and c1.startswith(prefix):
            cs = tr_cells(tr)
            set_tc_text(cs[2], str(hours))
            return
    raise RuntimeError(f'тема {prefix!r} не найдена')


def update_mdk_total_row(t2, mdk_prefix, total):
    for tr in t2._tbl.findall(W + 'tr'):
        c0, c1 = row_c0(tr), row_c1(tr)
        if c1.startswith(mdk_prefix):
            set_tc_text(tr_cells(tr)[2], str(total))
            return
    raise RuntimeError(f'строка {mdk_prefix} не найдена')


def update_itogo(t2, total):
    for tr in t2._tbl.findall(W + 'tr'):
        if row_c1(tr).startswith('Итого'):
            set_tc_text(tr_cells(tr)[2], str(total))
            return
    raise RuntimeError('Итого не найдено')


def renumber(t2):
    n = 0
    for tr in data_rows(t2):
        n += 1
        set_tc_text(tr_cells(tr)[0], str(n))
    return n


def set_t1_numbers(t1, total, theory, pract, theory_total):
    sem = tot = None
    for tr in t1._tbl.findall(W + 'tr'):
        c0 = row_c0(tr)
        if c0.startswith('МДК'):
            sem = tr
        elif c0 == 'Всего':
            tot = tr
    assert sem is not None and tot is not None
    for tr, th in ((sem, theory), (tot, theory_total)):
        cs = tr_cells(tr)
        set_tc_text(cs[3], str(total))
        set_tc_text(cs[4], str(total))
        set_tc_text(cs[5], str(th))
        set_tc_text(cs[7], str(pract))


# ---------------- основной сценарий по файлу ----------------
def process(path, mdk, plan, rp):
    """plan: dict с параметрами правок для конкретного МДК."""
    doc = Document(path)
    t1, t2 = doc.tables[1], doc.tables[2]
    log = []

    # 1. Текстовые замены (объединённые/обновлённые пункты РП)
    for prefix, key in plan['replace']:
        tr = find_row(t2, prefix)
        set_tc_text(tr_cells(tr)[1], rp[key])
        log.append(f'текст строки "{prefix[:35]}..." ← РП {key}')

    # 2. Вставки
    for anchor_prefix, before, key, eq, task in plan['insert']:
        anchor = find_row(t2, anchor_prefix)
        tpl = find_row(t2, plan['template'])
        new_tr = copy.deepcopy(tpl)
        fill_lesson_row(new_tr, rp[key], eq, task)
        if before:
            anchor.addprevious(new_tr)
        else:
            anchor.addnext(new_tr)
        log.append(f'вставка РП {key} ({eq} {task}) {"перед" if before else "после"} "{anchor_prefix[:30]}..."')

    # 3. Удаления
    for prefix in plan['delete']:
        tr = find_row(t2, prefix)
        tr.getparent().remove(tr)
        log.append(f'удалена строка "{prefix[:40]}..."')

    # 4. Часы: темы, МДК, Итого
    for prefix, hours, pract in plan['themes']:
        update_theme_hours(t2, prefix, hours)
        log.append(f'{prefix}: часы → {hours}')
    update_mdk_total_row(t2, 'МДК ' + mdk, plan['total'])
    update_itogo(t2, plan['total'])
    log.append(f'МДК/Итого → {plan["total"]}')

    # 5. Перенумерация
    n = renumber(t2)
    log.append(f'перенумеровано занятий: {n}')

    # 6. Т1
    set_t1_numbers(t1, plan['total'], plan['theory'], plan['pract'], plan['theory_total'])
    log.append(f"Т1: {plan['total']}/{plan['total']}/{plan['theory']}/-/{plan['pract']}; "
               f"Всего теория {plan['theory_total']}")

    # 7. Титул
    replace_para_number(doc, 'Объем образовательной программы', '50', '48')
    replace_para_number(doc, 'Учебная нагрузка во взаимодействии', '50', '48')
    replace_para_number(doc, 'теоретическое обучение', plan['old_theory'], str(plan['theory']),
                        word=plan['theory_word'])
    log.append('титул: 50→48 (×2), теория ' + plan['old_theory'] + '→' + str(plan['theory']))

    doc.save(path)
    return log, n


PLAN_01 = {
    'total': 48, 'theory': 34, 'pract': 12, 'theory_total': 36, 'old_theory': '36',
    'theory_word': 'часа',  # 34 → часа (7е)
    'template': 'Метрологическое обеспечение испытаний',
    'replace': [
        ('Основные понятия о гибких автоматизированных производствах', '1.1.5'),
        ('Структурная и принципиальная электрическая схема и принципы работы', '1.1.7'),
        ('Поузловая приемка и испытания конструктивных', '1.2.2'),
        ('Наладка оборудования измерения и контроля температуры', '1.2.7'),
    ],
    'insert': [
        # якорь — УЖЕ обновлённый текст 1.1.5 (начинается с «Классификация…»)
        ('Классификация автоматических станочных систем', False,
         '1.1.6', 'ДИ3', 'с. 67-84'),
    ],
    'delete': [
        'Индивидуальные испытания приборов для измерения и контроля уровня',
        'Пробные пуски оборудования автоматического пожаротушения',
    ],
    'themes': [('Тема 1.1', 20, 0), ('Тема 1.2', 26, 12)],
}

PLAN_02 = {
    'total': 48, 'theory': 20, 'pract': 26, 'theory_total': 22, 'old_theory': '22',
    'theory_word': 'часов',  # 20 → часов (7е)
    'template': 'Системы автоматического регулирования. Принципы регулирования',
    'replace': [
        ('Логарифмические частотные характеристики', '1.3.3'),
        ('Техническое обеспечение систем автоматического регулирования', '1.3.4'),
        ('Схемы визуального моделирования в MicrosoftVisio. Назначение системы КОМПАС', '1.4.4'),
        ('Построений сопряжений и нанесение размеров', '1.4.6'),
    ],
    'insert': [
        ('Форматирование фигуры в MS Visio. Текстовые элементы', True,
         '1.4.1', 'ДИ5', 'с. 3-12'),
        ('Схемы визуального моделирования в MicrosoftVisio. Назначение системы КОМПАС', False,
         '1.4.5', 'ДИ2', 'с. 85-102'),
    ],
    'delete': [
        'Виды систем управления. Понятие об адаптивном',
        'Промышленные микропроцессорные контроллеры',
        'Создание 3D-модели с использованием вспомогательных',
    ],
    'themes': [('Тема 1.3', 20, 12), ('Тема 1.4', 26, 14)],
}


def main():
    rp = load_rp_texts()
    jobs = [
        ('KTP/КТП МДК 02.01 С-21.docx', '02.01', PLAN_01),
        ('KTP/КТП МДК 02.01 С-22.docx', '02.01', PLAN_01),
        ('KTP/КТП МДК 02.02 С-21.docx', '02.02', PLAN_02),
        ('KTP/КТП МДК 02.02 С-22.docx', '02.02', PLAN_02),
    ]
    for path, mdk, plan in jobs:
        p = f'{BASE}/{path}'
        print('=' * 90)
        print(path)
        log, n = process(p, mdk, plan, rp)
        for line in log:
            print('  •', line)
        print(f'  ✔ сохранено, занятий: {n}')

    # 8. Пересборка data JSON из docx (источник истины — docx после правок)
    rebuild_data_json('KTP/КТП МДК 02.01 С-21.docx', 'scripts/data_02_01_s21.json',
                      total=48, theory=34, pract=12, rp_comment=(
                          'Источник — RP/15.01.37_РП_ПМ.02_2025.docx (новая версия владельца от 09.09.2026, '
                          'вместо «(2)»), таблица 3.2 «МДК 02.01», построчно: 48 ч = Тема 1.1 (20 = 10 уроков×2) '
                          '+ Тема 1.2 (26 = 7 уроков×2 = 14 + 7 ПЗ = 12) + КДЗ 2. Практ. подготовка 12 = кол.4 '
                          'строки МДК. Титул/Т1/Т2 = 48 (46 занятий + 2 КДЗ). Правки: тексты 1.1.5/1.1.7/1.2.2/1.2.7 '
                          'заменены на тексты РП, вставлен урок 1.1.6 (ДИ3 с. 67-84 🟡), объединены 1.2.2 (старые '
                          '№11+№12) и 1.2.7 (старые №17+№18), удалены старые №12/№18. ВНИМАНИЕ: новый РП даёт 46 '
                          'уч.занятий (34 теор.), УП С-21 формально 48 уч.зан (36 теор.) + 2 КДЗ = 50 — '
                          'расхождение 2 ч теории вынесено владельцу.'))
    rebuild_data_json('KTP/КТП МДК 02.02 С-21.docx', 'scripts/data_02_02_s21.json',
                      total=48, theory=20, pract=26, rp_comment=(
                          'Источник — RP/15.01.37_РП_ПМ.02_2025.docx (новая версия владельца от 09.09.2026), '
                          'таблица 3.2 «МДК 02.02», построчно: 48 ч = Тема 1.3 (20 = 4 урока×2 = 8 + 6 ПЗ = 12) '
                          '+ Тема 1.4 (26 = 6 уроков×2 = 12 + 6 ПЗ = 14) + КДЗ 2. Практ. подготовка 26 = кол.4. '
                          'Правки: тексты 1.3.3 (объед. старых №3+№4) и 1.3.4 (объед. №5+№6), вставлены уроки '
                          '1.4.1 (ДИ5 с. 3-12 🟡) и 1.4.5 (ДИ2 с. 85-102 🟡), 1.4.4 дополнен «Создание файлов», '
                          '1.4.6 = объед. старых №16+№17. УП С-21 формально 48 уч.зан (22 теор.) + 2 КДЗ = 50 — '
                          'расхождение 2 ч теории вынесено владельцу.'))


def rebuild_data_json(doc_path, json_path, total, theory, pract, rp_comment):
    doc = Document(f'{BASE}/{doc_path}')
    t2 = doc.tables[2]
    themes, cur = [], None
    for tr in t2._tbl.findall(W + 'tr'):
        cs = tr_cells(tr)
        if len(cs) < 4:
            continue
        c0, c1 = tc_text(cs[0]), tc_text(cs[1])
        if c1.startswith('МДК'):
            continue
        if not c0 and c1.startswith('Тема'):
            nums = [tc_text(c) for c in cs[2:5]]
            hrs = [x for x in nums if x.strip().isdigit()]
            cur = {'name': c1,
                   'total_hours': int(hrs[0]) if hrs else 0,
                   'practice_hours': int(hrs[1]) if len(hrs) > 1 else 0,
                   'lessons': []}
            themes.append(cur)
            continue
        if c1.startswith('Итого'):
            break
        if c0.isdigit() and not c1.strip().isdigit() and cur is not None \
                and 'зачет' not in c1.lower() and 'зачёт' not in c1.lower():
            assert len(cs) == 11, (c0, c1[:40], len(cs))
            cur['lessons'].append({
                'num': int(c0), 'name': c1,
                'hours': int(tc_text(cs[2]) or 0),
                'practice_prep': int(tc_text(cs[3]) or 0),
                'type': tc_text(cs[4]), 'ok': tc_text(cs[5]), 'pk': tc_text(cs[6]),
                'equipment': tc_text(cs[7]), 'task': tc_text(cs[8]), 'control': tc_text(cs[9]),
            })

    with open(f'{BASE}/{json_path}', encoding='utf-8') as f:
        data = json.load(f)
    data['total_hours'] = total
    data['theory_hours'] = theory
    data['practice_hours'] = pract
    data['table1_semesters'][0].update({
        'total': total, 'with_teacher': total, 'theory': theory, 'practice': pract,
    })
    data['topics'] = themes
    data['rp_sync'] = {'_comment': rp_comment}
    with open(f'{BASE}/{json_path}', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    s = sum(l['hours'] for t in themes for l in t['lessons'])
    sp = sum(l['practice_prep'] for t in themes for l in t['lessons'])
    print(f'{json_path}: тем {len(themes)}, занятий {sum(len(t["lessons"]) for t in themes)}, '
          f'Σ часов {s} (+КДЗ 2 = {s + 2}), Σ практ.подг {sp}')


if __name__ == '__main__':
    main()
