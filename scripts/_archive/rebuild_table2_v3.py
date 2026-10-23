from docx import Document
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
from docx.shared import Pt
from lxml import etree

SRC_ORIGINAL = '/tmp/original_ktp.docx'
DST = '/home/z/my-project/YMK/KTP/МДК_05_02_КТП.docx'

# New column widths (EMU)
NEW_COLS = {
    # After col 4 (Вид учебного занятия) insert:
    5: 937895,   # ОК
    6: 937895,   # ПК
    # At the end (was col 7, now col 10):
    10: 543560,  # ФИО преподавателя
}

# Adjust existing column widths to match reference
COL_WIDTHS = [
    318135,   # 0: № занятия
    1862455,  # 1: Наименование
    1403985,  # 2: Кол-во часов
    1403985,  # 3: В форме практической подготовки
    1403985,  # 4: Вид занятия
    937895,   # 5: ОК (NEW)
    937895,   # 6: ПК (NEW)
    577215,   # 7: Материальное и информационное обеспечение
    585470,   # 8: Задания для студентов
    600710,   # 9: Формы и методы контроля
    543560,   # 10: ФИО преподавателя (NEW)
]


def make_empty_tc(width_emu):
    """Create minimal w:tc with just width, vertical centering, and empty paragraph."""
    tc = OxmlElement('w:tc')
    tcPr = OxmlElement('w:tcPr')
    tw = OxmlElement('w:tcW')
    tw.set(qn('w:w'), str(width_emu))
    tw.set(qn('w:type'), 'dxa')
    tcPr.append(tw)
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), 'center')
    tcPr.append(va)
    tc.append(tcPr)
    p = OxmlElement('w:p')
    tc.append(p)
    return tc


def set_tc_text(tc, text, bold=False, size=10, align='center'):
    """Set text in a tc, clearing existing content."""
    # Remove existing paragraphs
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    # Add new paragraph
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), align)
    pPr.append(jc)
    p.append(pPr)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size * 2))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(size * 2))
    rPr.append(szCs)
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
        bCs = OxmlElement('w:bCs')
        rPr.append(bCs)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    r.append(t)
    p.append(r)
    tc.append(p)


def set_tc_multiline(tc, lines, bold=False, size=10, align='center'):
    """Set multi-line text with <w:br/> between lines."""
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), align)
    pPr.append(jc)
    p.append(pPr)
    for i, line in enumerate(lines):
        if i > 0:
            br = OxmlElement('w:br')
            p.append(br)
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(size * 2))
        rPr.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), str(size * 2))
        rPr.append(szCs)
        if bold:
            rPr.append(OxmlElement('w:b'))
            rPr.append(OxmlElement('w:bCs'))
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = line
        r.append(t)
        p.append(r)
    tc.append(p)


def add_vmerge(tc, restart=True):
    """Add vMerge to cell's tcPr."""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    vm = OxmlElement('w:vMerge')
    if restart:
        vm.set(qn('w:val'), 'restart')
    tcPr.append(vm)


def add_gridspan(tc, span):
    """Add gridSpan to cell's tcPr."""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    gs = OxmlElement('w:gridSpan')
    gs.set(qn('w:val'), str(span))
    tcPr.append(gs)


def main():
    doc = Document(SRC_ORIGINAL)
    body = doc.element.body
    tbl = list(body)[54]  # Table 2
    tblGrid = tbl.find(qn('w:tblGrid'))
    trs = tbl.findall(qn('w:tr'))

    # === EXTRACT DATA FROM OLD TABLE (8 cols) ===
    # Old cols: 0=№, 1=Наименование, 2=Часы, 3=Практ, 4=Вид, 5=Оснащение, 6=Задание, 7=Примечание
    old_table = doc.tables[2]
    old_data = []
    for row in old_table.rows:
        cells = [c.text.strip() for c in row.cells]
        old_data.append(cells)
    data_rows = old_data[2:]  # Skip 2 header rows

    # === STEP 1: Update tblGrid - replace 8 gridCols with 11 ===
    for gc in tblGrid.findall(qn('w:gridCol')):
        tblGrid.remove(gc)
    for w in COL_WIDTHS:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)

    # === STEP 2: Remove old rows and rebuild ===
    for tr in list(trs):
        tbl.remove(tr)

    def make_data_row(num_pp, name, hours, pract, vid, osnash, zadanie, prim):
        """Build a data row with 11 tc elements."""
        tr = OxmlElement('w:tr')
        is_header = (num_pp == '' and vid == '')
        is_total = name.startswith('Итого')
        bold = is_header or is_total

        for col_idx, (text, align) in enumerate([
            (num_pp, 'center'),
            (name, 'left'),
            (hours, 'center'),
            (pract, 'center'),
            (vid, 'center'),
            ('', 'center'),   # ОК
            ('', 'center'),   # ПК
            (osnash, 'center'),
            (zadanie, 'center'),
            (prim, 'center'),
            ('', 'center'),   # ФИО преподавателя
        ]):
            tc = make_empty_tc(COL_WIDTHS[col_idx])
            if text:
                set_tc_text(tc, text, bold=bold, align=align)
            tr.append(tc)
        return tr

    # === BUILD ALL ROWS ===
    # Row 0: 8 tc (gridSpan merges)
    r0 = OxmlElement('w:tr')
    tc0_0 = make_empty_tc(COL_WIDTHS[0]); add_vmerge(tc0_0, True); set_tc_text(tc0_0, '№ занятия', bold=True); r0.append(tc0_0)
    tc0_1 = make_empty_tc(COL_WIDTHS[1]); add_vmerge(tc0_1, True); set_tc_multiline(tc0_1, ['Наименование разделов', 'профессионального модуля,', 'тем и занятий по МДК'], bold=True); r0.append(tc0_1)
    tc0_2 = make_empty_tc(COL_WIDTHS[2]); add_gridspan(tc0_2, 3); set_tc_text(tc0_2, 'Обязательная учебная нагрузка', bold=True); r0.append(tc0_2)
    tc0_5 = make_empty_tc(COL_WIDTHS[5]); add_gridspan(tc0_5, 2); add_vmerge(tc0_5, True); set_tc_text(tc0_5, 'Коды формируемых компетенции', bold=True); r0.append(tc0_5)
    tc0_7 = make_empty_tc(COL_WIDTHS[7]); add_vmerge(tc0_7, True); set_tc_multiline(tc0_7, ['Материальное и информационное', 'обеспечение занятий'], bold=True); r0.append(tc0_7)
    tc0_8 = make_empty_tc(COL_WIDTHS[8]); add_vmerge(tc0_8, True); set_tc_text(tc0_8, 'Задания для студентов', bold=True); r0.append(tc0_8)
    tc0_9 = make_empty_tc(COL_WIDTHS[9]); add_vmerge(tc0_9, True); set_tc_text(tc0_9, 'Формы и методы контроля', bold=True); r0.append(tc0_9)
    tc0_10 = make_empty_tc(COL_WIDTHS[10]); add_vmerge(tc0_10, True); set_tc_text(tc0_10, 'ФИО преподавателя', bold=True); r0.append(tc0_10)
    tbl.append(r0)

    # Row 1: 10 tc
    r1 = OxmlElement('w:tr')
    tc1_0 = make_empty_tc(COL_WIDTHS[0]); add_vmerge(tc1_0, False); r1.append(tc1_0)
    tc1_1 = make_empty_tc(COL_WIDTHS[1]); add_vmerge(tc1_1, False); r1.append(tc1_1)
    tc1_2 = make_empty_tc(COL_WIDTHS[2]); add_vmerge(tc1_2, True); set_tc_multiline(tc1_2, ['Кол-во', 'часов'], bold=True); r1.append(tc1_2)
    tc1_3 = make_empty_tc(COL_WIDTHS[3]); add_vmerge(tc1_3, True); set_tc_text(tc1_3, 'В форме практической подготовки', bold=True); r1.append(tc1_3)
    tc1_4 = make_empty_tc(COL_WIDTHS[4]); add_vmerge(tc1_4, True); set_tc_text(tc1_4, 'Вид занятия', bold=True); r1.append(tc1_4)
    tc1_5 = make_empty_tc(COL_WIDTHS[5]); add_gridspan(tc1_5, 2); add_vmerge(tc1_5, False); r1.append(tc1_5)
    tc1_7 = make_empty_tc(COL_WIDTHS[7]); add_vmerge(tc1_7, False); r1.append(tc1_7)
    tc1_8 = make_empty_tc(COL_WIDTHS[8]); add_vmerge(tc1_8, False); r1.append(tc1_8)
    tc1_9 = make_empty_tc(COL_WIDTHS[9]); add_vmerge(tc1_9, False); r1.append(tc1_9)
    tc1_10 = make_empty_tc(COL_WIDTHS[10]); add_vmerge(tc1_10, False); r1.append(tc1_10)
    tbl.append(r1)

    # Row 2: 11 tc
    r2 = OxmlElement('w:tr')
    tc2_0 = make_empty_tc(COL_WIDTHS[0]); add_vmerge(tc2_0, False); r2.append(tc2_0)
    tc2_1 = make_empty_tc(COL_WIDTHS[1]); add_vmerge(tc2_1, False); r2.append(tc2_1)
    tc2_2 = make_empty_tc(COL_WIDTHS[2]); add_vmerge(tc2_2, False); r2.append(tc2_2)
    tc2_3 = make_empty_tc(COL_WIDTHS[3]); add_vmerge(tc2_3, False); r2.append(tc2_3)
    tc2_4 = make_empty_tc(COL_WIDTHS[4]); add_vmerge(tc2_4, False); r2.append(tc2_4)
    tc2_5 = make_empty_tc(COL_WIDTHS[5]); set_tc_text(tc2_5, 'ОК', bold=True); r2.append(tc2_5)
    tc2_6 = make_empty_tc(COL_WIDTHS[6]); set_tc_text(tc2_6, 'ПК', bold=True); r2.append(tc2_6)
    tc2_7 = make_empty_tc(COL_WIDTHS[7]); add_vmerge(tc2_7, False); r2.append(tc2_7)
    tc2_8 = make_empty_tc(COL_WIDTHS[8]); add_vmerge(tc2_8, False); r2.append(tc2_8)
    tc2_9 = make_empty_tc(COL_WIDTHS[9]); add_vmerge(tc2_9, False); r2.append(tc2_9)
    tc2_10 = make_empty_tc(COL_WIDTHS[10]); add_vmerge(tc2_10, False); r2.append(tc2_10)
    tbl.append(r2)

    # Row 3: 11 tc, number row
    r3 = OxmlElement('w:tr')
    for c, num in enumerate(['1','2','3','4','5','6','7','8','9','10','11']):
        tc = make_empty_tc(COL_WIDTHS[c])
        set_tc_text(tc, num, bold=True)
        r3.append(tc)
    tbl.append(r3)

    # Data rows
    for data in data_rows:
        row = make_data_row(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7])
        tbl.append(row)

    # === SAVE ===
    doc.save(DST)

    # Verify
    tbl_check = list(doc.element.body)[54]
    trs_check = tbl_check.findall(qn('w:tr'))
    grid_check = tbl_check.find(qn('w:tblGrid')).findall(qn('w:gridCol'))
    print('Saved:', DST)
    print('Grid cols:', len(grid_check))
    print('Rows:', len(trs_check))
    print('Row 0 tc:', len(trs_check[0].findall(qn('w:tc'))))
    print('Row 1 tc:', len(trs_check[1].findall(qn('w:tc'))))
    print('Row 2 tc:', len(trs_check[2].findall(qn('w:tc'))))

    xml_size = len(etree.tostring(tbl_check, encoding='unicode'))
    print('Table 2 XML size:', xml_size)


if __name__ == '__main__':
    main()
