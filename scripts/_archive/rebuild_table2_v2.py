from docx import Document
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

SRC = '/home/z/my-project/YMK/KTP/МДК_05_02_КТП.docx'
DST = '/home/z/my-project/YMK/KTP/МДК_05_02_КТП.docx'

COL_WIDTHS = [
    318135, 1862455, 1403985, 1403985, 1403985,
    937895, 937895, 577215, 585470, 600710, 543560,
]


def make_tcpr_xml(width_emu, vmerge=None, gridspan=None):
    """Build w:tcPr XML string."""
    parts = ['<w:tcPr ' + nsdecls('w') + '>']
    parts.append(f'<w:tcW w:w="{width_emu}" w:type="dxa"/>')
    parts.append('<w:vAlign w:val="center"/>')
    if vmerge is not None:
        if vmerge == 'restart':
            parts.append('<w:vMerge w:val="restart"/>')
        else:
            parts.append('<w:vMerge/>')  # continue
    if gridspan is not None:
        parts.append(f'<w:gridSpan w:val="{gridspan}"/>')
    parts.append('<w:tcMar>')
    parts.append('<w:top w:w="0" w:type="dxa"/>')
    parts.append('<w:bottom w:w="0" w:type="dxa"/>')
    parts.append('<w:left w:w="40" w:type="dxa"/>')
    parts.append('<w:right w:w="40" w:type="dxa"/>')
    parts.append('</w:tcMar>')
    parts.append('</w:tcPr>')
    return ''.join(parts)


def make_run_xml(text, bold=False, size=10):
    """Build w:r with w:rPr and w:t."""
    b_tag = '<w:b/>' if bold else ''
    bcs_tag = '<w:bCs/>' if bold else ''
    return (
        '<w:r>'
        '<w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman" w:eastAsia="Times New Roman"/>'
        f'<w:sz w:val="{size * 2}"/>'
        f'<w:szCs w:val="{size * 2}"/>'
        f'{b_tag}{bcs_tag}'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t>'
        '</w:r>'
    )


def make_paragraph_xml(runs_xml, align='center'):
    """Build w:p with runs and alignment."""
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:jc w:val="{align}"/>'
        '</w:pPr>'
        + runs_xml +
        '</w:p>'
    )


def make_cell_xml(width_emu, text, bold=False, size=10, align='center',
                  vmerge=None, gridspan=None):
    """Build a complete w:tc element."""
    tcpr = make_tcpr_xml(width_emu, vmerge, gridspan)
    # Handle multiline text with <w:br/>
    lines = text.split('\n')
    runs = []
    for i, line in enumerate(lines):
        if i > 0:
            runs.append('<w:r><w:br/></w:r>')
        runs.append(make_run_xml(line, bold, size))
    runs_xml = ''.join(runs)
    p_xml = make_paragraph_xml(runs_xml, align)
    return f'<w:tc>{tcpr}{p_xml}</w:tc>'


def make_empty_cell_xml(width_emu, vmerge=None, gridspan=None):
    """Build an empty w:tc (used for hidden spanned cells or vMerge continues)."""
    tcpr = make_tcpr_xml(width_emu, vmerge, gridspan)
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr></w:p></w:tc>'


def main():
    doc = Document(SRC)

    # === EXTRACT DATA ===
    old_table = doc.tables[2]
    old_data = []
    for row in old_table.rows:
        cells = [c.text.strip().replace('\n', ' | ') for c in row.cells]
        old_data.append(cells)
    data_rows = old_data[2:]  # skip 2 header rows

    # === REMOVE OLD TABLE ===
    old_elem = old_table._tbl
    parent = old_elem.getparent()
    next_sib = old_elem.getnext()
    parent.remove(old_elem)

    # === BUILD NEW TABLE XML ===
    num_data = len(data_rows)
    total_w = sum(COL_WIDTHS)

    grid_xml = ''.join(f'<w:gridCol w:w="{w}"/>' for w in COL_WIDTHS)

    borders_xml = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
    )

    # --- Row 0: 8 tc elements (matching reference exactly) ---
    # [0] vMerge=restart, [1] vMerge=restart, [2] gridSpan=3, [3] gridSpan=2 vMerge=restart,
    # [4] vMerge=restart, [5] vMerge=restart, [6] vMerge=restart, [7] vMerge=restart
    row0_xml = (
        make_cell_xml(COL_WIDTHS[0], '№ занятия', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[1], 'Наименование разделов\nпрофессионального модуля,\nтем и занятий по МДК', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[2], 'Обязательная учебная нагрузка', bold=True, gridspan=3)
        + make_cell_xml(COL_WIDTHS[5], 'Коды формируемых компетенции', bold=True, vmerge='restart', gridspan=2)
        + make_cell_xml(COL_WIDTHS[7], 'Материальное и информационное\nобеспечение занятий', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[8], 'Задания для студентов', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[9], 'Формы и методы контроля', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[10], 'ФИО преподавателя', bold=True, vmerge='restart')
    )

    # --- Row 1: 10 tc elements ---
    # [0] vMerge=cont, [1] vMerge=cont, [2] vMerge=restart, [3] vMerge=restart, [4] vMerge=restart,
    # [5] gridSpan=2 vMerge=continue, [6] vMerge=cont, [7] vMerge=cont, [8] vMerge=cont, [9] vMerge=cont
    row1_xml = (
        make_empty_cell_xml(COL_WIDTHS[0], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[1], vmerge='continue')
        + make_cell_xml(COL_WIDTHS[2], 'Кол-во\nчасов', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[3], 'В форме практической подготовки', bold=True, vmerge='restart')
        + make_cell_xml(COL_WIDTHS[4], 'Вид занятия', bold=True, vmerge='restart')
        + make_empty_cell_xml(COL_WIDTHS[5], vmerge='continue', gridspan=2)
        + make_empty_cell_xml(COL_WIDTHS[7], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[8], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[9], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[10], vmerge='continue')
    )

    # --- Row 2: 11 tc elements ---
    row2_xml = (
        make_empty_cell_xml(COL_WIDTHS[0], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[1], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[2], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[3], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[4], vmerge='continue')
        + make_cell_xml(COL_WIDTHS[5], 'ОК', bold=True)
        + make_cell_xml(COL_WIDTHS[6], 'ПК', bold=True)
        + make_empty_cell_xml(COL_WIDTHS[7], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[8], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[9], vmerge='continue')
        + make_empty_cell_xml(COL_WIDTHS[10], vmerge='continue')
    )

    # --- Row 3: 11 tc, number row ---
    nums = ['1','2','3','4','5','6','7','8','9','10','11']
    row3_xml = ''.join(
        make_cell_xml(COL_WIDTHS[i], nums[i], bold=True)
        for i in range(11)
    )

    # --- Data rows: 11 tc each ---
    data_xml_parts = []
    for data in data_rows:
        num_pp = data[0]
        name = data[1]
        hours = data[2]
        pract = data[3]
        vid = data[4]
        osnash = data[5]
        zadanie = data[6]
        prim = data[7]

        is_header = (num_pp == '' and vid == '')
        is_total = name.startswith('Итого')
        bold = is_header or is_total
        name_align = 'left'

        row_xml = (
            make_cell_xml(COL_WIDTHS[0], num_pp, bold=bold)
            + make_cell_xml(COL_WIDTHS[1], name, bold=bold, align=name_align)
            + make_cell_xml(COL_WIDTHS[2], hours, bold=bold)
            + make_cell_xml(COL_WIDTHS[3], pract, bold=bold)
            + make_cell_xml(COL_WIDTHS[4], vid, bold=bold)
            + make_cell_xml(COL_WIDTHS[5], '', bold=bold)
            + make_cell_xml(COL_WIDTHS[6], '', bold=bold)
            + make_cell_xml(COL_WIDTHS[7], osnash, bold=bold)
            + make_cell_xml(COL_WIDTHS[8], zadanie, bold=bold)
            + make_cell_xml(COL_WIDTHS[9], prim, bold=bold)
            + make_cell_xml(COL_WIDTHS[10], '', bold=bold)
        )
        data_xml_parts.append(row_xml)

    # Assemble all rows
    all_rows_xml = (
        '<w:tr>' + row0_xml + '</w:tr>'
        '<w:tr>' + row1_xml + '</w:tr>'
        '<w:tr>' + row2_xml + '</w:tr>'
        '<w:tr>' + row3_xml + '</w:tr>'
    )
    for dr in data_xml_parts:
        all_rows_xml += '<w:tr>' + dr + '</w:tr>'

    # Full table XML
    tbl_xml = (
        '<w:tbl ' + nsdecls('w') + '>'
        '<w:tblPr>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        + borders_xml +
        '</w:tblPr>'
        '<w:tblGrid>' + grid_xml + '</w:tblGrid>'
        + all_rows_xml +
        '</w:tbl>'
    )

    new_tbl = parse_xml(tbl_xml)

    # Insert before the next sibling
    if next_sib is not None:
        idx = list(parent).index(next_sib)
        parent.insert(idx, new_tbl)
    else:
        parent.append(new_tbl)

    doc.save(DST)
    print(f'Saved: {DST}')
    print(f'Data rows: {num_data}, Total rows: {4 + num_data}')


if __name__ == '__main__':
    main()
