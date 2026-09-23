W_TOTAL = 9411
FILL_HEAD = '1A3A5C'

# ---------------------------------------------------------------- helper docx
def _sz(run, half_pts):
    rPr = run._r.get_or_add_rPr()
    s = OxmlElement('w:sz'); s.set(qn('w:val'), str(half_pts)); rPr.append(s)

def _color_auto(run):
    rPr = run._r.get_or_add_rPr()
    c = OxmlElement('w:color'); c.set(qn('w:val'), 'auto'); rPr.append(c)

bm_id = [max(int(x) for x in re.findall(r'w:id="(\d+)"', body.xml)) + 1]
toc_no = [max(int(x) for x in re.findall(r'_Toc(\d+)', body.xml)) + 1]
toc_entries = []   # (level, text, anchor)

def heading(text, level):
    p = d.add_paragraph(style=f'Heading {level}')
    r = p.add_run(text); _color_auto(r)
    # penanda paragraf (rPr di pPr) juga 'auto' seperti heading yang ada
    pPr = p._p.get_or_add_pPr()
    rpr = OxmlElement('w:rPr'); c = OxmlElement('w:color'); c.set(qn('w:val'), 'auto')
    rpr.append(c); pPr.append(rpr)
    if level <= 2:
        anchor = f'_Toc{toc_no[0]}'; toc_no[0] += 1
        bs = OxmlElement('w:bookmarkStart'); bs.set(qn('w:id'), str(bm_id[0])); bs.set(qn('w:name'), anchor)
        be = OxmlElement('w:bookmarkEnd'); be.set(qn('w:id'), str(bm_id[0])); bm_id[0] += 1
        r._r.addprevious(bs); r._r.addnext(be)
        toc_entries.append((level, text, anchor))
    return p

def para(text, catatan=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    if catatan:
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55); _sz(r, 18)
    return p

def bullet(text):
    p = d.add_paragraph(text, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    return p

def tabel(header, rows, widths):
    assert abs(sum(widths) - W_TOTAL) <= 2, widths
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    t.style = d.styles['Table Grid']
    tbl = t._tbl
    tblPr = tbl.tblPr
    for el in list(tblPr):
        if el.tag in (qn('w:tblW'), qn('w:jc'), qn('w:tblLayout'), qn('w:tblLook')):
            tblPr.remove(el)
    tw = OxmlElement('w:tblW'); tw.set(qn('w:w'), '0'); tw.set(qn('w:type'), 'auto'); tblPr.append(tw)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'center'); tblPr.append(jc)
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    look = OxmlElement('w:tblLook')
    for k, v in [('val', '04A0'), ('firstRow', '1'), ('lastRow', '0'), ('firstColumn', '1'),
                 ('lastColumn', '0'), ('noHBand', '0'), ('noVBand', '1')]:
        look.set(qn('w:' + k), v)
    tblPr.append(look)
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(w))
    for ri, row in enumerate([header] + rows):
        tr = tbl.tr_lst[ri]
        trPr = tr.get_or_add_trPr()
        j = OxmlElement('w:jc'); j.set(qn('w:val'), 'center'); trPr.append(j)
        for ci, (cell, w) in enumerate(zip(t.rows[ri].cells, widths)):
            tcPr = cell._tc.get_or_add_tcPr()
            tcw = OxmlElement('w:tcW'); tcw.set(qn('w:w'), str(w)); tcw.set(qn('w:type'), 'dxa'); tcPr.append(tcw)
            if ri == 0:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), FILL_HEAD); tcPr.append(shd)
            p = cell.paragraphs[0]
            txt = str(row[ci])
            r = p.add_run(txt)
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _sz(r, 18)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def fmt(x, nd=1):
    s = f'{x:,.{nd}f}' if nd else f'{int(round(x)):,}'
    return s.replace(',', '¤').replace('.', ',').replace('¤', '.')

def run(cmd):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True).stdout.strip()
