#!/usr/bin/env python3
"""
Builds the TFS Customer Care API Ranking Methodology document for one
workgroup, driven entirely by the WORKGROUPS config at the bottom.

    pip install python-docx
    python build_api_methodology.py english
    python build_api_methodology.py english --logo tfs_logo.png
    python build_api_methodology.py english -o /path/Custom_Name.docx

Worked-example figures are computed from the config's raw values using the
same rules as the production query (dense ranking, weighted points, divide
by KPI count, tie-break), so the numbers in the document cannot drift.

Design tokens (keep these in sync with the HTML):
    accent red    EB0A1E      heading black  141414
    body ink      262626      label gray     8A8A8A
    rule gray     D8D8D8      hairline       E2E2E2
    Headings/UI   Arial       Body/formulas  Georgia
"""

import argparse

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# ── tokens ────────────────────────────────────────────────────────────────────
RED = "EB0A1E"
HEAD = "141414"
INK = "262626"
BODY_SOFT = "333333"
TABLE_INK = "3D3D3D"
MUTED = "4A4A4A"
LABEL = "8A8A8A"
LABEL_DK = "6B6B6B"
FOOT = "9A9A9A"
RULE = "D8D8D8"
HAIR = "E2E2E2"
FILL = "F7F7F7"
BLACK_RULE = "1A1A1A"

SANS = "Arial"
SERIF = "Georgia"
MONO = "Consolas"

CONTENT_W = Inches(6.8)  # 8.5in letter - 0.85in left - 0.85in right


def rgb(hexstr):
    return RGBColor.from_string(hexstr)


# ── low-level OOXML helpers (python-docx has no public API for these) ─────────
def _pPr(paragraph):
    return paragraph._p.get_or_add_pPr()


def shade_paragraph(paragraph, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    _pPr(paragraph).append(shd)


def border_paragraph(paragraph, edges):
    """edges: {'bottom': (pt, 'RRGGBB', space_pt), ...}  pt = line weight."""
    pPr = _pPr(paragraph)
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        # w:pBdr must precede w:shd in a valid pPr
        shd = pPr.find(qn("w:shd"))
        if shd is not None:
            shd.addprevious(pBdr)
        else:
            pPr.append(pBdr)
    for edge in ("top", "left", "bottom", "right"):
        if edge not in edges:
            continue
        weight, color, space = edges[edge]
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(int(round(weight * 8))))  # eighths of a point
        el.set(qn("w:space"), str(int(round(space))))
        el.set(qn("w:color"), color)
        pBdr.append(el)


def letter_space(run, pts):
    """Expanded character spacing, in points."""
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), str(int(round(pts * 20))))  # twentieths of a point
    run._r.get_or_add_rPr().append(sp)


def cell_border(cell, edges):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        if edge not in edges:
            continue
        weight, color = edges[edge]
        el = OxmlElement("w:" + edge)
        if weight == 0:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(int(round(weight * 8))))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        borders.append(el)


def shade_cell(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def kill_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def set_grid(table, widths):
    """Write w:tblGrid explicitly. Required: with a fixed layout Word sizes
    columns from the grid, not from per-cell w:tcW."""
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(round(w * 1440))))
        grid.append(col)
    tbl.insert(tbl.index(tbl.tblPr) + 1, grid)


def fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def cell_margins(table, top=4, bottom=4, left=0, right=6):
    """Cell padding in points."""
    mar = OxmlElement("w:tblCellMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement("w:" + name)
        el.set(qn("w:w"), str(int(round(val * 20))))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    table._tbl.tblPr.append(mar)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def add_field(paragraph, instr):
    """Insert a Word field, e.g. PAGE or NUMPAGES."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr_el)
    run._r.append(end)
    return run


def style_fonts(style, name):
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), name)


# ── named styles ──────────────────────────────────────────────────────────────
def make_style(doc, name, font, size, color, *, bold=False, italic=False,
               caps=False, before=0, after=0, line=None, exact=False,
               keep_next=False, keep_lines=True, spacing=None, base="Normal"):
    try:
        st = doc.styles[name]
        # Built-in styles (e.g. "Body Text 2") already exist in the default
        # template; reuse and fully redefine them instead of adding a duplicate.
        if st.type != WD_STYLE_TYPE.PARAGRAPH:
            raise ValueError(
                "style %r already exists and is not a paragraph style" % name)
        st.element.remove(st.element.get_or_add_pPr())
        st.element.remove(st.element.get_or_add_rPr())
        st.hidden = False
        st.quick_style = True
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    st.base_style = doc.styles[base]
    st.font.name = font
    st.font.size = Pt(size)
    st.font.bold = bold
    st.font.italic = italic
    st.font.all_caps = caps
    st.font.color.rgb = rgb(color)
    style_fonts(st, font)
    pf = st.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        if exact:
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(line)
        else:
            pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    pf.widow_control = True
    if spacing is not None:
        rPr = st.element.get_or_add_rPr()
        for old in rPr.findall(qn("w:spacing")):
            rPr.remove(old)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(round(spacing * 20))))
        # w:spacing must precede w:sz etc. in CT_RPr's element order.
        rPr.insert_element_before(
            sp, "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight",
            "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
            "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
            "w:specVanish", "w:oMath")
    return st


def build_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = SERIF
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = rgb(INK)
    style_fonts(normal, SERIF)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.15

    make_style(doc, "Cover Program", SANS, 31, HEAD, bold=True, line=34,
               exact=True, after=28, keep_next=True)
    make_style(doc, "Cover Subject", SANS, 17, INK, bold=True, line=1.3, after=14)
    make_style(doc, "Cover Deck", SERIF, 12, MUTED, line=1.55, after=0)
    make_style(doc, "Brand Name", SANS, 13, HEAD, bold=True, line=1.15, after=3)
    make_style(doc, "Brand Org", SANS, 8.5, LABEL, caps=True, spacing=1.0, after=0)
    make_style(doc, "Eyebrow Red", SANS, 8.5, RED, bold=True, caps=True,
               spacing=1.6, after=10, keep_next=True)
    make_style(doc, "Eyebrow Gray", SANS, 8.5, LABEL, bold=True, caps=True,
               spacing=1.6, after=10, keep_next=True)
    make_style(doc, "Section Label", SANS, 7.5, RED, bold=True, caps=True,
               spacing=1.5, after=8, keep_next=True)
    make_style(doc, "Heading Rank", SANS, 18, HEAD, bold=True, before=0,
               after=12, keep_next=True)
    make_style(doc, "Body Text 2", SERIF, 10.5, INK, line=1.15, after=11)
    make_style(doc, "Formula", SERIF, 13, HEAD, italic=True, before=12,
               after=12, line=1.2, keep_next=True)
    make_style(doc, "Formula Note", SERIF, 9.5, MUTED, line=1.5, after=20)
    make_style(doc, "Callout", SERIF, 10, BODY_SOFT, line=1.55, after=0)
    make_style(doc, "Callout Label", SANS, 7.5, RED, bold=True, caps=True,
               spacing=1.5, after=6, keep_next=True)
    make_style(doc, "Table Head", SANS, 7.5, BLACK_RULE, bold=True, caps=True,
               spacing=1.2, after=0)
    make_style(doc, "Table Body", SANS, 9, TABLE_INK, line=1.2, after=0)
    make_style(doc, "Table Key", SANS, 9, BLACK_RULE, bold=True, line=1.2, after=0)
    make_style(doc, "Table Mono", MONO, 7.5, TABLE_INK, line=1.2, after=0)
    make_style(doc, "Def Term", SANS, 9, BLACK_RULE, bold=True, line=1.2, after=0)
    make_style(doc, "Def Body", SERIF, 10, BODY_SOFT, line=1.5, after=0)
    make_style(doc, "Def Term Sub", SANS, 7, LABEL, line=1.2, after=0)
    make_style(doc, "Flow Label", SANS, 9, HEAD, bold=True, caps=True,
               spacing=1.2, line=1.1, after=0)
    make_style(doc, "Flow Value", SANS, 15, RED, bold=True, line=1.0, after=0)
    make_style(doc, "Flow Note", SANS, 7.5, LABEL, line=1.25, after=0)
    make_style(doc, "Flow Arrow", SANS, 9, RULE, line=1.0, after=0)
    make_style(doc, "Flow Label Rev", SANS, 9.5, "FFFFFF", bold=True, caps=True,
               spacing=1.6, line=1.1, after=0)
    make_style(doc, "Flow Note Rev", SANS, 7.5, "C4C4C4", line=1.25, after=0)
    make_style(doc, "Takeaway Num", SANS, 13, RED, bold=True, line=1.0, after=0)
    make_style(doc, "Takeaway Text", SERIF, 10, INK, line=1.45, after=0)
    make_style(doc, "Def Group", SANS, 7.5, RED, bold=True, caps=True,
               spacing=1.5, after=0, keep_next=True)
    make_style(doc, "Bullet Item", SERIF, 10, INK, line=1.55, after=8)
    make_style(doc, "Scope Label In", SANS, 7.5, RED, bold=True, caps=True,
               spacing=1.5, after=9, keep_next=True)
    make_style(doc, "Scope Label Out", SANS, 7.5, LABEL_DK, bold=True, caps=True,
               spacing=1.5, after=9, keep_next=True)
    make_style(doc, "Scope Item", SERIF, 9.5, INK, line=1.45, after=7)
    make_style(doc, "Scope Note", SERIF, 9, MUTED, italic=True, line=1.5, after=0)
    make_style(doc, "Meta Label", SANS, 7.5, LABEL, caps=True, spacing=1.2, after=4)
    make_style(doc, "Meta Value", SANS, 9, BLACK_RULE, bold=True, after=0)
    make_style(doc, "Tile Label", SANS, 7.5, LABEL, caps=True, spacing=1.2, after=6)
    make_style(doc, "Tile Value", SANS, 19, HEAD, bold=True, after=0)
    make_style(doc, "Swatch Name", SANS, 8, BLACK_RULE, bold=True, after=1)
    make_style(doc, "Swatch Hex", MONO, 7.5, LABEL_DK, after=0)
    make_style(doc, "Fine Print", SANS, 7.5, FOOT, spacing=0.5, after=0)
    make_style(doc, "Logo Placeholder", MONO, 7.5, FOOT, after=0)
    make_style(doc, "Rule Bar", SANS, 1, HEAD, after=0, line=1, exact=True)
    make_style(doc, "Code", MONO, 8.5, HEAD, before=6, after=6, line=1.3)


# ── content helpers ───────────────────────────────────────────────────────────
def para(container, style, text="", align=None, after=None, before=None,
         keep_next=None):
    p = container.add_paragraph(text, style=style)
    if align is not None:
        p.alignment = align
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if keep_next is not None:
        p.paragraph_format.keep_with_next = keep_next
    return p


def red_bar(doc, weight=9, after=36):
    """The cover's red band: an empty paragraph carrying a thick bottom border."""
    p = para(doc, "Rule Bar", "", after=after)
    border_paragraph(p, {"bottom": (weight, RED, 0)})
    return p


def hairline(doc, color=RULE, weight=0.5, after=18, before=0):
    p = para(doc, "Rule Bar", "", after=after, before=before)
    border_paragraph(p, {"bottom": (weight, color, 2)})
    return p


def short_rule(doc, width_in=0.9, after=26):
    """The cover's 64px red accent dash, as a narrow bottom-bordered paragraph."""
    p = para(doc, "Rule Bar", "", after=after)
    p.paragraph_format.right_indent = CONTENT_W - Inches(width_in)
    border_paragraph(p, {"bottom": (2.25, RED, 0)})
    return p


def section_heading(doc, number, title, note=None, before=24, new_page=False):
    """Red section number as a hanging prefix on the Heading Rank paragraph."""
    p = doc.add_paragraph(style="Heading Rank")
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.page_break_before = new_page
    p.paragraph_format.left_indent = Inches(0.4)
    p.paragraph_format.first_line_indent = Inches(-0.4)
    num = p.add_run(number + "\t")
    num.font.name = SANS
    num.font.size = Pt(11)
    num.font.bold = True
    num.font.color.rgb = rgb(RED)
    p.add_run(title)
    if note:
        r = p.add_run("  " + note)
        r.font.size = Pt(11)
        r.font.bold = False
        r.font.color.rgb = rgb(LABEL)
    tabs = p.paragraph_format.tab_stops
    tabs.add_tab_stop(Inches(0.4))
    return p


def callout(doc, label, text, after=0):
    lab = doc.add_paragraph(label, style="Callout Label")
    lab.paragraph_format.left_indent = Inches(0.16)
    lab.paragraph_format.right_indent = Inches(0.12)
    lab.paragraph_format.space_before = Pt(10)
    shade_paragraph(lab, FILL)
    border_paragraph(lab, {"left": (2.25, RED, 6), "top": (0.5, FILL, 6)})

    body = doc.add_paragraph(text, style="Callout")
    body.paragraph_format.left_indent = Inches(0.16)
    body.paragraph_format.right_indent = Inches(0.12)
    body.paragraph_format.space_after = Pt(10 + after)
    shade_paragraph(body, FILL)
    border_paragraph(body, {"left": (2.25, RED, 6), "bottom": (0.5, FILL, 6)})
    return body


def bullet(doc, lead, rest):
    p = doc.add_paragraph(style="Bullet Item")
    p.paragraph_format.left_indent = Inches(0.26)
    p.paragraph_format.first_line_indent = Inches(-0.26)
    dot = p.add_run("\u25aa\t")  # small square, matching the design's red tick
    dot.font.name = SANS
    dot.font.size = Pt(9)
    dot.font.color.rgb = rgb(RED)
    lead_run = p.add_run(lead + " ")
    lead_run.font.name = SANS
    lead_run.font.size = Pt(9)
    lead_run.font.bold = True
    lead_run.font.color.rgb = rgb(BLACK_RULE)
    p.add_run(rest)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(0.26))
    return p


def data_table(doc, headers, rows, widths, aligns=None, key_cols=(0,),
               total_row=False):
    """Hairline table: bold black rule under the header, 0.5pt rules between rows."""
    aligns = aligns or ["left"] * len(headers)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    cell_margins(table, top=5, bottom=5, left=0, right=8)
    set_grid(table, widths)

    amap = {"left": WD_ALIGN_PARAGRAPH.LEFT,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER}

    head = table.rows[0]
    repeat_header(head)
    row_cant_split(head)
    for i, text in enumerate(headers):
        cell = head.cells[i]
        cell.width = Inches(widths[i])
        cell.paragraphs[0].style = doc.styles["Table Head"]
        cell.paragraphs[0].text = text
        cell.paragraphs[0].alignment = amap[aligns[i]]
        cell_border(cell, {"bottom": (1.5, BLACK_RULE)})

    for r_i, row_data in enumerate(rows):
        row = table.add_row()
        row_cant_split(row)
        last = total_row and r_i == len(rows) - 1
        for i, text in enumerate(row_data):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            p = cell.paragraphs[0]
            if last:
                p.style = doc.styles["Table Key"]
            elif i in key_cols:
                p.style = doc.styles["Table Key"]
            else:
                p.style = doc.styles["Table Body"]
            p.text = text
            p.alignment = amap[aligns[i]]
            if last:
                cell_border(cell, {"bottom": (1.5, BLACK_RULE)})
                if i == len(row_data) - 1:
                    for run in p.runs:
                        run.font.color.rgb = rgb(RED)
            else:
                cell_border(cell, {"bottom": (0.5, HAIR)})
    return table


def scope_column(cell, label, label_style, rule_color, rule_weight,
                 items, marker, marker_color):
    """One half of the scope table: ruled header + hanging-indent item list."""
    cell_border(cell, {"top": (rule_weight, rule_color)})
    lab = cell.paragraphs[0]
    lab.style = label_style
    lab.text = label
    for lead, rest in items:
        p = cell.add_paragraph(style="Scope Item")
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.first_line_indent = Inches(-0.2)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(0.2))
        m = p.add_run(marker + "\t")
        m.font.name = SANS
        m.font.size = Pt(8.5)
        m.font.color.rgb = rgb(marker_color)
        lr = p.add_run(lead)
        lr.font.name = SANS
        lr.font.size = Pt(8.5)
        lr.font.bold = True
        lr.font.color.rgb = rgb(BLACK_RULE)
        if rest:
            p.add_run("  " + rest)
    cell.paragraphs[-1].paragraph_format.space_after = Pt(2)


def scope_table(doc, in_items, out_items, after=10):
    """Two-column in-scope / out-of-scope comparison, split by a gutter."""
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    cell_margins(table, top=9, bottom=0, left=0, right=12)
    row = table.rows[0]
    widths = (3.2, 0.4, 3.2)
    set_grid(table, widths)
    for i, w in enumerate(widths):
        row.cells[i].width = Inches(w)

    scope_column(row.cells[0], "In scope \u2014 what this document governs",
                 "Scope Label In", RED, 2.25, in_items, "\u25aa", RED)
    scope_column(row.cells[2], "Out of scope \u2014 governed elsewhere",
                 "Scope Label Out", RULE, 2.25, out_items, "\u2014", LABEL)

    spacer(doc, after)
    return table


def _flow_cells(row, spans):
    """Merge a 6-column row into the given spans, returning the merged cells."""
    out = []
    start = 0
    for n in spans:
        cell = row.cells[start]
        if n > 1:
            cell = cell.merge(row.cells[start + n - 1])
        out.append(cell)
        start += n
    return out


def _flow_box(cell, label, value=None, note=None, dark=False, accent=True):
    if dark:
        shade_cell(cell, HEAD)
        cell_border(cell, {e: (0, HEAD) for e in
                           ("top", "left", "bottom", "right")})
        lab_style, note_style = "Flow Label Rev", "Flow Note Rev"
    else:
        shade_cell(cell, FILL)
        cell_border(cell, {
            "top": (2.25, RED) if accent else (0.5, RULE),
            "left": (0.5, RULE), "right": (0.5, RULE), "bottom": (0.5, RULE)})
        lab_style, note_style = "Flow Label", "Flow Note"
    p = cell.paragraphs[0]
    p.style = lab_style
    p.text = label
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if value:
        vp = cell.add_paragraph(value, style="Flow Value")
        vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        vp.paragraph_format.space_before = Pt(3)
    if note:
        np_ = cell.add_paragraph(note, style=note_style)
        np_.alignment = WD_ALIGN_PARAGRAPH.CENTER
        np_.paragraph_format.space_before = Pt(3)


def flow_diagram(doc, metrics, divisor_note, national=True):
    """Native-Word flow of the API calculation: fan-in, spine, fan-out."""
    ncol = 12
    col = round(CONTENT_W.inches / ncol, 4)
    table = doc.add_table(rows=0, cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    set_grid(table, [col] * ncol)
    cell_margins(table, top=5, bottom=5, left=6, right=6)

    def band(spans):
        row = table.add_row()
        row_cant_split(row)
        for c in row.cells:
            c.width = Inches(col)
        return _flow_cells(row, spans)

    def arrows(spans):
        row = table.add_row()
        row.height = Pt(14)
        for c in row.cells:
            c.width = Inches(col)
        for cell in _flow_cells(row, spans):
            cell_border(cell, {e: (0, RULE) for e in
                               ("top", "left", "bottom", "right")})
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.style = doc.styles["Flow Arrow"]
            p.text = "\u2193"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    n = len(metrics)
    if ncol % n:
        raise ValueError("flow_diagram supports 2, 3, 4 or 6 metrics")
    spans = [ncol // n] * n
    for cell, (name, weight, note) in zip(band(spans), metrics):
        _flow_box(cell, name, weight, note)
    arrows(spans)
    _flow_box(band([ncol])[0], "Weighted rank points",
              note="each KPI rank \u00d7 its weight", accent=False)
    arrows([ncol])
    _flow_box(band([ncol])[0], "Total points \u00f7 number of KPIs",
              note=divisor_note, accent=False)
    arrows([ncol])
    _flow_box(band([ncol])[0], "API score", note="lower is better", dark=True)
    if national:
        arrows([6, 6])
        a, b = band([6, 6])
        _flow_box(a, "Site rank", note="same workgroup, same site")
        _flow_box(b, "National rank", note="same workgroup, all sites")
    else:
        arrows([ncol])
        _flow_box(band([ncol])[0], "Site rank",
                  note="same workgroup, same site \u2014 no national rank")
    return table


def takeaways(doc, items):
    """Numbered closing points: red numeral, hairline rule between."""
    widths = (0.45, 6.35)
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    set_grid(table, widths)
    cell_margins(table, top=9, bottom=9, left=0, right=0)
    for i, text in enumerate(items, start=1):
        row = table.add_row()
        row_cant_split(row)
        for j, w in enumerate(widths):
            row.cells[j].width = Inches(w)
            cell_border(row.cells[j], {"bottom": (0.5, HAIR)})
        np_ = row.cells[0].paragraphs[0]
        np_.style = doc.styles["Takeaway Num"]
        np_.text = str(i)
        tp = row.cells[1].paragraphs[0]
        tp.style = doc.styles["Takeaway Text"]
        tp.text = text
    return table


def def_list(doc, groups, term_w=1.75, gutter=0.25):
    """Grouped definition list: red group labels, hairline rule under each term.

    groups: [(group_label | None, [(term, expansion | None, definition), ...]), ...]
    """
    body_w = round(CONTENT_W.inches - term_w - gutter, 3)
    widths = (term_w, gutter, body_w)
    table = doc.add_table(rows=0, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    cell_margins(table, top=6, bottom=6, left=0, right=0)
    set_grid(table, widths)

    def _row(rule_weight, rule_color):
        row = table.add_row()
        row_cant_split(row)
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)
            cell_border(row.cells[i], {"bottom": (rule_weight, rule_color)})
        return row

    for g_i, (label, entries) in enumerate(groups):
        if label:
            row = _row(1.0, BLACK_RULE)
            merged = row.cells[0].merge(row.cells[2])
            p = merged.paragraphs[0]
            p.style = doc.styles["Def Group"]
            p.text = label
            p.paragraph_format.space_before = Pt(0 if g_i == 0 else 14)
        for term, expansion, definition in entries:
            row = _row(0.5, HAIR)
            tp = row.cells[0].paragraphs[0]
            tp.style = doc.styles["Def Term"]
            tp.text = term
            if expansion:
                sub = row.cells[0].add_paragraph(expansion, style="Def Term Sub")
                sub.paragraph_format.space_before = Pt(1.5)
            bp = row.cells[2].paragraphs[0]
            bp.style = doc.styles["Def Body"]
            bp.text = definition
    return table


def spacer(doc, pts):
    p = para(doc, "Rule Bar", "", after=0)
    p.paragraph_format.line_spacing = Pt(pts)
    return p



def cell_dashed(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "dashed")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def kpi_table(doc, cfg):
    """Feature table for the KPI set: dark header band, red weight column."""
    kpis = cfg["kpis"]
    headers = ["KPI", "What it measures", "System of record",
               "Rank 1 goes to", "Weight"]
    widths = [1.4, 1.95, 1.5, 1.1, 0.85]
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    set_grid(table, widths)
    cell_margins(table, top=8, bottom=8, left=8, right=8)

    head = table.rows[0]
    repeat_header(head)
    row_cant_split(head)
    for i, text in enumerate(headers):
        cell = head.cells[i]
        cell.width = Inches(widths[i])
        shade_cell(cell, HEAD)
        cell_border(cell, {"top": (2.25, RED)})
        p = cell.paragraphs[0]
        p.style = doc.styles["Table Head"]
        p.text = text
        p.alignment = (WD_ALIGN_PARAGRAPH.RIGHT if i == 4
                       else WD_ALIGN_PARAGRAPH.LEFT)
        p.runs[0].font.color.rgb = rgb("FFFFFF")

    for r_i, k in enumerate(kpis):
        row = table.add_row()
        row_cant_split(row)
        vals = [k["name"], k["measures"], k["system"], k["direction"],
                pct(k["weight"])]
        for i, text in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if r_i % 2 == 0:
                shade_cell(cell, FILL)
            cell_border(cell, {"bottom": (0.5, RULE)})
            p = cell.paragraphs[0]
            p.text = text
            if i == 0:
                p.style = doc.styles["Table Key"]
                p.runs[0].font.size = Pt(10)
                cell_border(cell, {"left": (2.25, RED), "bottom": (0.5, RULE)})
            elif i == 4:
                p.style = doc.styles["Flow Value"]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.style = doc.styles["Table Body"]

    row = table.add_row()
    row_cant_split(row)
    for i in range(len(headers)):
        cell = row.cells[i]
        cell.width = Inches(widths[i])
        cell_border(cell, {"top": (1.5, BLACK_RULE)})
        p = cell.paragraphs[0]
        p.style = doc.styles["Table Key"]
        if i == 0:
            p.text = "Total"
        elif i == 4:
            p.text = "100%"
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.runs[0].font.color.rgb = rgb(RED)
    return table


def numbered(doc, items, num_w=0.45, pad=8):
    """Numbered rows: red numeral, optional bold lead, hairline between."""
    widths = (num_w, round(CONTENT_W.inches - num_w, 3))
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    kill_table_borders(table)
    fixed_layout(table)
    set_grid(table, widths)
    cell_margins(table, top=pad, bottom=pad, left=0, right=0)
    for i, item in enumerate(items, start=1):
        lead, text = item if isinstance(item, tuple) else (None, item)
        row = table.add_row()
        row_cant_split(row)
        for j, w in enumerate(widths):
            row.cells[j].width = Inches(w)
            cell_border(row.cells[j], {"bottom": (0.5, HAIR)})
        np_ = row.cells[0].paragraphs[0]
        np_.style = doc.styles["Takeaway Num"]
        np_.text = str(i)
        tp = row.cells[1].paragraphs[0]
        tp.style = doc.styles["Takeaway Text"]
        if lead:
            lr = tp.add_run(lead + "  ")
            lr.font.name = SANS
            lr.font.size = Pt(9)
            lr.font.bold = True
            lr.font.color.rgb = rgb(BLACK_RULE)
        tp.add_run(text)
    return table


def tiles(doc, items):
    """Three result tiles: (label, value, color)."""
    n = len(items)
    w = round(CONTENT_W.inches / n, 3)
    t = doc.add_table(rows=2, cols=n)
    t.autofit = False
    fixed_layout(t)
    set_grid(t, [w] * n)
    cell_margins(t, top=7, bottom=7, left=10, right=10)
    for r in t.rows:
        row_cant_split(r)
    for i, (lab, val, color) in enumerate(items):
        top, bot = t.cell(0, i), t.cell(1, i)
        for cell in (top, bot):
            cell.width = Inches(w)
            cell_border(cell, {
                "top": (0.5, RULE) if cell is top else (0, RULE),
                "left": (0.5, RULE), "right": (0.5, RULE),
                "bottom": (0.5, RULE) if cell is bot else (0, RULE)})
        lp = top.paragraphs[0]
        lp.style = doc.styles["Tile Label"]
        lp.text = lab
        lp.paragraph_format.keep_with_next = True
        letter_space(lp.runs[0], 1.2)
        vp = bot.paragraphs[0]
        vp.style = doc.styles["Tile Value"]
        vp.text = val
        vp.runs[0].font.color.rgb = rgb(color)
    return t


def result_line(doc, text):
    p = doc.add_paragraph(style="Formula")
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    border_paragraph(p, {"top": (0.5, RULE, 6), "bottom": (0.5, RULE, 6)})
    r = p.add_run(text)
    r.font.italic = False
    r.font.size = Pt(12)
    return p


# ── calculation engine (mirrors the production query) ────────────────────────
def mmss(seconds):
    return "%d:%02d" % divmod(int(round(seconds)), 60)


def fmt_value(kpi, v):
    return "\u2014" if v is None else kpi["fmt"](v)


def dense_ranks(values, higher_better):
    """Dense rank; None (no activity) ranks last, as NULLS LAST does."""
    distinct = sorted({v for v in values.values() if v is not None},
                      reverse=higher_better)
    return {k: (distinct.index(v) + 1 if v is not None else len(distinct) + 1)
            for k, v in values.items()}


def compute(cfg, ids):
    """Rank the given associates against each other. Returns rows in id order."""
    kpis = cfg["kpis"]
    ex = cfg["example"]["associates"]
    ranks = {k["key"]: dense_ranks({i: ex[i][k["key"]] for i in ids},
                                   k["higher_better"]) for k in kpis}
    rows = []
    for i in ids:
        pts = {k["key"]: ranks[k["key"]][i] * k["weight"] for k in kpis}
        total = round(sum(pts.values()), 4)
        # Divisor is the workgroup's KPI count. A missing value is ranked last
        # (NULLS LAST) and still counts; see the FY2028 change consideration.
        n_kpis = len(kpis)
        score = round(total / n_kpis, 3)
        rows.append({"id": i, "site": ex[i]["site"],
                     "ranks": {k: ranks[k][i] for k in ranks},
                     "points": pts, "total": total, "n_kpis": n_kpis,
                     "score": score, "tb": ex[i][cfg["tiebreak"]["key"]]})
    tb_hi = not cfg["tiebreak"]["lower_better"]
    keys = sorted({(r["score"], (-r["tb"] if tb_hi else r["tb"])) for r in rows})
    for r in rows:
        r["rank"] = keys.index((r["score"], (-r["tb"] if tb_hi else r["tb"]))) + 1
    return rows


def pct(w):
    return "%d%%" % int(round(w * 100))


# ── page 1: cover ─────────────────────────────────────────────────────────────
def build_cover(doc, cfg, logo_path=None, logo_width=1.4):
    red_bar(doc, weight=9, after=34)

    lockup = doc.add_table(rows=1, cols=2)
    lockup.autofit = False
    fixed_layout(lockup)
    kill_table_borders(lockup)
    set_grid(lockup, (1.6, 5.2))
    cell_margins(lockup, top=0, bottom=0, left=0, right=14)

    logo = lockup.cell(0, 0)
    logo.width = Inches(1.6)
    logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    lp = logo.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_path:
        lp.style = doc.styles["Normal"]
        lp.paragraph_format.space_before = Pt(0)
        lp.paragraph_format.space_after = Pt(0)
        lp.add_run().add_picture(logo_path, width=Inches(logo_width))
    else:
        cell_dashed(logo, "B9B9B9")
        shade_cell(logo, "FAFAFA")
        lp.style = doc.styles["Logo Placeholder"]
        lp.paragraph_format.space_before = Pt(14)
        lp.paragraph_format.space_after = Pt(14)
        lp.text = "logo goes here"

    brand = lockup.cell(0, 1)
    brand.width = Inches(5.2)
    bp = brand.paragraphs[0]
    bp.style = doc.styles["Brand Name"]
    bp.text = "Data & Analytics"
    bp.paragraph_format.left_indent = Inches(0.18)
    org = brand.add_paragraph("Service Operations", style="Brand Org")
    org.paragraph_format.left_indent = Inches(0.18)
    letter_space(org.runs[0], 1.0)

    hairline(doc, RULE, 0.5, after=0, before=14)
    spacer(doc, 56)

    para(doc, "Eyebrow Gray", "Reporting Program")
    prog = doc.add_paragraph(style="Cover Program")
    prog.add_run("TFS Customer Care")
    prog.add_run().add_break()
    prog.add_run("Performance Reporting")

    short_rule(doc, width_in=0.9, after=24)

    para(doc, "Eyebrow Red", "Subject")
    subj = para(doc, "Cover Subject",
                "Associate Performance Index\nRanking Methodology")
    subj.paragraph_format.right_indent = Inches(1.4)
    deck = para(doc, "Cover Deck",
                "How each KPI is ranked, weighted, and combined into the monthly "
                "API score and rank.")
    deck.paragraph_format.right_indent = Inches(1.8)

    spacer(doc, 22)
    para(doc, "Eyebrow Red", "Focus")
    foc = para(doc, "Cover Subject", cfg["name"], after=2)
    foc.runs[0].font.italic = True
    foc.runs[0].font.size = Pt(15)
    para(doc, "Def Term Sub", cfg.get("label", "Workgroup"))

    spacer(doc, 44)
    hairline(doc, RULE, 0.5, after=14)

    meta = doc.add_table(rows=2, cols=4)
    meta.autofit = False
    fixed_layout(meta)
    kill_table_borders(meta)
    set_grid(meta, (1.7, 1.7, 1.7, 1.7))
    cell_margins(meta, top=0, bottom=0, left=0, right=14)
    labels = ["Prepared by", "Document", "Version / effective", "Status"]
    values = ["Data & Analytics", "Methodology reference", DOC_VERSION_LABEL,
              DOC_STATUS]
    for i, (lab, val) in enumerate(zip(labels, values)):
        c0 = meta.cell(0, i)
        c0.width = Inches(1.7)
        p0 = c0.paragraphs[0]
        p0.style = doc.styles["Meta Label"]
        p0.text = lab
        letter_space(p0.runs[0], 1.2)
        c1 = meta.cell(1, i)
        c1.width = Inches(1.7)
        p1 = c1.paragraphs[0]
        p1.style = doc.styles["Meta Value"]
        p1.text = val
        if lab == "Status":
            p1.runs[0].font.color.rgb = rgb(RED)

    fine = para(doc, "Fine Print",
                "Internal use only \u2014 contains associate performance methodology. "
                "Not for external distribution.", before=14)
    letter_space(fine.runs[0], 0.5)


# ── running header / footer ───────────────────────────────────────────────────
def _small(run, size=7.5, color=FOOT, bold=False):
    run.font.name = SANS
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)


def build_running(section, cfg):
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.style = section.part.document.styles["Table Head"]
    hp.paragraph_format.space_after = Pt(2)
    hp.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    left = hp.add_run("TFS CUSTOMER CARE PERFORMANCE REPORTING")
    _small(left, 8, BLACK_RULE, True)
    letter_space(left, 1.4)
    hp.add_run("\t")
    right = hp.add_run("DATA & ANALYTICS")
    _small(right, 8, LABEL_DK)
    letter_space(right, 1.4)

    h2 = header.add_paragraph(style="Table Head")
    h2.paragraph_format.space_after = Pt(4)
    sub = h2.add_run("API PERFORMANCE RANKING METHODOLOGY (%s)"
                     % cfg["name"].upper())
    _small(sub, 7.5, LABEL_DK)
    letter_space(sub, 1.4)
    border_paragraph(h2, {"bottom": (2, RED, 6)})

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.paragraph_format.space_before = Pt(4)
    fp.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    note = fp.add_run("Internal use only \u2014 not for external distribution")
    _small(note)
    letter_space(note, 0.5)
    fp.add_run("\t")
    _small(fp.add_run(DOC_VERSION + " \u00b7 Page "))
    _small(add_field(fp, "PAGE"))
    _small(fp.add_run(" of "))
    _small(add_field(fp, "NUMPAGES"))
    border_paragraph(fp, {"top": (0.5, RULE, 6)})


# ── body ──────────────────────────────────────────────────────────────────────
def kpi_names(cfg, joiner=", ", last=" and "):
    names = [k["name"] for k in cfg["kpis"]]
    return joiner.join(names[:-1]) + last + names[-1]


def weights_phrase(cfg):
    parts = ["%s at %s" % (k["name"], pct(k["weight"])) for k in cfg["kpis"]]
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


def scope_words(cfg):
    return "site rank and national rank" if "national" in cfg["scopes"] \
        else "site rank"


def page_purpose(doc, cfg):
    section_heading(doc, "01", "What the API is", before=0)
    para(doc, "Body Text 2",
         "The Associate Performance Index (API) is how TFS Customer Care ranks "
         "associates on monthly performance. Rather than comparing raw numbers "
         "directly, it converts each key performance indicator (KPI) into a rank "
         "within a peer group and combines those ranks into one score. Every "
         "associate is measured against colleagues in the same workgroup, "
         "scored on the same KPIs, in the same month.")
    para(doc, "Body Text 2",
         "For the %s workgroup the API is built from %s. This document explains "
         "how those KPIs become a score, how the score becomes a rank, and who is "
         "included. It is the reference for the monthly report; it does not "
         "decide how the rank is used afterward." % (cfg["name"], kpi_names(cfg)),
         after=16)

    scope_table(
        doc,
        in_items=[
            ("Monthly associate rankings.",
             "How the %s API score and rank are produced each month."
             % cfg["name"]),
            ("KPIs and weights.",
             "The %s KPI set, the direction that earns rank 1, and the weight "
             "each carries." % cfg["name"]),
            ("Site and national rank." if "national" in cfg["scopes"]
             else "Site rank.",
             "Who is in each peer group and how the two ranks relate."
             if "national" in cfg["scopes"] else
             "Who is in the peer group and how the rank is produced."),
            ("Eligibility, exclusions, and tie-breakers.",
             "Who is ranked, what data is withheld, and how a tie is resolved."),
        ],
        out_items=[
            ("Compensation decisions.",
             "Rank may be an input; the applicable plan documents govern "
             "eligibility and award."),
            ("Performance management actions.",
             "Coaching, improvement plans, and disciplinary thresholds."),
            ("Attendance policies.",
             "Governed by their own policy documents."),
            ("Goal and target setting.",
             "Targets are established outside this process."),
            ("Merit review processes.",
             "Reviews may reference API rank but are governed elsewhere."),
        ],
        after=6)
    para(doc, "Scope Note",
         "Out of scope does not mean unrelated. Where API rank feeds another "
         "process, that process owns its own criteria and governance; this "
         "document defines only the number it receives.")


def page_kpis(doc, cfg):
    section_heading(doc, "02", "The KPIs", before=0, new_page=True)
    para(doc, "Body Text 2",
         "All performance data comes from official systems of record and is "
         "compiled after month-end processing is complete. No manual adjustments "
         "are applied to source data. The %s workgroup is ranked on the KPIs "
         "below; the weights sum to 100%%." % cfg["name"], after=14)
    kpi_table(doc, cfg)
    notes = [k["note"] for k in cfg["kpis"] if k.get("note")]
    if notes:
        spacer(doc, 8)
        para(doc, "Formula Note", " ".join(notes), after=6)

    spacer(doc, 14)
    para(doc, "Section Label", "Terms used in this document")
    def_list(doc, [(None, [
        ("API", "Associate Performance Index",
         "The score and rank this document produces."),
        ("KPI", "Key Performance Indicator",
         "One performance measure included in the API. Each has a weight."),
        ("Peer group", None,
         "The associates an individual is ranked against: the same workgroup, in "
         "the same month" + (", at the same site (site rank) or across all sites "
                             "(national rank)." if "national" in cfg["scopes"]
                             else ", at the same site.")),
        ("KPI rank", None,
         "An associate\u2019s position on one KPI within the peer group. Rank 1 is "
         "the best result."),
        ("API score", None,
         "The weighted ranks added together and divided by the number of KPIs. "
         "Lower is better."),
        ("API rank", None,
         "The associate\u2019s position when the peer group is ordered by API "
         "score, lowest first."),
    ])], term_w=1.55, gutter=0.2)


def page_how(doc, cfg, site_rows):
    kpis = cfg["kpis"]
    a = site_rows[0]
    section_heading(doc, "03", "How the ranking works", before=0,
                    new_page=True)
    para(doc, "Body Text 2",
         "Five steps, run once for every peer group each month. The numbers are "
         "Associate A\u2019s from the worked example on the next page.",
         after=8)

    flow_diagram(
        doc,
        metrics=[(k.get("flow", k["name"]), pct(k["weight"]),
                  k["direction"].lower() + " ranks 1") for k in kpis],
        divisor_note="%d KPIs in the %s set" % (len(kpis), cfg["name"]),
        national="national" in cfg["scopes"])

    spacer(doc, 6)
    directions = ", ".join(k["dir_phrase"] for k in kpis)
    k2 = kpis[1]
    r2 = a["ranks"][k2["key"]]
    numbered(doc, [
        ("Rank each KPI within the peer group.",
         "Rank 1 is the best result: %s. Ties share a rank with no gap after "
         "(1, 2, 2, 3). No value for a KPI ranks last on it." % directions),
        ("Multiply each rank by the KPI\u2019s weight.",
         "Associate A ranked %d on %s at %s weight: %d \u00d7 %s = %.2f points."
         % (r2, k2["name"], pct(k2["weight"]), r2, pct(k2["weight"]),
            a["points"][k2["key"]])),
        ("Add up the points.",
         "%s = %.2f total points." % (
             " + ".join("%.2f" % a["points"][k["key"]] for k in kpis), a["total"])),
        ("Divide by the number of KPIs in the %s set." % cfg["name"],
         "%.2f \u00f7 %d = %.3f. The lower the API score, the better."
         % (a["total"], a["n_kpis"], a["score"])),
        ("Order the peer group from lowest score to highest.",
         "The lowest API score is rank 1."),
    ], pad=5)

    tb = cfg["tiebreak"]
    callout(doc, "Tie-breaker \u2014 %s" % tb["name"],
            "When two associates finish with the same API score, the one with the "
            "%s %s ranks first. This is the only tie-breaker, and it applies to %s."
            % ("lower" if tb["lower_better"] else "higher", tb["long"],
               scope_words(cfg)))


def page_detail(doc, cfg):
    kpis = cfg["kpis"]
    section_heading(doc, "04", "Ranking rules in detail", before=0,
                    new_page=True)
    para(doc, "Body Text 2",
         "The rules behind step 1, stated precisely. These are what the "
         "production query applies; the wording here is the reference if a "
         "result is ever questioned.", after=12)

    para(doc, "Section Label", "Ties \u2014 dense ranking")
    para(doc, "Body Text 2",
         "KPI ranks use DENSE_RANK, not RANK or ROW_NUMBER. Tied values share a "
         "rank and the next value takes the next number. If three associates tie "
         "at rank 1, the next associate is rank 2 \u2014 never rank 4 \u2014 and "
         "no tied associate is placed ahead of another by chance.", after=8)
    data_table(
        doc,
        ["Associate", kpis[0]["name"], "DENSE_RANK  (used)", "RANK", "ROW_NUMBER"],
        [["W", "96.0", "1", "1", "1"],
         ["X", "96.0", "1", "1", "2"],
         ["Y", "96.0", "1", "1", "3"],
         ["Z", "94.5", "2", "4", "4"]],
        widths=[1.3, 1.2, 1.7, 1.2, 1.4],
        aligns=["left", "right", "right", "right", "right"])
    spacer(doc, 14)
    para(doc, "Section Label", "Sort direction and missing values")
    para(doc, "Body Text 2",
         "Each KPI is sorted so the best result is first, within one reporting "
         "month, workgroup, and site. An associate with no value on a KPI is "
         "sorted after every associate who has one, so they take the last rank "
         "on that KPI.", after=8)
    data_table(
        doc,
        ["KPI", "Sorted", "So rank 1 is", "No value"],
        [[k["name"], k["sort"], k["direction"].lower(), "ranks last"]
         for k in kpis],
        widths=[1.7, 1.6, 2.0, 1.5])
    spacer(doc, 6)
    code = doc.add_paragraph(style="Code")
    code.text = ("DENSE_RANK() OVER (PARTITION BY reporting_month, workgroup, "
                 "site ORDER BY %s %s NULLS LAST)"
                 % (kpis[0]["field"],
                    "DESC" if kpis[0]["higher_better"] else "ASC"))
    code.paragraph_format.left_indent = Inches(0.16)
    code.paragraph_format.right_indent = Inches(0.12)
    shade_paragraph(code, FILL)

    spacer(doc, 12)
    para(doc, "Section Label", "The divisor")
    para(doc, "Body Text 2",
         "Total points are divided by the number of KPIs in the %s set (%d). "
         "Because a missing value still receives a rank, the count does not "
         "change when a KPI is missing: the associate is ranked last on it and "
         "the divisor stays at %d."
         % (cfg["name"], len(kpis), len(kpis)), after=6)

    callout(doc, "Change consideration \u2014 FY 2028",
            "Under review, not in effect: an associate with no value on a KPI "
            "would receive no rank on it and the KPI would drop out of the "
            "divisor, so a month with no ASAT is scored on the remaining KPIs. "
            "Any change will be issued as a new version of this document.")


def _rank_table(doc, cfg, rows, scope_label, show_site):
    kpis = cfg["kpis"]
    ex = cfg["example"]["associates"]
    headers = ["Associate"] + (["Site"] if show_site else [])
    widths = [1.15] + ([0.6] if show_site else [])
    for k in kpis:
        headers += [k.get("short", k["name"]), "Rank"]
    body = []
    for r in rows:
        line = ["Associate " + r["id"]] + ([r["site"]] if show_site else [])
        for k in kpis:
            line += [fmt_value(k, ex[r["id"]][k["key"]]), str(r["ranks"][k["key"]])]
        body.append(line)
    rest = CONTENT_W.inches - sum(widths)
    per = rest / len(kpis)
    for _ in kpis:
        widths += [round(per * 0.62, 3), round(per * 0.38, 3)]
    aligns = ["left"] + (["left"] if show_site else []) + ["right"] * (2 * len(kpis))
    para(doc, "Section Label", scope_label)
    data_table(doc, headers, body, widths=widths, aligns=aligns)


def _result_table(doc, cfg, rows, rank_label, show_site):
    kpis = cfg["kpis"]
    tb = cfg["tiebreak"]
    ordered = sorted(rows, key=lambda r: r["rank"])
    headers = (["Associate"] + (["Site"] if show_site else [])
               + [k.get("short", k["name"]) + " rank" for k in kpis]
               + ["Total points", "API score", tb.get("short", tb["name"]),
                  rank_label])
    body = []
    for r in ordered:
        body.append(["Associate " + r["id"]] + ([r["site"]] if show_site else [])
                    + [str(r["ranks"][k["key"]]) for k in kpis]
                    + ["%.2f" % r["total"], "%.3f" % r["score"],
                       tb["fmt"](r["tb"]), str(r["rank"])])
    n = len(headers)
    fixed = {0: 1.15, n - 4: 0.85, n - 3: 0.8, n - 2: 0.6, n - 1: 0.85}
    if show_site:
        fixed[1] = 0.55
    rest = CONTENT_W.inches - sum(fixed.values())
    kw = round(rest / len(kpis), 3)
    widths = [fixed.get(i, kw) for i in range(n)]
    aligns = ["left"] + (["left"] if show_site else []) + ["right"] * (n - 1 - int(show_site))
    data_table(doc, headers, body, widths=widths, aligns=aligns)


def page_example_site(doc, cfg, site_rows):
    kpis = cfg["kpis"]
    ex = cfg["example"]
    a = site_rows[0]
    section_heading(doc, "05", "Worked example \u2014 site rank", before=0,
                    new_page=True)
    para(doc, "Body Text 2",
         "A %d-associate %s peer group at the %s site for one reporting month. "
         "Figures are illustrative." % (len(site_rows), cfg["name"], ex["site"]),
         after=14)

    _rank_table(doc, cfg, site_rows, "Step 1 \u2014 raw values and KPI ranks",
                show_site=False)

    spacer(doc, 8)
    para(doc, "Section Label", "Steps 2\u20134 \u2014 Associate A\u2019s points and score")
    pts_rows = [[k["name"], str(a["ranks"][k["key"]]), pct(k["weight"]),
                 "%.2f" % a["points"][k["key"]]] for k in kpis]
    pts_rows.append(["Total weighted points", "", "100%", "%.2f" % a["total"]])
    data_table(doc, ["KPI", "Rank", "Weight", "Points"], pts_rows,
               widths=[2.9, 1.2, 1.3, 1.4],
               aligns=["left", "right", "right", "right"], total_row=True)
    result_line(doc, "%.2f total points  \u00f7  %d KPIs  =  API score %.3f"
                % (a["total"], a["n_kpis"], a["score"]))

    spacer(doc, 6)
    para(doc, "Section Label", "Step 5 \u2014 the peer group, ordered")
    _result_table(doc, cfg, site_rows, "Site rank", show_site=False)
    spacer(doc, 8)
    para(doc, "Formula Note", ex["site_note"], after=0)

    if "national" not in cfg["scopes"]:
        spacer(doc, 8)
        para(doc, "Section Label", "Associate A \u2014 as reported")
        tiles(doc, [("API score", "%.3f" % a["score"], HEAD),
                    ("Site rank", "%d of %d" % (a["rank"], len(site_rows)), RED),
                    ("National rank", "not ranked", LABEL)])


def page_example_national(doc, cfg, site_rows, nat_rows):
    ex = cfg["example"]
    a_site = site_rows[0]
    a_nat = [r for r in nat_rows if r["id"] == a_site["id"]][0]
    section_heading(doc, "05", "Worked example \u2014 national rank",
                    note="continued", before=0, new_page=True)
    para(doc, "Body Text 2",
         "The same month at national scope: %d associates across %s. Every KPI "
         "is re-ranked against all of them; nothing from the site calculation "
         "carries over."
         % (len(nat_rows), ", ".join(cfg["sites"][:-1]) + " and "
            + cfg["sites"][-1]), after=12)

    _rank_table(doc, cfg, nat_rows, "Raw values and national KPI ranks",
                show_site=True)
    spacer(doc, 10)
    para(doc, "Section Label", "Points, score, and national rank")
    _result_table(doc, cfg, nat_rows, "National rank", show_site=True)
    spacer(doc, 6)
    para(doc, "Formula Note", ex["national_note"], after=0)

    spacer(doc, 12)
    para(doc, "Section Label", "Associate A \u2014 as reported")
    spacer(doc, 2)
    tiles(doc, [("Site API score", "%.3f" % a_site["score"], HEAD),
                ("Site rank", "%d of %d" % (a_site["rank"], len(site_rows)), HEAD),
                ("National rank", "%d of %d" % (a_nat["rank"], len(nat_rows)), RED)])
    spacer(doc, 8)
    para(doc, "Formula Note",
         "Site rank %d and national rank %d are answers to two different "
         "questions, measured against two different peer groups. A strong site "
         "rank does not guarantee a strong national rank, and that is expected."
         % (a_site["rank"], a_nat["rank"]), after=0)


def page_rules(doc, cfg):
    tb = cfg["tiebreak"]
    section_heading(doc, "06", "Eligibility, exclusions & tie-breakers",
                    before=0, new_page=True)
    para(doc, "Body Text 2",
         "Only associates the production month reports in an ACTIVE status are "
         "ranked. Anyone in another status is left out of the peer group "
         "entirely \u2014 they hold no rank and affect no one else\u2019s.",
         after=12)
    items = [
        ("Active status",
         "\u2014 required for the production month, for %s." % scope_words(cfg)),
        ("No reported value",
         "\u2014 an associate with no value on a KPI for the month is ranked last "
         "on that KPI. The KPI still counts in the divisor (see \u00a704)."),
        ("Management-reported exclusions",
         "\u2014 performance data reported as excluded is withheld before ranking, "
         "such as a day of %s dropped for a vendor or system issue, scheduled "
         "training, or another approved circumstance. The associate stays in the "
         "peer group; only the affected data is removed."
         % " or ".join(k["name"].lower() if len(k["name"]) > 4 else k["name"]
                       for k in cfg["kpis"][1:])),
        ("Tie-breaker",
         "\u2014 two associates with the same API score are separated by %s: the "
         "%s value ranks first. The rule is the same for %s. Associates still "
         "level share a rank and the next rank follows with no gap."
         % (tb["long"], "lower" if tb["lower_better"] else "higher",
            scope_words(cfg))),
    ]
    for extra in cfg.get("rule_notes", []):
        items.append(extra)
    for lead, rest in items:
        bullet(doc, lead, rest)

    spacer(doc, 14)
    section_heading(doc, "07", "Key takeaways", before=6)
    spacer(doc, 4)
    t = [
        "The %s API is built from %s." % (cfg["name"], weights_phrase(cfg)),
        "Raw values are never compared directly. Each KPI becomes a rank within "
        "the peer group, so associates are measured against their peers, not "
        "against a fixed target.",
        "Ranks are weighted, added, and divided by the number of KPIs. A lower "
        "API score is a better result.",
    ]
    if "national" in cfg["scopes"]:
        t.append("Site rank and national rank are two separate calculations "
                 "against two different peer groups. Placing first at a site and "
                 "lower nationally is normal, not an error.")
    else:
        t.append("The %s workgroup is ranked at site scope only; no national "
                 "rank is produced." % cfg["name"])
    t.append("Only associates in ACTIVE status are ranked, and ties are broken "
             "by %s." % tb["long"])
    numbered(doc, t, pad=7)


def build_body(doc, cfg):
    ex = cfg["example"]
    site_ids = [i for i, v in ex["associates"].items() if v["site"] == ex["site"]]
    site_rows = compute(cfg, site_ids)
    nat_rows = compute(cfg, list(ex["associates"])) if "national" in cfg["scopes"] \
        else None

    page_purpose(doc, cfg)
    page_kpis(doc, cfg)
    page_how(doc, cfg, site_rows)
    page_detail(doc, cfg)
    page_example_site(doc, cfg, site_rows)
    if nat_rows:
        page_example_national(doc, cfg, site_rows, nat_rows)
    page_rules(doc, cfg)
    return site_rows, nat_rows


# ── assembly ──────────────────────────────────────────────────────────────────
DOC_VERSION = "v1.0"
DOC_VERSION_LABEL = "v1.0 / FY 2027"
DOC_STATUS = "Draft for review"


def build(cfg, path, logo_path=None, logo_width=1.4):
    doc = Document()
    build_styles(doc)

    cover = doc.sections[0]
    cover.top_margin = Inches(0.9)
    cover.bottom_margin = Inches(0.75)
    cover.left_margin = Inches(0.85)
    cover.right_margin = Inches(0.85)
    build_cover(doc, cfg, logo_path, logo_width)

    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.top_margin = Inches(0.7)
    body.bottom_margin = Inches(0.6)
    body.left_margin = Inches(0.85)
    body.right_margin = Inches(0.85)
    body.header_distance = Inches(0.4)
    body.footer_distance = Inches(0.35)
    build_running(body, cfg)
    site_rows, nat_rows = build_body(doc, cfg)

    doc.save(path)
    print("wrote", path)
    for label, rows in (("site", site_rows), ("national", nat_rows)):
        if rows:
            print("  %-8s" % label, "  ".join(
                "%s=%.3f/#%d" % (r["id"], r["score"], r["rank"])
                for r in sorted(rows, key=lambda r: r["rank"])))


# ── workgroup configs ─────────────────────────────────────────────────────────
def _f1(v):
    return "%.1f" % v


WORKGROUPS = {
    "english": {
        "name": "English",
        "slug": "English",
        "sites": ["ECC", "ECW", "ECE"],
        "scopes": ["site", "national"],
        "kpis": [
            {"key": "asat", "name": "ASAT", "dir_phrase": "the highest ASAT",
             "measures": "Post-interaction customer satisfaction",
             "system": "Medallia", "direction": "Highest score",
             "sort": "highest first", "field": "agent_satisfaction",
             "higher_better": True, "weight": 0.50, "fmt": _f1,
             "note": "ASAT includes final survey results after any applicable "
                     "disputes."},
            {"key": "aht", "name": "AHT", "dir_phrase": "the lowest AHT",
             "measures": "Average handle time across attributed interactions",
             "system": "NICE InContact / CXOne", "direction": "Lowest time",
             "sort": "lowest first", "field": "avg_handle_time",
             "higher_better": False, "weight": 0.25, "fmt": mmss},
            {"key": "conf", "name": "Break conformance", "short": "Conf.",
             "dir_phrase": "the break conformance closest to schedule",
             "measures": "Adherence to scheduled break windows",
             "system": "NICE CXOne", "direction": "Closest to schedule",
             "sort": "lowest first", "field": "conformance_pct_diff_from_perfect",
             "higher_better": False, "weight": 0.25, "fmt": _f1,
             "note": "Break conformance is ranked on the percentage difference "
                     "from perfect (100%) conformance, so the smallest "
                     "difference earns rank 1."},
        ],
        "tiebreak": {"key": "acw", "name": "ACW",
                     "long": "average after-call work (ACW)",
                     "lower_better": True, "fmt": mmss},
        "example": {
            "site": "ECC",
            # AHT and ACW in seconds; conf = % difference from perfect.
            "associates": {
                "A": {"site": "ECC", "asat": 94.0, "aht": 432, "conf": 1.2, "acw": 38},
                "B": {"site": "ECC", "asat": 91.5, "aht": 408, "conf": 2.8, "acw": 42},
                "C": {"site": "ECC", "asat": 88.0, "aht": 485, "conf": 4.5, "acw": 55},
                "D": {"site": "ECW", "asat": 96.2, "aht": 455, "conf": 1.2, "acw": 31},
                "E": {"site": "ECE", "asat": 91.5, "aht": 380, "conf": 3.1, "acw": 47},
            },
            "site_note":
                "Associate A places first without leading every KPI. Ranking "
                "second on AHT costs 0.50 points, but AHT carries 25% while ASAT "
                "carries 50%, so Associate B\u2019s first-place AHT cannot offset "
                "placing second on the heavier KPI.",
            "national_note":
                "Two pairs tie on a KPI (B and E on ASAT, A and D on break "
                "conformance); each pair shares a rank and the next rank follows "
                "with no gap. B and E then tie on API score at 0.833, so ACW "
                "decides: B\u2019s 0:42 beats E\u2019s 0:47. Associate A drops "
                "to second because D posted the highest ASAT of any site, and "
                "ASAT carries half the weight.",
        },
    },
}


def _pct(v):
    return "%.1f%%" % v


WORKGROUPS["support_center"] = {
    "name": "Support Center",
    "slug": "Support_Center",
    "sites": ["ECC", "ECW"],
    "scopes": ["site", "national"],
    "kpis": [
        {"key": "asat", "name": "ASAT", "dir_phrase": "the highest ASAT",
         "measures": "Post-interaction customer satisfaction (email)",
         "system": "Medallia", "direction": "Highest score",
         "sort": "highest first", "field": "email_agent_satisfaction",
         "higher_better": True, "weight": 0.50, "fmt": _f1,
         "note": "ASAT includes final survey results after any applicable "
                 "disputes."},
        {"key": "eph", "name": "Email productivity", "short": "EPH",
         "dir_phrase": "the most emails per hour",
         "measures": "Emails handled per hour of logged email time",
         "system": "Salesforce CEP / Activity Tracker",
         "direction": "Highest rate",
         "sort": "highest first", "field": "emails_per_hour",
         "higher_better": True, "weight": 0.25, "fmt": _f1},
        {"key": "conf", "name": "Break conformance", "short": "Conf.",
         "dir_phrase": "the break conformance closest to schedule",
         "measures": "Adherence to scheduled break windows",
         "system": "NICE CXOne", "direction": "Closest to schedule",
         "sort": "lowest first", "field": "conformance_pct_diff_from_perfect",
         "higher_better": False, "weight": 0.25, "fmt": _f1,
         "note": "Break conformance is ranked on the percentage difference "
                 "from perfect (100%) conformance, so the smallest "
                 "difference earns rank 1."},
    ],
    "tiebreak": {"key": "adh", "name": "Adherence", "short": "Adh.",
                 "long": "schedule adherence",
                 "lower_better": False, "fmt": _pct},
    "example": {
        "site": "ECC",
        "associates": {
            "A": {"site": "ECC", "asat": 93.0, "eph": 6.4, "conf": 1.5, "adh": 96.2},
            "B": {"site": "ECC", "asat": 90.5, "eph": 7.1, "conf": 2.4, "adh": 94.8},
            "C": {"site": "ECC", "asat": 87.0, "eph": 5.2, "conf": 3.9, "adh": 91.0},
            "D": {"site": "ECW", "asat": 95.1, "eph": 5.9, "conf": 1.5, "adh": 97.3},
            "E": {"site": "ECW", "asat": 90.5, "eph": 7.8, "conf": 2.9, "adh": 95.5},
        },
        "site_note":
            "Associate A places first without leading every KPI. Ranking "
            "second on email productivity costs 0.50 points, but that KPI "
            "carries 25% while ASAT carries 50%, so Associate B\u2019s "
            "first-place productivity cannot offset placing second on the "
            "heavier KPI.",
        "national_note":
            "Two pairs tie on a KPI (B and E on ASAT, A and D on break "
            "conformance); each pair shares a rank and the next rank follows "
            "with no gap. B and E then tie on API score at 0.833, so adherence "
            "decides: E\u2019s 95.5% beats B\u2019s 94.8%. Associate A drops "
            "to second because D posted the highest ASAT of either site, and "
            "ASAT carries half the weight.",
    },
}

WORKGROUPS["dcl_bsg"] = {
    "name": "DCL/BSG",
    "slug": "DCL_BSG",
    "sites": ["ECW"],
    "scopes": ["site"],
    "kpis": [
        {"key": "asat", "name": "ASAT", "dir_phrase": "the highest ASAT",
         "measures": "Post-interaction customer satisfaction",
         "system": "NICE Satmetrix", "direction": "Highest score",
         "sort": "highest first", "field": "satmetrix_asat",
         "higher_better": True, "weight": 0.30, "fmt": _f1,
         "note": "ASAT includes final survey results after any applicable "
                 "disputes."},
        {"key": "res", "name": "RES", "dir_phrase": "the highest RES",
         "measures": "Customer-reported issue resolution",
         "system": "NICE Satmetrix", "direction": "Highest score",
         "sort": "highest first", "field": "satmetrix_res",
         "higher_better": True, "weight": 0.30, "fmt": _f1},
        {"key": "aht", "name": "AHT", "dir_phrase": "the lowest AHT",
         "measures": "Average handle time across attributed interactions",
         "system": "NICE InContact / CXOne", "direction": "Lowest time",
         "sort": "lowest first", "field": "avg_handle_time",
         "higher_better": False, "weight": 0.20, "fmt": mmss},
        {"key": "conf", "name": "Break conformance", "short": "Conf.",
         "flow": "Conformance",
         "dir_phrase": "the break conformance closest to schedule",
         "measures": "Adherence to scheduled break windows",
         "system": "NICE CXOne", "direction": "Closest to schedule",
         "sort": "lowest first", "field": "conformance_pct_diff_from_perfect",
         "higher_better": False, "weight": 0.20, "fmt": _f1,
         "note": "Break conformance is ranked on the percentage difference "
                 "from perfect (100%) conformance, so the smallest "
                 "difference earns rank 1."},
    ],
    "tiebreak": {"key": "acw", "name": "ACW",
                 "long": "average after-call work (ACW)",
                 "lower_better": True, "fmt": mmss},
    "rule_notes": [
        ("Site rank only",
         "\u2014 this workgroup is ranked within its site. No national rank is "
         "calculated or reported for it."),
    ],
    "example": {
        "site": "ECW",
        "associates": {
            "A": {"site": "ECW", "asat": 92.0, "res": 88.5, "aht": 510, "conf": 1.8, "acw": 44},
            "B": {"site": "ECW", "asat": 92.0, "res": 85.0, "aht": 470, "conf": 2.6, "acw": 39},
            "C": {"site": "ECW", "asat": 89.5, "res": 90.0, "aht": 545, "conf": 1.8, "acw": 52},
        },
        "site_note":
            "A and B tie on ASAT, and A and C on break conformance; each pair "
            "shares a rank with no gap after. Associate A never leads a KPI "
            "outright yet places first: joint first on ASAT and conformance "
            "plus second on RES outweighs C\u2019s first-place RES and "
            "B\u2019s fastest AHT.",
    },
}


def _clone(base, **overrides):
    """A workgroup that ranks exactly like another, under its own name."""
    import copy
    cfg = copy.deepcopy(WORKGROUPS[base])
    cfg.update(overrides)
    # Keep the example's sites within the workgroup's site list.
    for assoc in cfg["example"]["associates"].values():
        if assoc["site"] not in cfg["sites"]:
            assoc["site"] = cfg["sites"][-1]
    return cfg


WORKGROUPS["title_experiment"] = _clone(
    "english", name="Title Experiment", slug="Title_Experiment",
    label="Pilot workgroup")
WORKGROUPS["bilingual"] = _clone("english", name="Bilingual", slug="Bilingual")
WORKGROUPS["web_support"] = _clone("english", name="Web Support",
                                   slug="Web_Support")
WORKGROUPS["dcl2"] = _clone("dcl_bsg", name="DCL2", slug="DCL2")
WORKGROUPS["flex_support"] = _clone("english", name="Flex Support",
                                    slug="Flex_Support", sites=["ECC", "ECE"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("workgroup", choices=sorted(WORKGROUPS),
                    help="which workgroup document to build")
    ap.add_argument("-o", "--out", default=None,
                    help="output path (default API_Ranking_Methodology_<Workgroup>.docx)")
    ap.add_argument("-l", "--logo", default=None,
                    help="path to a logo image (PNG/JPG); omit for the placeholder box")
    ap.add_argument("--logo-width", type=float, default=1.4,
                    help="logo width in inches (default 1.4; column is 1.6)")
    args = ap.parse_args()
    cfg = WORKGROUPS[args.workgroup]
    out = args.out or "API_Ranking_Methodology_%s.docx" % cfg["slug"]
    build(cfg, out, args.logo, args.logo_width)
