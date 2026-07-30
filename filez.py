def section_heading(doc, number, title, note=None, before=24):
    """Red section number as a hanging prefix on the Heading Rank paragraph."""
    p = doc.add_paragraph(style="Heading Rank")
    p.paragraph_format.space_before = Pt(before)
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


def formula(doc, parts, note=None):
    """parts: list of (text, is_subscript) tuples."""
    p = doc.add_paragraph(style="Formula")
    border_paragraph(p, {"top": (0.5, RULE, 8), "bottom": (0.5, RULE, 8)})
    for text, sub in parts:
        r = p.add_run(text)
        if sub:
            r.font.subscript = True
            r.font.italic = False
            r.font.size = Pt(8.5)
    if note:
        para(doc, "Formula Note", note)
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


def grid_table(doc, rows, widths, styles, borderless=True):
    table = doc.add_table(rows=0, cols=len(widths))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    fixed_layout(table)
    if borderless:
        kill_table_borders(table)
    cell_margins(table, top=3, bottom=3, left=0, right=10)
    for row_data in rows:
        row = table.add_row()
        for i, text in enumerate(row_data):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            p = cell.paragraphs[0]
            p.style = doc.styles[styles[i]]
            p.text = text
    return table


def spacer(doc, pts):
    p = para(doc, "Rule Bar", "", after=0)
    p.paragraph_format.line_spacing = Pt(pts)
    return p


# ── page 1: cover ─────────────────────────────────────────────────────────────
def build_cover(doc):
    red_bar(doc, weight=9, after=34)

    lockup = doc.add_table(rows=1, cols=2)
    lockup.autofit = False
    fixed_layout(lockup)
    kill_table_borders(lockup)
    cell_margins(lockup, top=0, bottom=0, left=0, right=14)

    logo = lockup.cell(0, 0)
    logo.width = Inches(1.6)
    cell_dashed(logo, "B9B9B9")
    shade_cell(logo, "FAFAFA")
    lp = logo.paragraphs[0]
    lp.style = doc.styles["Logo Placeholder"]
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp.paragraph_format.space_before = Pt(14)
    lp.paragraph_format.space_after = Pt(14)
    lp.text = "logo goes here"

    brand = lockup.cell(0, 1)
    brand.width = Inches(5.2)
    bp = brand.paragraphs[0]
    bp.style = doc.styles["Brand Name"]
    bp.text = "Data & Analytics"
    org = brand.add_paragraph("Toyota Financial Services", style="Brand Org")
    letter_space(org.runs[0], 1.0)

    hairline(doc, RULE, 0.5, after=0, before=14)
    spacer(doc, 64)

    para(doc, "Eyebrow Gray", "Reporting Program")
    prog = doc.add_paragraph(style="Cover Program")
    prog.add_run("TFS Customer Care")
    prog.add_run().add_break()
    prog.add_run("Performance Reporting")

    short_rule(doc, width_in=0.9, after=24)

    para(doc, "Eyebrow Red", "Subject")
    subj = para(doc, "Cover Subject",
                "End of Month API Performance Rank Calculation")
    subj.paragraph_format.right_indent = Inches(1.4)
    deck = para(doc, "Cover Deck",
                "How the Associate Performance Index is normalized, weighted, "
                "combined, and ranked for the monthly report.")
    deck.paragraph_format.right_indent = Inches(1.8)

    spacer(doc, 72)
    hairline(doc, RULE, 0.5, after=14)

    meta = doc.add_table(rows=2, cols=4)
    meta.autofit = False
    fixed_layout(meta)
    kill_table_borders(meta)
    cell_margins(meta, top=0, bottom=0, left=0, right=14)
    labels = ["Prepared by", "Document", "Version / effective", "Status"]
    values = ["Data & Analytics", "Methodology reference", "2.1 · July 2026",
              "Draft for review"]
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
                "Internal use only — contains associate performance methodology. "
                "Not for external distribution.", before=14)
    letter_space(fine.runs[0], 0.5)


# ── running header / footer for the body section ───────────────────────────────
def build_running(section, subject):
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.style = section.part.document.styles["Table Head"]
    hp.paragraph_format.space_after = Pt(4)
    hp.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    left = hp.add_run(subject.upper())
    left.font.name = SANS
    left.font.size = Pt(8)
    left.font.bold = True
    left.font.color.rgb = rgb(BLACK_RULE)
    letter_space(left, 1.4)
    hp.add_run("\t")
    right = hp.add_run("DATA & ANALYTICS")
    right.font.name = SANS
    right.font.size = Pt(8)
    right.font.bold = False
    right.font.color.rgb = rgb(LABEL_DK)
    letter_space(right, 1.4)
    border_paragraph(hp, {"bottom": (2, RED, 6)})

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.paragraph_format.space_before = Pt(4)
    fp.paragraph_format.tab_stops.add_tab_stop(CONTENT_W, WD_TAB_ALIGNMENT.RIGHT)
    note = fp.add_run("Internal use only — not for external distribution")
    for run_target in (note,):
        run_target.font.name = SANS
        run_target.font.size = Pt(7.5)
        run_target.font.color.rgb = rgb(FOOT)
    letter_space(note, 0.5)
    fp.add_run("\t")
    ver = fp.add_run("v2.1 · Page ")
    ver.font.name = SANS
    ver.font.size = Pt(7.5)
    ver.font.color.rgb = rgb(FOOT)
    page_run = add_field(fp, "PAGE")
    page_run.font.name = SANS
    page_run.font.size = Pt(7.5)
    page_run.font.color.rgb = rgb(FOOT)
    of = fp.add_run(" of ")
    of.font.name = SANS
    of.font.size = Pt(7.5)
    of.font.color.rgb = rgb(FOOT)
    total_run = add_field(fp, "NUMPAGES")
    total_run.font.name = SANS
    total_run.font.size = Pt(7.5)
    total_run.font.color.rgb = rgb(FOOT)
    border_paragraph(fp, {"top": (0.5, RULE, 6)})


# ── body pages ────────────────────────────────────────────────────────────────
def build_body(doc):
    # ── 01 purpose & scope ───────────────────────────────────────────────────
    section_heading(doc, "01", "Purpose & scope", before=0)
    para(doc, "Body Text 2",
         "The Associate Performance Index (API) is a single 0–100 composite score "
         "summarizing an associate's monthly performance across five weighted "
         "dimensions. This document is the authoritative description of how that "
         "score — and the rank derived from it — is produced for the End of Month "
         "report.")
    para(doc, "Body Text 2",
         "The index exists so performance can be compared fairly across associates "
         "whose portfolios differ in size, aging, and risk mix. Raw metrics are never "
         "compared directly; every metric is first converted to a peer-relative "
         "score, so an associate is measured against colleagues working comparable "
         "inventory in the same period.")
    callout(doc, "In scope / out of scope",
            "In scope: monthly ranking of eligible associates within a peer group. "
            "Out of scope: compensation calculation, disciplinary thresholds, and "
            "team- or site-level rollups, each governed by its own methodology "
            "document.", after=8)

    # ── 02 data sources ──────────────────────────────────────────────────────
    section_heading(doc, "02", "Data sources & inputs")
    para(doc, "Body Text 2",
         "All inputs are pulled after the month-end close lock, from the systems of "
         "record below. No manual adjustments are applied to source data; "
         "corrections are handled by restating the affected month.", after=14)
    data_table(
        doc,
        ["Input", "System of record", "Grain", "Notes"],
        [
            ["Account outcomes", "Servicing platform", "Account × month",
             "Status as of close lock; excludes legal hold"],
            ["Contact activity", "Telephony / CTI", "Interaction",
             "Attributed to the owning associate"],
            ["Quality reviews", "QA scorecard", "Review",
             "Minimum 4 scored reviews for coverage"],
            ["Compliance events", "Compliance log", "Event",
             "Confirmed findings only; open items unscored"],
            ["Customer feedback", "Post-interaction survey", "Response",
             "Rolling 90-day window stabilizes small samples"],
            ["Roster & schedule", "Workforce management", "Associate × day",
             "Drives eligibility, peer group, scored days"],
        ],
        widths=[1.5, 1.55, 1.3, 2.45],
    )

    doc.add_page_break()

    # ── 03 definitions ───────────────────────────────────────────────────────
    section_heading(doc, "03", "Definitions", before=0)
    grid_table(
        doc,
        [
            ["API",
             "Associate Performance Index. The 0–100 weighted composite this "
             "document defines."],
            ["Peer group",
             "The comparison set for normalization: associates sharing queue "
             "family, tenure band, and full/part-time status in the scored month."],
            ["Component score",
             "A single metric expressed as a 0–100 peer percentile, "
             "direction-corrected so higher is always better."],
            ["Coverage",
             "The share of an associate's weight with scoreable data. Below 80%, "
             "the associate is unranked for the month."],
            ["Scored day",
             "A scheduled production day on which the associate handled at least "
             "one attributed interaction."],
            ["Inverse metric",
             "A metric where a lower raw value is better (for example, roll rate). "
             "Its percentile is flipped before weighting."],
            ["Percentile band",
             "The reporting tier assigned from final rank: Top 10%, Top 10–25%, "
             "Middle 50%, Bottom 25%."],
        ],
        widths=[1.5, 5.3],
        styles=["Def Term", "Def Body"],
    )

    # ── 04 the calculation (steps 1–2) ───────────────────────────────────────
    section_heading(doc, "04", "The calculation")
    para(doc, "Body Text 2",
         "The index is produced in four ordered steps. Steps 1 and 2 run per "
         "metric; steps 3 and 4 run per associate.", after=8)

    para(doc, "Section Label", "Step 1 — Normalize each metric to a peer percentile")
    formula(doc,
            [("P", False), ("ik", True), ("  =  ( r", False), ("ik", True),
             (" − 0.5 ) / N", False), ("k", True), ("  ×  100", False)],
            "Where r-ik is associate i's ascending rank on metric k within the peer "
            "group and N-k is the count of associates with scoreable data for that "
            "metric. Ties receive the average of the tied ranks.")

    para(doc, "Section Label", "Step 2 — Correct direction")
    formula(doc,
            [("S", False), ("ik", True), ("  =  P", False), ("ik", True),
             ("  (standard)   |   100 − P", False), ("ik", True),
             ("  (inverse)", False)],
            "After this step every component score reads the same way: 100 is best, "
            "0 is worst.")

    doc.add_page_break()

    # ── 04 continued (steps 3–4) ─────────────────────────────────────────────
    section_heading(doc, "04", "The calculation", note="continued", before=0)

    para(doc, "Section Label", "Step 3 — Weight and combine")
    formula(doc,
            [("API", False), ("i", True), ("  =  Σ", False), ("k", True),
             (" ( w", False), ("k", True), (" · S", False), ("ik", True),
             (" )  /  Σ", False), ("k", True), (" w", False), ("k", True)],
            "Dividing by the sum of available weights re-bases the index when a "
            "component is missing, so a partial-coverage associate is not penalized "
            "twice. The result is rounded to one decimal.")

    para(doc, "Section Label", "Step 4 — Rank")
    formula(doc,
            [("Rank", False), ("i", True), ("  =  descending position of API", False),
             ("i", True)],
            "Ties are broken by the ordered rules in §07. Rank is always reported "
            "with its denominator (for example, 38 of 214) and its percentile band.")

    # ── 05 components & weights ──────────────────────────────────────────────
    section_heading(doc, "05", "Components & weights")
    para(doc, "Body Text 2",
         "Weights are set annually and held constant for the calendar year. Any "
         "change is versioned in this document and applied prospectively only.",
         after=14)
    data_table(
        doc,
        ["Component", "Underlying metric", "Direction", "Weight"],
        [
            ["Portfolio performance",
             "Net delinquency roll rate on owned inventory", "Inverse", "30%"],
            ["Resolution effectiveness",
             "Share of contacted accounts reaching a kept arrangement",
             "Standard", "25%"],
            ["Quality & compliance",
             "Mean QA score, less confirmed compliance findings", "Standard", "20%"],
            ["Productivity", "Attributed interactions per scored hour",
             "Standard", "15%"],
            ["Customer experience",
             "Rolling 90-day post-interaction satisfaction", "Standard", "10%"],
            ["Total", "", "", "100%"],
        ],
        widths=[1.78, 3.34, 0.95, 0.73],
        aligns=["left", "left", "center", "right"],
        total_row=True,
    )

    doc.add_page_break()

    # ── 06 exclusions ────────────────────────────────────────────────────────
    section_heading(doc, "06", "Exclusions & filters", before=0)
    para(doc, "Body Text 2",
         "Filters are applied before normalization, so excluded records never "
         "influence another associate's percentile.", after=13)
    for lead, rest in [
        ("Fewer than 15 scored days",
         "in the month — the associate is reported as unranked."),
        ("First 90 days of tenure",
         "— new associates are scored for coaching but held out of the ranked "
         "population."),
        ("Approved leave over 10 business days",
         "— scored days are prorated; if the 15-day floor is missed, unranked."),
        ("Accounts in bankruptcy, litigation, or legal hold",
         "— removed from portfolio and resolution metrics entirely."),
        ("Test, training, and internal accounts",
         "— filtered by account class before any aggregation."),
        ("Coverage below 80%",
         "of total weight — insufficient data to produce a comparable index."),
        ("Peer groups smaller than 8",
         "— merged upward to the next-broader queue family before normalizing."),
    ]:
        bullet(doc, lead, rest)

    # ── 07 tie-breakers ──────────────────────────────────────────────────────
    section_heading(doc, "07", "Tie-breakers & edge cases", before=20)
    para(doc, "Body Text 2",
         "Two associates can round to the same index. Tie-breakers are applied "
         "strictly in order, stopping at the first that separates them.", after=13)
    data_table(
        doc,
        ["Order", "Criterion", "Rationale"],
        [
            ["1", "Unrounded index to 4 decimals",
             "Resolves most ties without judgment"],
            ["2", "Higher portfolio performance score",
             "Heaviest-weighted component governs"],
            ["3", "Higher quality & compliance score",
             "Protects against volume-driven outcomes"],
            ["4", "Greater eligible account volume",
             "The larger sample is the more reliable result"],
            ["5", "Shared rank, next rank skipped",
             "Standard competition ranking (1, 2, 2, 4)"],
        ],
        widths=[0.55, 2.72, 3.53],
        key_cols=(0, 1),
    )
    callout(doc, "Restatements",
            "If a source correction changes any associate's index by more than 1.0 "
            "point, the full month is recalculated and reissued as a numbered "
            "revision. Smaller corrections are carried into the following cycle and "
            "noted in the report appendix.")

    doc.add_page_break()

    # ── 08 worked example ────────────────────────────────────────────────────
    section_heading(doc, "08", "Worked example", before=0)
    para(doc, "Body Text 2",
         "Associate A-4417, June 2026 cycle, peer group of 214 with full coverage. "
         "Figures are illustrative.", after=16)
    data_table(
        doc,
        ["Component", "Raw value", "Peer rank", "Score", "Weight", "Contribution"],
        [
            ["Portfolio performance", "4.1% roll", "47 of 214", "78.3", "30%", "23.49"],
            ["Resolution effectiveness", "61.4%", "168 of 214", "78.3", "25%", "19.58"],
            ["Quality & compliance", "94.0 QA", "190 of 214", "88.6", "20%", "17.72"],
            ["Productivity", "6.2 / hr", "120 of 214", "55.8", "15%", "8.37"],
            ["Customer experience", "4.51 / 5", "150 of 214", "69.9", "10%", "6.99"],
            ["Composite index", "", "", "", "100%", "76.15"],
        ],
        widths=[1.98, 1.0, 0.98, 0.8, 0.76, 1.28],
        aligns=["left", "right", "right", "right", "right", "right"],
        total_row=True,
    )

    spacer(doc, 14)
    tiles = doc.add_table(rows=2, cols=3)
    tiles.autofit = False
    fixed_layout(tiles)
    cell_margins(tiles, top=8, bottom=8, left=10, right=10)
    tile_data = [("Reported index", "76.2", HEAD),
                 ("Rank", "38 of 214", HEAD),
                 ("Percentile band", "Top 10–25%", RED)]
    for i, (lab, val, color) in enumerate(tile_data):
        top = tiles.cell(0, i)
        bot = tiles.cell(1, i)
        for cell in (top, bot):
            cell.width = Inches(2.27)
            cell_border(cell, {
                "top": (0.5, RULE) if cell is top else (0, RULE),
                "left": (0.5, RULE),
                "right": (0.5, RULE),
                "bottom": (0.5, RULE) if cell is bot else (0, RULE),
            })
        lp = top.paragraphs[0]
        lp.style = doc.styles["Tile Label"]
        lp.text = lab
        letter_space(lp.runs[0], 1.2)
        vp = bot.paragraphs[0]
        vp.style = doc.styles["Tile Value"]
        vp.text = val
        vp.runs[0].font.color.rgb = rgb(color)

    spacer(doc, 14)
    para(doc, "Formula Note",
         "Read the portfolio row as the direction correction in action: a 4.1% roll "
         "rate is the 47th-lowest of 214, a strong result, so the flipped percentile "
         "lands at 78.3 rather than 21.7.", after=0)

    doc.add_page_break()

    # ── appendix A ───────────────────────────────────────────────────────────
    section_heading(doc, "A", "Word template specification", before=0)
    para(doc, "Body Text 2",
         "Everything needed to rebuild this design as a reusable .dotx. All fonts "
         "ship with Office on Windows and macOS. Define each row as a named "
         "paragraph style so the document restyles from one place.", after=18)

    swatches = [("Accent red", RED), ("Heading black", HEAD), ("Body ink", INK),
                ("Label gray", LABEL), ("Callout fill", FILL)]
    sw = doc.add_table(rows=3, cols=5)
    sw.autofit = False
    fixed_layout(sw)
    kill_table_borders(sw)
    cell_margins(sw, top=0, bottom=0, left=0, right=10)
    for i, (name, hexval) in enumerate(swatches):
        chip = sw.cell(0, i)
        chip.width = Inches(1.36)
        shade_cell(chip, hexval)
        cp = chip.paragraphs[0]
        cp.style = doc.styles["Rule Bar"]
        cp.paragraph_format.line_spacing = Pt(28)
        if hexval == FILL:
            cell_border(chip, {edge: (0.5, HAIR)
                               for edge in ("top", "left", "bottom", "right")})
        np_ = sw.cell(1, i)
        np_.width = Inches(1.36)
        np_.paragraphs[0].style = doc.styles["Swatch Name"]
        np_.paragraphs[0].text = name
        hp2 = sw.cell(2, i)
        hp2.width = Inches(1.36)
        hp2.paragraphs[0].style = doc.styles["Swatch Hex"]
        hp2.paragraphs[0].text = hexval

    spacer(doc, 16)
    data_table(
        doc,
        ["Style name", "Font", "Size / weight", "Color", "Spacing & rules"],
        [
            ["Cover Program", "Arial", "31 pt Bold", "141414",
             "Exactly 34 pt line; 28 pt after"],
            ["Cover Subject", "Arial", "17 pt Bold", "262626", "1.3 line; 14 pt after"],
            ["Cover Deck", "Georgia", "12 pt Regular", "4A4A4A",
             "1.55 line; max 5\" wide"],
            ["Eyebrow", "Arial", "8.5 pt Bold, caps", "EB0A1E / 8A8A8A",
             "Character spacing expanded 1.6 pt"],
            ["Heading Rank", "Arial", "18 pt Bold", "141414",
             "24 pt before, 12 pt after; keep with next"],
            ["Section Number", "Arial", "11 pt Bold", "EB0A1E",
             "Hanging prefix on Heading Rank, 0.4\" indent"],
            ["Body Text 2", "Georgia", "10.5 pt Regular", "262626",
             "Multiple 1.15; 11 pt after; widow/orphan on"],
            ["Table Head", "Arial", "7.5 pt Bold, caps", "1A1A1A",
             "1.5 pt bottom rule 1A1A1A; repeat header row"],
            ["Table Body", "Arial", "9 pt Regular", "3D3D3D",
             "0.5 pt bottom rule E2E2E2; no vertical rules"],
            ["Formula", "Georgia italic", "13 pt", "141414",
             "0.5 pt rules above/below; keep lines together"],
            ["Callout", "Georgia", "10 pt Regular", "333333",
             "F7F7F7 shading; 2.25 pt left border EB0A1E"],
            ["Header / Footer", "Arial", "8 / 7.5 pt, caps", "6B6B6B / 9A9A9A",
             "Header 2 pt bottom rule EB0A1E; footer 0.5 pt top rule"],
        ],
        widths=[1.28, 1.0, 1.16, 1.06, 2.3],
    )
    callout(doc, "Page setup & rebuild notes",
            "Letter portrait; margins 0.7\" top, 0.85\" left and right, 0.6\" bottom. "
            "The cover sits in its own section so the running header starts on page 2. "
            "Page X of Y fields are already in the footer. The cover's red band is a "
            "paragraph with a thick bottom border, not an image — drop your logo into "
            "the dashed placeholder box and delete the border.")


# ── assembly ──────────────────────────────────────────────────────────────────
def build(path):
    doc = Document()
    build_styles(doc)

    cover = doc.sections[0]
    cover.top_margin = Inches(0.9)
    cover.bottom_margin = Inches(0.75)
    cover.left_margin = Inches(0.85)
    cover.right_margin = Inches(0.85)
    build_cover(doc)

    body_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    body_sec.top_margin = Inches(0.7)
    body_sec.bottom_margin = Inches(0.6)
    body_sec.left_margin = Inches(0.85)
    body_sec.right_margin = Inches(0.85)
    body_sec.header_distance = Inches(0.4)
    body_sec.footer_distance = Inches(0.35)
    build_running(body_sec, "End of Month API Performance Rank Calculation")
    build_body(doc)

    doc.save(path)
    print("wrote", path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-o", "--out", default="API_Ranking_Methodology.docx")
    build(ap.parse_args().out)
