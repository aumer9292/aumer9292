"""Word formatting engine.

A Word input is formatted in place, so everything the tool does not
understand (images, fields, content controls, text boxes, footnotes) stays
exactly where it was. Other inputs are first written to a plain Word document
from the document model and then go through the same formatting pass.

Every text change is recorded; figures are rewritten only into the standard
presentation of the same value (1234 -> 1,234; -500 -> (500); 0 -> -).
"""

import re
from decimal import Decimal

import docx
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from .classify import PRIMARY, is_financial_table
from .extract import W, docx_body_items, unique_cells
from .numbers import YEAR, decimals_of, format_amount, parse_amount
from .proofread import safe_text_fixes, spelling_fix


NEW_PAGE = PRIMARY + ("NOTES", "DIRECTORS", "AUDITOR", "SUPPLEMENTARY")


class Changes:
    def __init__(self):
        self.text = []        # (location, before, after, reason)
        self.counts = {}

    def add(self, what, n=1):
        self.counts[what] = self.counts.get(what, 0) + n

    def edit(self, location, before, after, reason):
        self.text.append((location, before, after, reason))
        self.add(reason)


# ------------------------------------------------------------ build from model

def model_to_docx(model):
    d = docx.Document()
    for b in model.blocks:
        if b.kind == "paragraph":
            p = d.add_paragraph()
            text = b.text
            if b.role == "contents":
                # contents line from a PDF or text file: put the page number on a tab
                text = re.sub(r"(?<=[A-Za-z)])\s*(\d{1,3}(?:\s*-\s*\d{1,3})?)$", r"\t\1", text)
            if text:
                r = p.add_run(text)
                r.bold = b.bold or None
            if b.needs_review:
                p.runs[0].font.highlight_color = 7  # yellow
        else:
            w = max(1, b.width)
            t = d.add_table(rows=len(b.rows), cols=w)
            for i, row in enumerate(b.rows):
                for j in range(w):
                    txt = row.cells[j] if j < len(row.cells) else ""
                    cell = t.cell(i, j)
                    if txt:
                        run = cell.paragraphs[0].add_run(txt)
                        if row.bold:
                            run.bold = True
    return d


# ------------------------------------------------------------------ helpers

def _set_run_font(run, name=None, size=None, color=None, bold=None):
    f = run.font
    if name:
        f.name = name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(a), name)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rfonts.get(qn(a)) is not None:
                del rfonts.attrib[qn(a)]
    if size:
        f.size = Pt(size)
    if color:
        f.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        f.bold = bold


def _style_font(style, name, size, color=None, bold=None):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), name)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(a)) is not None:
            del rfonts.attrib[qn(a)]
    style.font.size = Pt(size)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        style.font.bold = bold


def _get_style(d, name, base="Normal"):
    from docx.enum.style import WD_STYLE_TYPE
    try:
        return d.styles[name]
    except KeyError:
        st = d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = d.styles[base]
        return st


def _has_object(p_el):
    return bool(p_el.findall(f".//{W}drawing") or p_el.findall(f".//{W}pict") or p_el.findall(f".//{W}object")
                or p_el.findall(f".//{W}fldChar") or p_el.findall(f".//{W}fldSimple")
                or p_el.findall(f".//{W}sectPr") or p_el.findall(f".//{W}txbxContent"))


def _has_page_break(p_el):
    for br in p_el.findall(f".//{W}br"):
        if br.get(qn("w:type")) == "page":
            return True
    ppr = p_el.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None


def _rewrite_runs(paragraph, fn, location, reason, changes):
    """Apply a text function run by run (keeps run formatting)."""
    for r in paragraph.runs:
        t = r.text
        if not t:
            continue
        new = fn(t)
        if new != t:
            r.text = new
            changes.edit(location, t, new, reason)


def _set_cell_text(cell, text):
    paras = [p for p in cell.paragraphs if p.text.strip()]
    if len(paras) != 1:
        return False
    p = paras[0]
    runs = [r for r in p.runs if r.text]
    if not runs:
        return False
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""
    return True


def _border(el_pr, tag, edges):
    """edges: {'top': ('single', 4), 'bottom': ('double', 6), ...}; value None removes."""
    b = el_pr.find(qn(tag))
    if b is None:
        b = OxmlElement(tag)
        el_pr.append(b)
    for edge, spec in edges.items():
        e = b.find(qn(f"w:{edge}"))
        if e is None:
            e = OxmlElement(f"w:{edge}")
            b.append(e)
        if spec is None:
            e.set(qn("w:val"), "nil")
            for a in ("w:sz", "w:space", "w:color"):
                if e.get(qn(a)) is not None:
                    del e.attrib[qn(a)]
        else:
            e.set(qn("w:val"), spec[0])
            e.set(qn("w:sz"), str(spec[1]))
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), "000000")
    # Word requires the order top, left/start, bottom, right/end, insideH, insideV
    order = ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"]
    kids = sorted(list(b), key=lambda k: order.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in order else 99)
    for k in list(b):
        b.remove(k)
    for k in kids:
        b.append(k)


def _cell_border(cell, **edges):
    tcpr = cell._tc.get_or_add_tcPr()
    _border(tcpr, "w:tcBorders", edges)
    _reorder_tcpr(tcpr)


_TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
               "tcPrChange"]


def _reorder_tcpr(tcpr):
    kids = list(tcpr)
    kids.sort(key=lambda k: _TCPR_ORDER.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in _TCPR_ORDER else 99)
    for k in list(tcpr):
        tcpr.remove(k)
    for k in kids:
        tcpr.append(k)


_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
              "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
              "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]


def _reorder_ppr(ppr):
    kids = list(ppr)
    kids.sort(key=lambda k: _PPR_ORDER.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in _PPR_ORDER else 99)
    for k in list(ppr):
        ppr.remove(k)
    for k in kids:
        ppr.append(k)


_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription", "tblPrChange"]


def _reorder_tblpr(tblpr):
    kids = list(tblpr)
    kids.sort(key=lambda k: _TBLPR_ORDER.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in _TBLPR_ORDER else 99)
    for k in list(tblpr):
        tblpr.remove(k)
    for k in kids:
        tblpr.append(k)


def _row_flag(row, tag):
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn(tag)) is None:
        trpr.append(OxmlElement(tag))


def _para_compact(p, align=None, size=None):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(1)
    pf.line_spacing = 1.0
    if align is not None:
        p.alignment = align


# ------------------------------------------------------------------ main pass

def format_document(d, model, settings, changes=None, from_word=True):
    changes = changes or Changes()
    items = docx_body_items(d)
    if len(items) != len(model.blocks):
        raise RuntimeError("Internal error: Word body and document model are out of step "
                           f"({len(items)} vs {len(model.blocks)} items).")
    _page_setup(d, settings, changes)
    _styles(d, settings)

    entity = (settings.get("entity_name") or "").strip().lower()
    prev_section = None
    note_style = _note_number_style(model)
    blank_run = 0
    to_delete = []
    for idx, (item, block) in enumerate(zip(items, model.blocks)):
        if block.kind == "paragraph":
            p = item
            loc = block.source
            if block.role == "blank" and not p.text.strip() and not _has_object(p._element):
                blank_run += 1
                if blank_run > 1 and not _has_page_break(p._element):
                    to_delete.append(p._element)
                    changes.add("Extra blank paragraphs removed")
                continue
            blank_run = 0
            if settings.get("apply_spacing_fixes", True) and block.role != "contents":
                _rewrite_runs(p, safe_text_fixes, loc, "Spacing corrected", changes)
            if settings.get("apply_spelling"):
                _rewrite_runs(p, spelling_fix, loc, "Spelling corrected", changes)
            _format_paragraph(p, block, settings, changes, note_style)
            # each primary statement on a new page
            if block.role == "title" and block.section in NEW_PAGE and block.section != prev_section:
                target = _title_block_start(items, model.blocks, idx, entity)
                if target is not None and target != idx and block.section in PRIMARY:
                    # the entity name above a statement title is part of the title block
                    tp = items[target]
                    tp.alignment = _title_alignment(settings)
                    tp.paragraph_format.keep_with_next = True
                    tp.paragraph_format.space_after = Pt(2)
                    for r in tp.runs:
                        _set_run_font(r, settings["style"]["heading_font"], settings["style"]["title_size"],
                                      settings["style"]["heading_color"], bold=True)
                if settings.get("statement_page_breaks") and target is not None \
                        and not _already_breaks(items, model.blocks, target):
                    items[target].paragraph_format.page_break_before = True
                    changes.add("Statements and reports started on a new page")
            if block.role == "title":
                prev_section = block.section
        else:
            blank_run = 0
            _format_table(item, block, settings, changes)
            # space after a table before the next paragraph
            if idx + 1 < len(items) and model.blocks[idx + 1].kind == "paragraph":
                items[idx + 1].paragraph_format.space_before = Pt(settings["space_after_pt"] + 2)
    if settings["style"].get("annexed_notes") or settings["style"].get("signatures"):
        _statement_endings(d, items, model.blocks, settings, changes, entity)
    for el in to_delete:
        el.getparent().remove(el)
    if settings.get("header_footer"):
        _header_footer(d, settings, changes)
    return changes


def _statement_endings(d, items, blocks, settings, changes, entity):
    """Big 4 finish for each primary statement: the 'annexed notes form an integral part' line
    and the signature block, added only where the draft does not already have them."""
    st = settings["style"]
    starts = []
    for i, b in enumerate(blocks):
        if b.kind == "paragraph" and b.role == "title" and b.section != "CONTENTS":
            if not starts or blocks[starts[-1][1]].section != b.section:
                s0 = _title_block_start(items, blocks, i, entity)
                starts.append((i if s0 is None else s0, i))
    last_note = max([int(b.note_number) for b in blocks if b.kind == "paragraph" and b.role == "note_heading"
                     and b.note_number and b.note_number.isdigit()] + [0])
    signers = ["Chief Executive", "Director"]
    if settings.get("category") == "PIC":
        signers.append("Chief Financial Officer")
    for k, (begin, title_idx) in enumerate(starts):
        sec = blocks[title_idx].section
        if sec not in PRIMARY:
            continue
        end = starts[k + 1][0] - 1 if k + 1 < len(starts) else len(blocks) - 1
        span = blocks[title_idx:end + 1]
        text = "\n".join(b.text if b.kind == "paragraph" else " ".join(" ".join(r.cells) for r in b.rows)
                         for b in span)
        if not any(b.kind == "table" for b in span):
            continue
        anchor = end
        while anchor > title_idx and (blocks[anchor].kind == "paragraph" and not blocks[anchor].text.strip()):
            anchor -= 1
        el = items[anchor]._element
        if st.get("annexed_notes") and last_note and not re.search(r"annexed notes|integral part", text, re.I):
            line = f"The annexed notes 1 to {last_note} form an integral part of these financial statements."
            p = d.add_paragraph()
            r = p.add_run(line)
            _set_run_font(r, st["font"], st["size"] - 1 if st["size"] > 9 else st["size"], st["color"])
            p.paragraph_format.space_before = Pt(10)
            el.addnext(p._p)
            el = p._p
            changes.edit(blocks[title_idx].source, "", line, "Annexed-notes statement added")
        if st.get("signatures") and not re.search(r"chief executive|director\b.*\n?.*(?:signature|sd/)", text, re.I):
            tbl = d.add_table(rows=2, cols=len(signers))
            for c, who in enumerate(signers):
                top, bottom = tbl.cell(0, c), tbl.cell(1, c)
                top.paragraphs[0].paragraph_format.space_before = Pt(42)
                para = bottom.paragraphs[0]
                r = para.add_run(who)
                _set_run_font(r, st["font"], st["size"], st["color"], bold=True)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.left_indent = para.paragraph_format.right_indent = Cm(0.6)
                _border(para._p.get_or_add_pPr(), "w:pBdr", {"top": ("single", 4)})
                _reorder_ppr(para._p.get_or_add_pPr())
            tblpr = tbl._tbl.tblPr
            _border(tblpr, "w:tblBorders", {e: None for e in ("top", "left", "bottom", "right", "insideH", "insideV")})
            _reorder_tblpr(tblpr)
            for row in tbl.rows:
                _row_flag(row, "w:cantSplit")
            el.addnext(tbl._tbl)
            changes.edit(blocks[title_idx].source, "", "  ".join(signers), "Signature block added")


def _title_block_start(items, blocks, idx, entity):
    """The paragraph a statement's page should start on: the entity name above the title, if any."""
    j = idx - 1
    while j >= 0 and blocks[j].kind == "paragraph" and blocks[j].role == "blank":
        j -= 1
    if j >= 0 and blocks[j].kind == "paragraph" and len(blocks[j].text) < 100 and \
            (blocks[j].role == "heading" or (entity and blocks[j].text.strip().lower() == entity)):
        k = j - 1
        while k >= 0 and blocks[k].kind == "paragraph" and blocks[k].role == "blank":
            k -= 1
        if k < 0:
            return None
        return j
    k = idx - 1
    while k >= 0 and blocks[k].kind == "paragraph" and blocks[k].role == "blank":
        k -= 1
    return None if k < 0 else idx


def _already_breaks(items, blocks, target):
    if _has_page_break(items[target]._element):
        return True
    k = target - 1
    while k >= 0 and blocks[k].kind == "paragraph" and blocks[k].role == "blank":
        if _has_page_break(items[k]._element):
            return True
        k -= 1
    if k >= 0 and blocks[k].kind == "paragraph":
        el = items[k]._element
        if el.find(f".//{W}sectPr") is not None or _has_page_break(el):
            return True
        last = list(el.iter(f"{W}br"))
        if last and last[-1].get(qn("w:type")) == "page":
            return True
    return False


def _page_setup(d, settings, changes):
    w, h = (Inches(8.5), Inches(11)) if settings.get("page_size") == "Letter" else (Cm(21.0), Cm(29.7))
    for sec in d.sections:
        landscape = sec.page_width is not None and sec.page_height is not None and sec.page_width > sec.page_height
        if landscape:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = h, w
        else:
            sec.orientation = WD_ORIENT.PORTRAIT
            sec.page_width, sec.page_height = w, h
        sec.top_margin = Cm(settings["margin_top_cm"])
        sec.bottom_margin = Cm(settings["margin_bottom_cm"])
        sec.left_margin = Cm(settings["margin_left_cm"])
        sec.right_margin = Cm(settings["margin_right_cm"])
        sec.gutter = Cm(settings.get("gutter_cm") or 0)
        sec.header_distance = Cm(1.2)
        sec.footer_distance = Cm(1.2)
    changes.add("Page size and margins standardised")


def _styles(d, settings):
    st = settings["style"]
    normal = d.styles["Normal"]
    _style_font(normal, st["font"], st["size"], st["color"])
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(settings["space_after_pt"])
    pf.line_spacing = settings["line_spacing"]
    pf.widow_control = True
    for name, size in (("Heading 1", st["title_size"]), ("Heading 2", st["h1_size"]), ("Heading 3", st["h2_size"]),
                       ("Title", st["title_size"])):
        s = _get_style(d, name)
        _style_font(s, st["heading_font"], size, st["heading_color"], bold=True)
        s.font.italic = False
        s.paragraph_format.keep_with_next = True
        s.paragraph_format.space_before = Pt(12 if name != "Heading 3" else 8)
        s.paragraph_format.space_after = Pt(4)
    for name in ("Table Grid",):
        try:
            _style_font(d.styles[name], st["font"], st["table_size"], st["color"])
        except KeyError:
            pass


def _note_number_style(model):
    dot = plain = 0
    for b in model.paragraphs():
        if b.role == "note_heading":
            if re.match(r"^\s*\d+\.\s", b.text):
                dot += 1
            elif re.match(r"^\s*\d+\s", b.text):
                plain += 1
    return "dot" if dot >= plain else "plain"


def _format_paragraph(p, block, settings, changes, note_style):
    st = settings["style"]
    role = block.role
    pf = p.paragraph_format
    pf.widow_control = True
    if role in ("contents",):
        for r in p.runs:
            _set_run_font(r, st["font"], st["size"], st["color"])
        pf.space_after = Pt(3)
        if "\t" in p.text:
            # page numbers on a right-aligned tab with a dot leader
            sec = p.part.document.sections[-1]
            width = sec.page_width - sec.left_margin - sec.right_margin
            ppr = p._p.get_or_add_pPr()
            old_tabs = ppr.find(qn("w:tabs"))
            if old_tabs is not None:
                ppr.remove(old_tabs)
            pf.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        return
    if role == "title":
        if settings["style"].get("heading_caps") == "upper_titles":
            _rewrite_runs(p, lambda t: t.upper(), block.source, "Statement titles capitalised", changes)
        for r in p.runs:
            _set_run_font(r, st["heading_font"], st["title_size"], st["heading_color"], bold=True)
            r.font.italic = False
        p.alignment = _title_alignment(settings) if block.section in PRIMARY else WD_ALIGN_PARAGRAPH.LEFT
        pf.keep_with_next = True
        pf.space_before = Pt(0 if block.section in PRIMARY else 12)
        pf.space_after = Pt(2)
        pf.left_indent = pf.first_line_indent = None
        return
    if role in ("period", "unit"):
        if role == "period" and st.get("period_caps") and block.section in PRIMARY + ("NOTES",):
            _rewrite_runs(p, lambda t: t.upper(), block.source, "Statement titles capitalised", changes)
        for r in p.runs:
            _set_run_font(r, st["heading_font"], st["size"], st["heading_color"], bold=(role == "period"))
        p.alignment = _title_alignment(settings) if block.section in PRIMARY else p.alignment
        pf.keep_with_next = True
        pf.space_after = Pt(2 if role == "period" else 8)
        return
    if role == "heading":
        is_entity = block.section in PRIMARY or block.section == "COVER"
        for r in p.runs:
            _set_run_font(r, st["heading_font"], st["title_size"] if is_entity and block.text.isupper() else st["h1_size"],
                          st["heading_color"], bold=True)
        if block.section in PRIMARY and block.text.isupper():
            p.alignment = _title_alignment(settings)
        elif p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.keep_with_next = True
        pf.space_before = Pt(10)
        pf.space_after = Pt(4)
        return
    if role == "note_heading":
        # one numbering style for note headings: "12. Revenue"
        txt = p.text
        m = re.match(r"^(\s*)(\d+)(\.?)(\s+)", txt)    # top-level notes only; 3.1 stays 3.1
        if m:
            want_dot = note_style == "dot"
            has_dot = bool(m.group(3))
            if want_dot != has_dot and p.runs and p.runs[0].text.startswith(m.group(0)):
                old = p.runs[0].text
                new_prefix = m.group(2) + ("." if want_dot else "") + " "
                p.runs[0].text = new_prefix + old[len(m.group(0)):]
                changes.edit(block.source, old, p.runs[0].text, "Note numbering made consistent")
        sub = block.note_number and "." in block.note_number
        if st.get("note_caps") and not sub:
            _rewrite_runs(p, lambda t: t.upper(), block.source, "Note headings capitalised", changes)
        if st.get("note_indent_cm"):
            _hang_number(p, st["note_indent_cm"])
        for r in p.runs:
            _set_run_font(r, st["heading_font"], st["h2_size"] if sub else st["h1_size"], st["heading_color"], bold=True)
        pf.keep_with_next = True
        pf.space_before = Pt(10 if not sub else 6)
        pf.space_after = Pt(4)
        if p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return
    # body text
    for r in p.runs:
        _set_run_font(r, st["font"], st["size"], st["color"])
    pf.space_after = Pt(settings["space_after_pt"])
    pf.line_spacing = settings["line_spacing"]
    if settings.get("justify_body") and len(block.text) > 110 and p.alignment in (None, WD_ALIGN_PARAGRAPH.LEFT,
                                                                                  WD_ALIGN_PARAGRAPH.JUSTIFY):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if block.text.rstrip().endswith(":"):
        pf.keep_with_next = True
    ind = st.get("note_indent_cm")
    if ind and block.section in ("NOTES", "POLICIES", "ESTIMATES"):
        if not _hang_number(p, ind):
            pf.left_indent = Cm(ind)
            pf.first_line_indent = None


def _title_alignment(settings):
    return WD_ALIGN_PARAGRAPH.LEFT if settings["style"].get("title_align") == "left" else WD_ALIGN_PARAGRAPH.CENTER


def _hang_number(p, indent_cm):
    """'3.1 Revenue is ...' -> number at the margin, text from the indent (Big 4 note layout)."""
    runs = [r for r in p.runs if r.text]
    if not runs:
        return False
    m = re.match(r"^(\s*)(\d{1,2}(?:\.\d{1,2}){0,3}\.?)([ \t]+)(?=\S)", runs[0].text)
    if not m:
        return False
    runs[0].text = m.group(2) + "\t" + runs[0].text[m.end():]
    pf = p.paragraph_format
    pf.left_indent = Cm(indent_cm)
    pf.first_line_indent = -Cm(indent_cm)
    ppr = p._p.get_or_add_pPr()
    old = ppr.find(qn("w:tabs"))
    if old is not None:
        ppr.remove(old)
    pf.tab_stops.add_tab_stop(Cm(indent_cm))
    return True


def _text_width(d):
    sec = d.sections[-1]
    return sec.page_width - sec.left_margin - sec.right_margin


def _format_table(t, block, settings, changes):
    st = settings["style"]
    fin = is_financial_table(block)
    rows = list(t.rows)
    # fonts and compact paragraphs everywhere
    for row in rows:
        for cell in unique_cells(row):
            for p in cell.paragraphs:
                _para_compact(p)
                if settings.get("apply_spacing_fixes", True):
                    _rewrite_runs(p, safe_text_fixes, block.source, "Spacing corrected", changes)
                if settings.get("apply_spelling"):
                    _rewrite_runs(p, spelling_fix, block.source, "Spelling corrected", changes)
                for r in p.runs:
                    _set_run_font(r, st["font"], st["table_size"], st["color"])
        _row_flag(row, "w:cantSplit")
    if not fin:
        changes.add("Tables restyled")
        return
    w = block.width
    amount, note = set(block.amount_cols), block.note_col
    # widths
    tw = _text_width(t.part.document)
    indent = Cm(st["note_indent_cm"]) if st.get("note_indent_cm") and block.section in ("NOTES", "POLICIES", "ESTIMATES") else 0
    tw -= indent
    n_amt = len(block.amount_cols)
    note_w = Cm(1.3) if note is not None else 0
    amt_w = Cm(2.6) if n_amt <= 3 else max(Cm(1.8), int((tw * 0.68 - note_w) / n_amt))
    label_w = max(Cm(4), tw - note_w - amt_w * n_amt - Cm(0.01) * (w - 1 - n_amt - (1 if note is not None else 0)))
    other = [c for c in range(1, w) if c not in amount and c != note]
    if other:
        label_w = max(Cm(3), label_w - Cm(1.5) * len(other))
    widths = []
    for c in range(w):
        widths.append(label_w if c == 0 else note_w if c == note else amt_w if c in amount else Cm(1.5))
    tbl = t._tbl
    tblpr = tbl.tblPr
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:type"), "dxa")
    tblw.set(qn("w:w"), str(int(sum(widths) / 635)))
    lay = tblpr.find(qn("w:tblLayout"))
    if lay is None:
        lay = OxmlElement("w:tblLayout")
        tblpr.append(lay)
    lay.set(qn("w:type"), "fixed")
    ind = tblpr.find(qn("w:tblInd"))
    if ind is not None:
        tblpr.remove(ind)
    if indent:
        ind = OxmlElement("w:tblInd")
        ind.set(qn("w:w"), str(int(indent / 635)))
        ind.set(qn("w:type"), "dxa")
        tblpr.append(ind)
    if settings.get("table_borders", "financial") == "financial":
        _border(tblpr, "w:tblBorders", {e: None for e in ("top", "left", "bottom", "right", "insideH", "insideV")})
    _reorder_tblpr(tblpr)
    grid = tbl.tblGrid
    gcols = grid.findall(qn("w:gridCol"))
    if len(gcols) == w:
        for gc, wd in zip(gcols, widths):
            gc.set(qn("w:w"), str(int(wd / 635)))

    # decimals per amount column, from the figures as written
    decimals = {}
    for c in amount:
        decs = [decimals_of(r.cells[c]) for r in block.rows
                if r.kind not in ("header",) and c < len(r.cells) and parse_amount(r.cells[c]) is not None]
        decimals[c] = max(decs) if decs else 0

    prev_kind = None
    for ri, (row, mrow) in enumerate(zip(rows, block.rows)):
        cells = unique_cells(row)
        if len(cells) != w:
            prev_kind = mrow.kind
            continue
        if settings.get("table_borders", "financial") == "financial":
            for cell in cells:
                _cell_border(cell, top=None, left=None, bottom=None, right=None)
        for c, cell in enumerate(cells):
            cell.width = widths[c]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
            tcpr = cell._tc.get_or_add_tcPr()
            _reorder_tcpr(tcpr)
        kind = mrow.kind
        if kind == "header":
            _row_flag(row, "w:tblHeader")
            for c, cell in enumerate(cells):
                for p in cell.paragraphs:
                    if c in amount:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    elif c == note:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.bold = True
        else:
            bold = kind in ("grand_total", "heading", "subtotal")
            for c, cell in enumerate(cells):
                for p in cell.paragraphs:
                    if c in amount:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    elif c == note:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:
                        r.font.bold = bold if (c == 0 or c in amount) else r.font.bold
                if c in amount and settings.get("reformat_figures", True) and kind != "heading":
                    _reformat_cell(cell, mrow.cells[c], decimals[c], settings, block, ri, changes)
            if kind in ("total", "grand_total") and prev_kind in ("item", "total", "subtotal"):
                for c in amount:
                    if mrow.cells[c].strip():
                        _cell_border(cells[c], top=("single", 4))
            if kind == "grand_total":
                for c in amount:
                    if mrow.cells[c].strip():
                        _cell_border(cells[c], bottom=("double", 6))
        if kind != "blank":
            prev_kind = kind
    if st.get("unit_banner"):
        _unit_banner(t, block, changes)
    # keep a statement table on one page where it fits
    if len(rows) <= 45 and block.section in PRIMARY:
        for row in rows[:-1]:
            for cell in unique_cells(row):
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True
    changes.add("Financial tables formatted")


def _unit_banner(t, block, changes):
    """A header row repeating the unit over each year ('Rupees  Rupees') becomes one
    centred banner across the figure columns: '------------ Rupees ------------'."""
    cols = block.amount_cols
    if len(cols) < 2 or cols != list(range(cols[0], cols[0] + len(cols))):
        return
    for ri, (row, mrow) in enumerate(zip(t.rows, block.rows)):
        if mrow.kind != "header":
            continue
        vals = [mrow.cells[c].strip() for c in cols]
        if not vals[0] or len(set(v.lower() for v in vals)) != 1 or re.search(r"\d{4}|note", vals[0], re.I):
            continue
        if not re.search(r"rupees|rs\.?|pkr|\$|usd|amount|['’]000|thousand|million", vals[0], re.I):
            continue
        cells = unique_cells(row)
        if len(cells) != block.width:
            return
        first, last = cells[cols[0]], cells[cols[-1]]
        merged = first.merge(last)
        text = f"{'-' * 12} {vals[0]} {'-' * 12}"
        for p in merged.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        para = merged.paragraphs[0]
        for r in para.runs[1:]:
            r._r.getparent().remove(r._r)
        if para.runs:
            para.runs[0].text = text
        else:
            para.add_run(text)
        para.runs[0].font.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        changes.edit(f"{block.source}, row {ri + 1}", " | ".join(vals), text, "Unit heading set across the columns")
        return


def _reformat_cell(cell, text, decimals, settings, block, ri, changes):
    t = text.strip()
    v = parse_amount(t)
    if v is None or not t:
        return
    if YEAR.match(t):
        return          # a bare four-digit number in a figures column may be a year: leave it
    if settings.get("zero_style", "-") == "0" and v == 0:
        new = format_amount(Decimal(0), decimals, zero="0")
    else:
        new = format_amount(v, decimals, settings.get("negative_style", "brackets"), settings.get("zero_style", "-"))
    if new == t:
        return
    if parse_amount(new) != v:
        return          # never change a value
    if _set_cell_text(cell, new):
        changes.edit(f"{block.source}, row {ri + 1}", t, new, "Figures presented consistently")



def _add_field(paragraph, instr, font):
    """A PAGE / NUMPAGES field; every part gets the footer font so the number matches the text."""
    for kind in ("begin", "instr", "separate", "text", "end"):
        run = paragraph.add_run()
        _set_run_font(run, *font)
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {instr} "
        elif kind == "text":
            el = OxmlElement("w:t")
            el.text = "1"
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        run._element.append(el)


def _header_footer(d, settings, changes):
    st = settings["style"]
    entity = (settings.get("entity_name") or "").strip()
    for i, sec in enumerate(d.sections):
        if i and sec.header.is_linked_to_previous and sec.footer.is_linked_to_previous:
            continue
        hdr = sec.header
        if not "".join(p.text for p in hdr.paragraphs).strip():
            if entity:
                p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
                r = p.add_run(entity)
                _set_run_font(r, st["font"], max(8, st["size"] - 2), "595959")
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                changes.add("Header added")
        else:
            for p in hdr.paragraphs:
                for r in p.runs:
                    _set_run_font(r, st["font"], max(8, st["size"] - 2))
        ftr = sec.footer
        has_page = any("PAGE" in (el.text or "") for el in ftr._element.iter(qn("w:instrText"))) or \
            any("PAGE" in (el.get(qn("w:instr")) or "") for el in ftr._element.iter(qn("w:fldSimple")))
        if settings.get("page_numbers") and not has_page:
            p = ftr.add_paragraph() if "".join(x.text for x in ftr.paragraphs).strip() else ftr.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            font = (st["font"], max(8, st["size"] - 2), "595959")
            for part in ("Page ", "PAGE", " of ", "NUMPAGES"):
                if part in ("PAGE", "NUMPAGES"):
                    _add_field(p, part, font)
                else:
                    _set_run_font(p.add_run(part), *font)
            changes.add("Page numbers added")
        for p in ftr.paragraphs:
            for r in p.runs:
                if r.text:
                    _set_run_font(r, st["font"], max(8, st["size"] - 2))


def format_to_file(model, settings, out_path, source_docx=None):
    if source_docx:
        d = docx.Document(source_docx)
    else:
        d = model_to_docx(model)
    changes = format_document(d, model, settings, from_word=bool(source_docx))
    d.save(out_path)
    return changes
