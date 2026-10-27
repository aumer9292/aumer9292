"""Lossless print set-up for an Excel workbook.

Saving a workbook through a spreadsheet library rewrites the whole file and can
drop what the library does not understand (logos in WMF/EMF, shapes, charts,
values under merged cells, full float precision). So the user's own workbook is
not re-saved. Instead the few XML elements that control printing are patched in
the statement and note sheets only, and every other part of the file is copied
byte for byte. The check afterwards proves it: all other parts are identical
and each patched sheet's cell data (<sheetData>) is unchanged.
"""

import hashlib
import posixpath
import zipfile

from lxml import etree

from .model import Issue

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
q = lambda tag: f"{{{NS}}}{tag}"

# CT_Worksheet child order (ECMA-376 Part 1, 18.3.1.99)
ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
         "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
         "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
         "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches",
         "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture",
         "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def sheet_paths(z):
    """Sheet name -> part path inside the package."""
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{REL_NS}}}Relationship")}
    out = {}
    for s in wb.iter(q("sheet")):
        t = target.get(s.get(f"{{{R_NS}}}id"), "")
        path = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
        out[s.get("name")] = path
    return out


def _child(root, tag):
    el = root.find(q(tag))
    if el is None:
        el = etree.SubElement(root, q(tag))
        _reorder(root)
    return el


def _reorder(root):
    kids = list(root)
    pos = {t: i for i, t in enumerate(ORDER)}
    kids.sort(key=lambda k: pos.get(etree.QName(k).localname, len(ORDER)) if isinstance(k.tag, str) else len(ORDER))
    for k in list(root):
        root.remove(k)
    for k in kids:
        root.append(k)


def _patch_sheet(xml, settings):
    root = etree.fromstring(xml)
    sheet_pr = root.find(q("sheetPr"))
    if sheet_pr is None:
        sheet_pr = etree.Element(q("sheetPr"))
        root.insert(0, sheet_pr)
    psp = sheet_pr.find(q("pageSetUpPr"))
    if psp is None:
        psp = etree.Element(q("pageSetUpPr"))
        # pageSetUpPr follows tabColor and outlinePr inside sheetPr
        sheet_pr.append(psp)
    psp.set("fitToPage", "1")
    for view in root.iter(q("sheetView")):
        view.set("showGridLines", "0")
    po = _child(root, "printOptions")
    po.set("horizontalCentered", "1")
    pm = _child(root, "pageMargins")
    gutter = (settings.get("gutter_cm") or 0) / 2.54
    for k, v in (("left", 0.7 + gutter), ("right", 0.5), ("top", 0.75), ("bottom", 0.75),
                 ("header", 0.3), ("footer", 0.3)):
        pm.set(k, f"{v:.3f}".rstrip("0").rstrip("."))
    ps = _child(root, "pageSetup")
    ps.set("paperSize", "1" if settings.get("page_size") == "Letter" else "9")
    ps.set("fitToWidth", "1")
    ps.set("fitToHeight", "0")
    hf = root.find(q("headerFooter"))
    if hf is None or not "".join(hf.itertext()).strip():
        hf = _child(root, "headerFooter")
        footer = hf.find(q("oddFooter"))
        if footer is None:
            footer = etree.SubElement(hf, q("oddFooter"))
        footer.text = '&C&"Arial,Regular"&8&P'
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_print_setup(src, out, sheet_names, settings):
    with zipfile.ZipFile(src) as z:
        paths = sheet_paths(z)
        targets = {paths[n] for n in sheet_names if n in paths}
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
            for info in z.infolist():
                data = z.read(info.filename)
                if info.filename in targets:
                    data = _patch_sheet(data, settings)
                w.writestr(info, data, compress_type=info.compress_type)
    return targets


def _sheet_data_digest(xml):
    root = etree.fromstring(xml)
    sd = root.find(q("sheetData"))
    return hashlib.sha256(etree.tostring(sd, method="c14n") if sd is not None else b"").hexdigest()


def verify(src, out, patched):
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        na, nb = set(a.namelist()), set(b.namelist())
        if na != nb:
            return [Issue("integrity", "error", "The formatted workbook does not contain the same parts as the "
                          "original. Do not use it.")]
        changed = []
        for name in sorted(na):
            da, db = a.read(name), b.read(name)
            if name in patched:
                if _sheet_data_digest(da) != _sheet_data_digest(db):
                    changed.append(name)
            elif da != db:
                changed.append(name)
    if changed:
        return [Issue("integrity", "error", f"Unexpected changes in {', '.join(changed[:5])}. Do not use the "
                      "formatted workbook.")]
    return [Issue("integrity", "info", f"Excel workbook: print settings set on {len(patched)} statement/note "
                  f"sheet(s); every cell, formula, logo and every other sheet is byte-for-byte unchanged.",
                  status="passed")]


def format_workbook(src, out, settings, changes, sheet_names, working=()):
    patched = patch_print_setup(src, out, sheet_names, settings)
    changes.add("Excel print settings set (paper, fit to page width, centred, page numbers)", len(patched))
    issues = verify(src, out, patched)
    if working:
        issues.append(Issue("formatting", "info", f"{len(working)} working-paper sheet(s) left exactly as they were: "
                            + ", ".join(w.strip() for w in working[:12]) + ("…" if len(working) > 12 else ""),
                            status="passed"))
    return issues
