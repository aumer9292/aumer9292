"""Command line: python -m fsreview INPUT [options]   or   python -m fsreview --serve"""

import argparse
import os
import sys

from .templates import PRESETS


def main(argv=None):
    ap = argparse.ArgumentParser(prog="fsreview", description="Format, proofread and review financial statements.")
    ap.add_argument("input", nargs="?", help="Financial statements: .docx, .xlsx, .xls, .pdf or .txt")
    ap.add_argument("-o", "--out", help="Output folder (default: '<input> - review' next to the input)")
    ap.add_argument("--serve", action="store_true", help="Start the browser interface")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--reference", help="A previously approved Word document to copy the house style from")
    ap.add_argument("--entity", dest="entity_name", help="Entity name (used in headers)")
    ap.add_argument("--category", choices=["PIC", "LSC", "MSC", "SSC"],
                    help="Entity category under the Companies Act 2017. If omitted you are asked (or, when not "
                         "interactive, the suggested category is used and flagged)")
    ap.add_argument("--add-reference", metavar="FILE",
                    help="Add FILE (e.g. an ICAP illustrative set) to the reference library and exit")
    ap.add_argument("--ref-category", choices=["PIC", "LSC", "MSC", "SSC", "ANY"], default="ANY",
                    help="Category of the reference being added")
    ap.add_argument("--ref-name", help="Name for the reference being added")
    ap.add_argument("--list-references", action="store_true", help="List the reference library and exit")
    ap.add_argument("--framework", default="IFRS", help="IFRS, IFRS for SMEs, US GAAP, UK GAAP (FRS 102), ...")
    ap.add_argument("--jurisdiction")
    ap.add_argument("--entity-type", dest="entity_type")
    ap.add_argument("--year", dest="financial_year", help="Financial year, e.g. 2025")
    ap.add_argument("--currency")
    ap.add_argument("--unit", help="e.g. '000")
    ap.add_argument("--template", choices=sorted(PRESETS), default="big4")
    ap.add_argument("--page-size", dest="page_size", choices=["A4", "Letter"], default="A4")
    ap.add_argument("--negatives", dest="negative_style", choices=["brackets", "minus"], default="brackets")
    ap.add_argument("--fix-spelling", dest="apply_spelling", action="store_true",
                    help="Apply high-confidence spelling corrections (all listed in the report)")
    ap.add_argument("--no-page-breaks", dest="statement_page_breaks", action="store_false")
    ap.add_argument("--no-header-footer", dest="header_footer", action="store_false")
    ap.add_argument("--keep-borders", action="store_true", help="Keep existing table borders")
    ap.add_argument("--ai", dest="ai_review", action="store_true",
                    help="Also proofread narrative text with Claude (needs ANTHROPIC_API_KEY)")
    ap.add_argument("--outputs", default="docx,xlsx,pdf,report")
    args = ap.parse_args(argv)

    if args.add_reference:
        from . import library
        meta = library.add(args.add_reference, args.ref_name, "reference", args.ref_category)
        print(f"Added '{meta['name']}' ({meta['category']}) with {len(meta['note_headings'])} note headings.")
        return 0
    if args.list_references:
        from . import library
        for m in library.items():
            print(f"{m['id']}  {m['category']:4}  {m['name']}  ({len(m['note_headings'])} notes)")
        return 0
    if args.serve:
        from .server import serve
        serve(args.port)
        return 0
    if not args.input:
        ap.print_help()
        return 2
    if not os.path.exists(args.input):
        print(f"File not found: {args.input}", file=sys.stderr)
        return 2
    settings = {k: v for k, v in vars(args).items()
                if k not in ("input", "out", "serve", "port", "reference", "keep_borders", "outputs", "add_reference",
                             "ref_category", "ref_name", "list_references")}
    if not args.category and sys.stdin.isatty():
        from .entity import CATEGORIES
        from .pipeline import analyse
        a = analyse(args.input, settings)
        print("\nWhich set of financial statements is this?")
        for r in a["reasons"]:
            print("  - " + r)
        for code, c in CATEGORIES.items():
            mark = "  <- suggested" if code == a["suggested"] else ""
            print(f"  {code}  {c['name']:<24} {c['framework']}{mark}")
        ans = input(f"Category [{a['suggested']}]: ").strip().upper() or a["suggested"]
        while ans not in CATEGORIES:
            ans = input("Enter PIC, LSC, MSC or SSC: ").strip().upper()
        settings["category"] = ans
    settings["table_borders"] = "keep" if args.keep_borders else "financial"
    settings["outputs"] = [x.strip() for x in args.outputs.split(",") if x.strip()]
    out = args.out or os.path.splitext(args.input)[0] + " - review"
    from .pipeline import run
    data = run(args.input, out, settings, args.reference)
    s = data["summary"]
    print(f"\n{s['error']} error(s), {s['warning']} warning(s), {s['info']} note(s), {s['passed']} check(s) passed.")
    for i in data["issues"]:
        if i["status"] != "passed" and i["severity"] in ("error", "warning"):
            print(f"  [{i['severity'].upper():7}] {i['category']}: {i['message']}"
                  + (f"  ({i['location']})" if i["location"] else ""))
    print(f"\nFiles written to: {out}")
    for k, v in data["outputs"].items():
        print(f"  {k}: {v}")
    return 1 if s["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
