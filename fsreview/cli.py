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
    ap.add_argument("--framework", default="IFRS", help="IFRS, IFRS for SMEs, US GAAP, UK GAAP (FRS 102), ...")
    ap.add_argument("--jurisdiction")
    ap.add_argument("--entity-type", dest="entity_type")
    ap.add_argument("--year", dest="financial_year", help="Financial year, e.g. 2025")
    ap.add_argument("--currency")
    ap.add_argument("--unit", help="e.g. '000")
    ap.add_argument("--template", choices=sorted(PRESETS), default="classic")
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
                if k not in ("input", "out", "serve", "port", "reference", "keep_borders", "outputs")}
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
