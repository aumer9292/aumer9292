#!/bin/sh
# Opens FS Review in your browser (macOS / Linux). Needs Python 3.10+.
cd "$(dirname "$0")"
if [ ! -f .installed ]; then
  echo "First run: installing the components FS Review needs (once)."
  python3 -m pip install -r requirements.txt && echo done > .installed || exit 1
fi
python3 -m fsreview --serve
