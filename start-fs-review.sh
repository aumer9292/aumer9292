#!/bin/sh
# Opens FS Review in your browser (macOS / Linux). Needs Python 3.10+.
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt
python3 -m fsreview --serve
