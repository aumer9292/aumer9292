"""Local browser interface. Runs on this computer only (127.0.0.1); files never leave it
unless the user turns on AI review."""

import email.parser
import email.policy
import json
import os
import re
import tempfile
import threading
import traceback
import uuid
import webbrowser
from urllib.parse import quote, unquote
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import __version__
from .ai_review import available as ai_available
from .pdf_engine import find_office
from .templates import DEFAULTS, PRESETS

WEB = os.path.join(os.path.dirname(__file__), "web")
JOBS = os.path.join(tempfile.gettempdir(), "fsreview-jobs")
ALLOWED = (".docx", ".docm", ".xlsx", ".xlsm", ".xls", ".pdf", ".txt")
MAX_UPLOAD = 100 * 1024 * 1024
_lock = threading.Lock()


class Handler(BaseHTTPRequestHandler):
    server_version = f"fsreview/{__version__}"

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(os.path.join(WEB, "index.html"), "rb") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if self.path == "/api/options":
            return self._send(200, {
                "version": __version__,
                "templates": {k: v["label"] for k, v in PRESETS.items()},
                "defaults": {k: v for k, v in DEFAULTS.items()},
                "ai_available": ai_available(),
                "office": bool(find_office()),
            })
        m = re.fullmatch(r"/api/file/([0-9a-f]{32})/([^/]+)", self.path)
        if m:
            job, name = m.group(1), os.path.basename(unquote(m.group(2)))
            path = os.path.join(JOBS, job, "out", name)
            if not os.path.isfile(path):
                return self._send(404, {"error": "File not found"})
            ctype = {".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
                     ".pdf": "application/pdf", ".json": "application/json"}.get(
                os.path.splitext(name)[1].lower(), "application/octet-stream")
            with open(path, "rb") as f:
                data = f.read()
            return self._send(200, data, ctype, {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"})
        return self._send(404, {"error": "Not found"})

    def do_POST(self):
        if self.path != "/api/review":
            return self._send(404, {"error": "Not found"})
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0 or length > MAX_UPLOAD:
            return self._send(413, {"error": "Upload is empty or larger than 100 MB."})
        ctype = self.headers.get("Content-Type", "")
        if not ctype.startswith("multipart/form-data"):
            return self._send(400, {"error": "Expected a form upload."})
        body = self.rfile.read(length)
        msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(
            b"Content-Type: " + ctype.encode() + b"\r\n\r\n" + body)
        fields, files = {}, {}
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            fname = part.get_filename()
            if fname:
                files[name] = (os.path.basename(fname), part.get_payload(decode=True) or b"")
            else:
                fields[name] = part.get_content().strip() if part.get_content_maintype() == "text" else ""
        if "file" not in files or not files["file"][1]:
            return self._send(400, {"error": "Choose the financial statements file to review."})
        fname, data = files["file"]
        if not fname.lower().endswith(ALLOWED):
            return self._send(400, {"error": f"'{fname}' is not a supported file. Use Word, Excel, PDF or text."})
        job = uuid.uuid4().hex
        jdir = os.path.join(JOBS, job)
        os.makedirs(os.path.join(jdir, "out"), exist_ok=True)
        src = os.path.join(jdir, fname)
        with open(src, "wb") as f:
            f.write(data)
        ref = None
        if "reference" in files and files["reference"][1]:
            rname, rdata = files["reference"]
            if rname.lower().endswith(".docx"):
                ref = os.path.join(jdir, "reference-" + rname)
                with open(ref, "wb") as f:
                    f.write(rdata)
        try:
            settings = json.loads(fields.get("settings") or "{}")
        except json.JSONDecodeError:
            settings = {}
        from .pipeline import run
        try:
            with _lock:     # LibreOffice runs one conversion at a time
                result = run(src, os.path.join(jdir, "out"), settings, ref, progress=lambda m: None)
        except Exception as e:
            traceback.print_exc()
            return self._send(500, {"error": f"The document could not be processed: {e}"})
        result["job"] = job
        return self._send(200, result)


def serve(port=8765, open_browser=True):
    os.makedirs(JOBS, exist_ok=True)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"FS Review is running at {url}  (press Ctrl+C to stop)")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
