from __future__ import annotations

import cgi
import json
import sqlite3
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class _FastApiStub:
    def __init__(self, *args, **kwargs): pass
    def add_middleware(self, *args, **kwargs): pass
    def on_event(self, *args, **kwargs): return lambda fn: fn
    def post(self, *args, **kwargs): return lambda fn: fn
    def get(self, *args, **kwargs): return lambda fn: fn
    def mount(self, *args, **kwargs): pass


class _HTTPException(Exception):
    def __init__(self, status_code: int, detail: str):
        self.status_code = status_code
        self.detail = detail


class _UploadFile:
    def __init__(self, filename, file):
        self.filename = filename
        self.file = file

class _NoopInit:
    def __init__(self, *args, **kwargs): pass


sys.modules.setdefault("fastapi", type(sys)("fastapi"))
sys.modules["fastapi"].FastAPI = _FastApiStub
sys.modules["fastapi"].File = lambda *a, **k: None
sys.modules["fastapi"].Form = lambda *a, **k: None
sys.modules["fastapi"].HTTPException = _HTTPException
sys.modules["fastapi"].UploadFile = _UploadFile
sys.modules.setdefault("fastapi.middleware", type(sys)("fastapi.middleware"))
sys.modules.setdefault("fastapi.middleware.cors", type(sys)("fastapi.middleware.cors"))
sys.modules["fastapi.middleware.cors"].CORSMiddleware = object
sys.modules.setdefault("fastapi.responses", type(sys)("fastapi.responses"))
sys.modules["fastapi.responses"].FileResponse = _NoopInit
sys.modules.setdefault("fastapi.staticfiles", type(sys)("fastapi.staticfiles"))
sys.modules["fastapi.staticfiles"].StaticFiles = _NoopInit

from backend.app.main import (  # noqa: E402
    checks,
    dashboard,
    import_price_plan,
    import_price_statistics,
    preview_price_plan,
    preview_price_statistics,
    reminders,
    startup,
)


def form_value(form, key, default=None):
    if key not in form:
        return default
    item = form[key]
    if isinstance(item, list):
        item = item[0]
    return item.value


def form_upload(form):
    item = form["file"]
    if isinstance(item, list):
        item = item[0]
    return _UploadFile(item.filename, item.file)


class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/dashboard":
                return self.send_json(dashboard())
            if parsed.path == "/api/checks":
                params = parse_qs(parsed.query)
                return self.send_json(checks(params.get("issue_type", [None])[0], params.get("sku", [None])[0], params.get("platform", [None])[0]))
            if parsed.path == "/api/reminders":
                return self.send_json(reminders())
            if parsed.path in {"/", "/index.html"}:
                return self.send_file(ROOT / "frontend" / "index.html", "text/html; charset=utf-8")
            target = (ROOT / "frontend" / parsed.path.lstrip("/")).resolve()
            if target.is_file() and ROOT in target.parents:
                return self.send_file(target, "application/octet-stream")
            self.send_json({"detail": "Not found"}, 404)
        except Exception as exc:
            self.send_json({"detail": str(exc)}, 500)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type")})
            if parsed.path == "/api/price-statistics/preview":
                return self.send_json(preview_price_statistics(form_upload(form), form_value(form, "snapshot_date")))
            if parsed.path == "/api/price-statistics/import":
                return self.send_json(import_price_statistics(form_upload(form), form_value(form, "snapshot_date"), form_value(form, "platform_columns", "")))
            if parsed.path == "/api/price-plan/preview":
                return self.send_json(preview_price_plan(form_upload(form), int(form_value(form, "plan_year"))))
            if parsed.path == "/api/price-plan/import":
                return self.send_json(import_price_plan(form_upload(form), int(form_value(form, "plan_year")), form_value(form, "stage_columns", "")))
            self.send_json({"detail": "Not found"}, 404)
        except _HTTPException as exc:
            self.send_json({"detail": exc.detail}, exc.status_code)
        except sqlite3.Error as exc:
            self.send_json({"detail": f"database error: {exc}"}, 500)
        except Exception as exc:
            self.send_json({"detail": str(exc)}, 500)

    def send_file(self, path: Path, content_type: str):
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    startup()
    host, port = "127.0.0.1", 8000
    print(f"Serving on http://{host}:{port}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    main()
