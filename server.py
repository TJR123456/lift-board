#!/usr/bin/env python3
"""Serves the Lift Board at http://localhost:8765 and powers its Refresh button.

POST /refresh downloads the New Lift Sheet straight from Google (it is shared
by link), checks it is a real workbook, keeps the old copy as a backup and
rebuilds index.html. Only listens on this Mac (127.0.0.1).

Run:  python3 server.py      (a LaunchAgent can keep it running; see README)
"""
import json
import shutil
import subprocess
import sys
import threading
import urllib.request
import zipfile
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE / "data" / "New Lift Sheet.xlsx"
# The sheet's ID (the long code in its URL) lives in sheet_id.txt, kept out of git.
SHEET_ID_FILE = HERE / "sheet_id.txt"
PORT = 8765
lock = threading.Lock()


def refresh():
    if not SHEET_ID_FILE.exists():
        raise RuntimeError("Put your Google Sheet's ID in sheet_id.txt next to server.py.")
    export = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_FILE.read_text().strip()}/export?format=xlsx"
    tmp = TARGET.with_suffix(".incoming.xlsx")
    with urllib.request.urlopen(export, timeout=60) as r:
        tmp.write_bytes(r.read())
    try:
        with zipfile.ZipFile(tmp) as z:
            if z.testzip() or "xl/workbook.xml" not in z.namelist():
                raise ValueError
    except (zipfile.BadZipFile, ValueError):
        tmp.unlink()
        raise RuntimeError("Google didn't send a workbook (is the sheet still shared by link?). Kept the old data.")
    shutil.copy2(TARGET, TARGET.with_suffix(".previous.xlsx"))
    tmp.replace(TARGET)
    out = subprocess.run([sys.executable, "-I", str(HERE / "build.py")], capture_output=True, text=True)
    if out.returncode:
        shutil.copy2(TARGET.with_suffix(".previous.xlsx"), TARGET)
        raise RuntimeError("Rebuild failed, kept the old data: " + out.stderr.strip().splitlines()[-1])
    return out.stdout.strip().splitlines()[-1]


class Handler(SimpleHTTPRequestHandler):
    def do_POST(self):
        origin = self.headers.get("Origin", "")
        if self.path != "/refresh" or (origin and origin not in (f"http://localhost:{PORT}", f"http://127.0.0.1:{PORT}")):
            return self.send_error(404)
        if not lock.acquire(blocking=False):
            return self.reply(409, ok=False, message="A refresh is already running.")
        try:
            self.reply(200, ok=True, message=refresh())
        except Exception as e:          # report anything back to the page
            self.reply(500, ok=False, message=str(e))
        finally:
            lock.release()

    def reply(self, code, **body):
        data = json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *args):
        pass


if __name__ == "__main__":
    ThreadingHTTPServer(("127.0.0.1", PORT), partial(Handler, directory=str(HERE))).serve_forever()
