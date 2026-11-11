#!/usr/bin/env python3
"""Swap in a fresh export of the New Lift Sheet, then rebuild index.html.

Usage: python3 refresh.py <path to the Drive download result>

The path is the JSON file the Google Drive connector saves when it downloads
the sheet as .xlsx ({"content": "<base64>", ...}). The old copy is kept as a
backup and only replaced once the new file opens as a real workbook.
"""
import base64
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE / "data" / "New Lift Sheet.xlsx"

raw = json.loads(Path(sys.argv[1]).read_text())
blob = base64.b64decode(raw["content"])
tmp = TARGET.with_suffix(".incoming.xlsx")
tmp.write_bytes(blob)
with zipfile.ZipFile(tmp) as z:
    bad = z.testzip()
    if bad or "xl/workbook.xml" not in z.namelist():
        tmp.unlink()
        sys.exit(f"Download is not a valid workbook ({bad or 'no xl/workbook.xml'}); kept the old copy.")
shutil.copy2(TARGET, TARGET.with_suffix(".previous.xlsx"))
tmp.replace(TARGET)
subprocess.run([sys.executable, "-I", str(HERE / "build.py")], check=True)
