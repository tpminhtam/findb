"""Google Sheets import for link-shared sheets: the export URL keeps formulas and needs no OAuth."""
from __future__ import annotations

import re
from pathlib import Path

import requests

ID_RE = re.compile(r"/spreadsheets/d/([A-Za-z0-9_-]+)")


def fetch_gsheet(url: str, dest_dir: Path) -> Path:
    m = ID_RE.search(url)
    if not m:
        raise ValueError("not a Google Sheets URL (expected /spreadsheets/d/<id>)")
    sid = m.group(1)
    export = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
    r = requests.get(export, timeout=60, allow_redirects=True)
    if r.status_code != 200 or not r.content.startswith(b"PK"):
        raise ValueError(f"could not export sheet {sid}: HTTP {r.status_code}. Is it shared with 'anyone with the link'?")
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / f"gsheet_{sid[:12]}.xlsx"
    path.write_bytes(r.content)
    return path
