"""Fail if anything that looks like real client/employer data or a secret is in the repo.

Runs locally (`python scripts/check_no_real_data.py`) and in GitHub Actions.
Edit FORBIDDEN to add names that must never appear (employer, vendors, clients).
Also scans text inside .xlsx files.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKIP_DIRS = {".git", ".venv", "__pycache__", "node_modules"}
SKIP_FILES = {"check_no_real_data.py"}

# Case-sensitive where noted; keep this list private-safe (no real client names in a public repo!).
FORBIDDEN = [
    (r"\bAVi\b", "employer name"),
    (r"(?i)aroflo", "vendor name"),
    (r"(?i)securecorp|securitas|\bESS\b", "monitoring company"),
    (r"(?i)stripe|xero", "billing vendor"),
    (r"(?i)api[_-]?key\s*[:=]\s*['\"]?\w{8,}", "API key"),
    (r"(?i)(secret|token|password)\s*[:=]\s*['\"]\w{6,}", "credential"),
    (r"\b(?:\+?61|0)4\d{2}[ -]?\d{3}[ -]?\d{3}\b", "AU mobile number"),
    (r"[\w.+-]+@(?!example\.)[\w-]+\.[\w.]+", "email address"),
]
TEXT_EXT = {".py", ".html", ".md", ".txt", ".json", ".yml", ".yaml", ".bat", ".css", ".js", ".csv"}


def texts():
    for p in ROOT.rglob("*"):
        if not p.is_file() or SKIP_DIRS & set(p.parts) or p.name in SKIP_FILES:
            continue
        if p.suffix in TEXT_EXT:
            yield p, p.read_text(errors="ignore")
        elif p.suffix == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(p, read_only=True, data_only=True)
            yield p, "\n".join(
                " ".join(str(c) for c in row if c is not None)
                for ws in wb.worksheets for row in ws.iter_rows(values_only=True))


bad = []
for path, text in texts():
    for pat, why in FORBIDDEN:
        for m in re.finditer(pat, text):
            line = text.count("\n", 0, m.start()) + 1
            bad.append(f"{path.relative_to(ROOT)}:{line}: {why}: {m.group(0)[:40]!r}")
if bad:
    print("POSSIBLE REAL DATA FOUND:\n" + "\n".join(bad))
    sys.exit(1)
print("clean: no forbidden strings found")
