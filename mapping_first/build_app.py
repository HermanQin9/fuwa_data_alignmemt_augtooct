"""Build the Windows app: dist/PMIX App/ holds the .exe and everything it reads.

Run from the project root with an interpreter that has pandas, openpyxl, XlsxWriter, pywin32
and PyInstaller (a clean virtual environment keeps the .exe small):

    python mapping_first/build_app.py

The folder is never committed (dist/ is ignored): it holds store sales figures in the saved
mapping. The reference workbook is a copy of the 'Menu Category' outline only; the originals
(USA PMix.xlsx, Menu Items List.xlsx) stay untouched.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DIST = ROOT / "dist"
APP = DIST / "PMIX App"
NAME = "PMIX Report Builder"
DECISION_FILES = ["mapping_overrides.json", "menu_name_overrides.json",
                  "review_decisions.json", "menu_list_additions.json"]


def make_reference(source: Path, target: Path) -> None:
    """Copy only the outline (column E of 'Menu Category', with its indent) to a small workbook."""
    book = openpyxl.load_workbook(source, read_only=True, data_only=True)
    out = openpyxl.Workbook()
    sheet = out.active
    sheet.title = "Menu Category"
    try:
        for row in book["Menu Category"].iter_rows():
            if len(row) < 5 or row[4].value is None:
                continue
            cell = row[4]
            copy = sheet.cell(row=cell.row, column=5, value=cell.value)
            copy.alignment = Alignment(indent=int(cell.alignment.indent or 0))
    finally:
        book.close()
    out.save(target)


def main() -> None:
    work = ROOT / "build"
    exe_dir = DIST / "_exe"
    for folder in (work, exe_dir, APP):
        shutil.rmtree(folder, ignore_errors=True)
    APP.mkdir(parents=True)
    subprocess.run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed",
        "--name", NAME, "--paths", str(HERE), "--distpath", str(exe_dir),
        "--workpath", str(work), "--specpath", str(work),
        "--hidden-import", "pmix_workflow", "--hidden-import", "validate_prepared",
        "--hidden-import", "verify_result", "--hidden-import", "win32com.client",
        "--hidden-import", "xlsxwriter", "--hidden-import", "tkinter",
        str(HERE / "pmix_app.py"),
    ], check=True, cwd=ROOT)
    shutil.copy2(exe_dir / f"{NAME}.exe", APP / f"{NAME}.exe")
    shutil.copy2(ROOT / "Menu Items List.xlsx", APP / "Menu Items List.xlsx")
    make_reference(ROOT / "USA PMix.xlsx", APP / "Menu Category reference.xlsx")
    (APP / "decisions").mkdir()
    for name in DECISION_FILES:
        shutil.copy2(HERE / name, APP / "decisions" / name)
    (APP / "Result").mkdir()
    mapping = ROOT / "Result" / "USA PMix - Menu Mapping.xlsx"
    if mapping.exists():
        shutil.copy2(mapping, APP / "Result" / mapping.name)
    shutil.copy2(HERE / "APP_GUIDE.txt", APP / "README.txt")
    print(f"Built {APP}")


if __name__ == "__main__":
    main()
