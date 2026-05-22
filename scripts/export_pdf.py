"""LibreOffice headless PDF export of the Comparison tab.

Renders the Comparison tab of a built workbook to a single-page A4
landscape PDF suitable for sharing with clients.

    python scripts/export_pdf.py dist/Division_296_Model_v0.1.0.xlsx
    python scripts/export_pdf.py dist/Division_296_Model_v0.1.0.xlsx --tab Comparison --out my_output.pdf
    python scripts/export_pdf.py dist/Division_296_Model_v0.1.0.xlsx --all-tabs

Requires LibreOffice installed and `soffice` on PATH. On Windows the
default install puts it at:
    C:\\Program Files\\LibreOffice\\program\\soffice.exe

Approach: copy the workbook to a temp dir, mark only the requested tab
as visible (others hidden), then run `soffice --convert-to pdf`. The
Comparison tab's print area + fit-to-page settings ensure it lands on
one A4 landscape page.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook


def _soffice_executable() -> str | None:
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    # Common Windows install path.
    win_default = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    if win_default.exists():
        return str(win_default)
    return None


def _isolate_tab(src: Path, work_dir: Path, tab_name: str | None) -> Path:
    """Copy `src` to `work_dir` and (optionally) hide every tab except `tab_name`."""
    staged = work_dir / src.name
    shutil.copy2(src, staged)
    if tab_name is None:
        return staged

    wb = load_workbook(staged)
    if tab_name not in wb.sheetnames:
        raise SystemExit(
            f"Tab {tab_name!r} not found in {src.name}. Available: {wb.sheetnames}"
        )
    wb.active = wb.sheetnames.index(tab_name)
    for name in wb.sheetnames:
        wb[name].sheet_state = "visible" if name == tab_name else "hidden"
    wb.save(staged)
    return staged


def export_pdf(src: Path, out: Path | None, tab: str | None) -> int:
    soffice = _soffice_executable()
    if soffice is None:
        print("ERROR: LibreOffice not found. Install it and put soffice on PATH.",
              file=sys.stderr)
        return 1

    if not src.exists():
        print(f"ERROR: input file does not exist: {src}", file=sys.stderr)
        return 1

    out_dir = (out.parent if out else src.parent).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="div296_pdf_") as td:
        work = Path(td)
        staged = _isolate_tab(src, work, tab)

        cmd = [
            soffice, "--headless", "--calc",
            "--convert-to", "pdf",
            "--outdir", str(work),
            str(staged),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            print(result.stdout, file=sys.stdout)
            print(result.stderr, file=sys.stderr)
            return result.returncode

        produced = work / (staged.stem + ".pdf")
        if not produced.exists():
            print(f"ERROR: LibreOffice did not produce {produced}", file=sys.stderr)
            print(result.stdout, file=sys.stdout)
            return 1

        final = out or (out_dir / f"{src.stem}{'_' + tab if tab else ''}.pdf")
        shutil.copy2(produced, final)
        print(f"Wrote {final}")
        return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Render a tab of the Div 296 workbook to PDF via LibreOffice headless.")
    parser.add_argument("xlsx", type=Path, help="Path to the built .xlsx")
    parser.add_argument(
        "--tab", default="Comparison",
        help="Tab to export (default: Comparison). Use --all-tabs to export every tab.",
    )
    parser.add_argument(
        "--all-tabs", action="store_true",
        help="Export the whole workbook (every visible tab) into one PDF.",
    )
    parser.add_argument(
        "--out", type=Path, default=None,
        help="Output PDF path (default: alongside the .xlsx).",
    )
    args = parser.parse_args(argv)

    tab = None if args.all_tabs else args.tab
    return export_pdf(args.xlsx.resolve(), args.out, tab)


if __name__ == "__main__":
    raise SystemExit(main())
