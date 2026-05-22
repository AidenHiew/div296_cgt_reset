"""LibreOffice headless recalc helper.

openpyxl writes formula text but not cached results. This script
opens the workbook in LibreOffice (if installed), forces a recalc,
and resaves. If LibreOffice is not on PATH, the script exits 0 with
a hint — Excel will recalc automatically the first time the user
opens the file (per spec §13).

    python scripts/recalc.py dist/Division_296_Model_v0.1.0.xlsx
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def _soffice_executable() -> str | None:
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return None


def recalc(path: Path) -> int:
    soffice = _soffice_executable()
    if soffice is None:
        print(
            "LibreOffice not found on PATH. Skipping recalc — "
            "Excel will recalculate on first open.",
            file=sys.stderr,
        )
        return 0

    cmd = [
        soffice, "--headless", "--calc",
        "--convert-to", "xlsx",
        "--outdir", str(path.parent),
        str(path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        print(result.stdout, file=sys.stdout)
        print(result.stderr, file=sys.stderr)
        return result.returncode
    print(f"Recalculated {path}")
    return 0


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print(f"Usage: {argv[0]} <path-to-xlsx>", file=sys.stderr)
        return 2
    return recalc(Path(argv[1]))


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
