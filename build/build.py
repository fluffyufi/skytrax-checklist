"""Assemble the workbook: python build/build.py"""
import importlib
import json
import os
import subprocess
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
OUT_DIR = os.path.join(HERE, "..", "output")
OUT = os.path.join(OUT_DIR, "MAGCS_Skytrax_2026_Catering_Readiness.xlsx")
RECALC = "/root/.claude/skills/synced/dbf11ce3-ba00-4160-a445-f73a3231a942_7e2a8d9e-80e6-4124-9d8d-d105bdc3f685/xlsx/scripts/recalc.py"

ORDER = ["Instructions", "Dashboard", "Flights", "Checks", "Documents", "ISOP Register", "Requirements", "Settings"]


def protect(wb):
    """Lock formula cells (no password); yellow input cells stay editable. Filtering, sorting,
    row/column sizing and inserting pictures remain allowed."""
    from openpyxl.styles import Protection
    unlocked = Protection(locked=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                fg = c.fill.fgColor.rgb if c.fill and c.fill.fill_type == "solid" else None
                if isinstance(fg, str) and fg.upper().endswith("FFF2CC"):
                    c.protection = unlocked
        p = ws.protection
        p.sheet = True
        p.autoFilter = False
        p.sort = True
        p.formatRows = False
        p.formatColumns = False
        p.formatCells = False
        p.objects = False
        p.selectLockedCells = False
        p.selectUnlockedCells = False


def main():
    data = json.load(open(os.path.join(HERE, "data.json")))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for mod in ("core", "dashboard", "printable"):
        if os.path.exists(os.path.join(HERE, mod + ".py")):
            importlib.import_module(mod).build(wb, data)
        else:
            print("skip", mod)
    names = [ws.title for ws in wb.worksheets]
    order = [n for n in ORDER if n in names] + sorted(n for n in names if n not in ORDER)
    wb._sheets = [wb[n] for n in order]
    wb.active = order.index("Dashboard") if "Dashboard" in order else 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == order[wb.active_sheet_index if hasattr(wb, 'active_sheet_index') else 0]
    protect(wb)
    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT)
    if "--no-recalc" not in sys.argv:
        res = subprocess.run([sys.executable, RECALC, OUT, "900"], capture_output=True, text=True)
        print(res.stdout[-3000:], res.stderr[-2000:])
    print(OUT)


if __name__ == "__main__":
    main()
