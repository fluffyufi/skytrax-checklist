"""Scan every formula: length <= 8192 and function nesting depth <= 64 (Excel limits)."""
import re, sys
import openpyxl

path = sys.argv[1] if len(sys.argv) > 1 else "/home/user/skytrax-checklist/output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx"
wb = openpyxl.load_workbook(path)
worst_len, worst_depth = (0, None), (0, None)
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            f = c.value
            if not (isinstance(f, str) and f.startswith("=")):
                continue
            s = re.sub(r'"[^"]*"', '""', f)          # drop string literals
            depth = cur = 0
            for m in re.finditer(r"[A-Za-z_][A-Za-z0-9_.]*\(|\(|\)", s):
                t = m.group(0)
                if t == ")":
                    cur -= 1
                elif t.endswith("(") and len(t) > 1:
                    cur += 1
                    depth = max(depth, cur)
                else:
                    cur += 1  # plain parenthesis: not a function level, balance only
                    cur -= 0
            if len(f) > worst_len[0]:
                worst_len = (len(f), f"{ws.title}!{c.coordinate}")
            if depth > worst_depth[0]:
                worst_depth = (depth, f"{ws.title}!{c.coordinate}")
print("longest formula", worst_len, "| deepest function nesting", worst_depth)
