"""Blank workbook checks: A350 tail counted as open clarification; Dashboard 'Checks row' points at the right row."""
import os, sys, openpyxl
src = os.environ.get("TEST_SRC", "/home/user/skytrax-checklist/output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx")
wb = openpyxl.load_workbook(src, data_only=True)
fl, ck, db = wb["Flights"], wb["Checks"], wb["Dashboard"]
for r in range(5, 27):
    print(fl[f"A{r}"].value, fl[f"L{r}"].value, repr(fl[f"N{r}"].value), "clarif", fl[f"AS{r}"].value) if fl[f"L{r}"].value == "A350" else None
bad = 0
for r in range(1, db.max_row + 1):
    v = db[f"T{r}"].value
    if isinstance(v, (int, float)) and v >= 5:
        cid = db[f"B{r}"].value if False else None
        rowvals = [db.cell(r, c).value for c in range(1, 20)]
        ids = [x for x in rowvals if isinstance(x, str) and x[:1] == "F" and "-" in x]
        if ids and ck[f"A{int(v)}"].value != ids[0]:
            bad += 1; print("mismatch", r, ids[0], v, ck[f"A{int(v)}"].value)
print("dashboard row mismatches:", bad)
