"""Regression tests for logic-critic round 2 findings."""
import shutil, subprocess, sys
from datetime import datetime, timedelta
import openpyxl
sys.path.insert(0, "/home/user/skytrax-checklist/build")
import test_ready as T

F11_STD = datetime(2026, 10, 12, 15, 50)  # PEN local


def fill_f11(wb):
    T.STD = F11_STD
    T.fill(wb["Checks"], wb["Documents"], fid="F11")
    ck = wb["Checks"]
    for r in range(5, ck.max_row + 1):
        if ck[f"B{r}"].value != "F11":
            continue
        cid = ck[f"A{r}"].value
        if cid == "F11-T7-10":  # carrying flight clarification: must be Pass (N/A not allowed)
            ck[f"U{r}"] = "Pass"; ck[f"AC{r}"] = None
        if cid == "F11-UPL-03":  # KUL loading line: before MH1140 dep 11:45 KUL
            ck[f"AD{r}"] = datetime(2026, 10, 12, 10, 0)
        if ck[f"G{r}"].value in ("T-24H", "T-12H PREP"):  # capped by loading (AY = 12-Oct 05:45 KUL)
            ck[f"AD{r}"] = datetime(2026, 10, 11, 23, 0) if ck[f"G{r}"].value == "T-12H PREP" else datetime(2026, 10, 11, 10, 0)
        if ck[f"H{r}"].value == "Physical uplift" and cid != "F11-UPL-03":
            ck[f"AD{r}"] = datetime(2026, 10, 12, 14, 0)  # PEN, after MH1140 arrival
    d = wb["Documents"]
    d["F15"], d["G15"], d["H15"], d["I15"] = "GLD-1", "R1", datetime(2026, 9, 1), "loc"
    d["K15"], d["L15"], d["M15"], d["N15"] = "MCL-1", "R1", datetime(2026, 9, 1), "loc"
    wb["Flights"]["AX15"] = datetime(2026, 10, 12, 3, 45)


def run(tag, mutate):
    p = T.TMP.format("l2_" + tag)
    shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p)
    wb["Settings"]["B4"] = datetime(2026, 10, 12, 7, 49)
    fill_f11(wb)
    rows = {wb["Checks"].cell(r, 1).value: r for r in range(5, wb["Checks"].max_row + 1)}
    if mutate:
        mutate(wb, rows)
    wb.save(p)
    out = subprocess.run([sys.executable, T.RECALC, p, "300"], capture_output=True, text=True).stdout
    wb = openpyxl.load_workbook(p, data_only=True)
    ck = wb["Checks"]
    bad = [(ck[f"A{r}"].value, ck[f"AH{r}"].value) for r in range(5, ck.max_row + 1)
           if ck[f"B{r}"].value == "F11" and not str(ck[f"AH{r}"].value).startswith(("COMPLETE", "N/A"))]
    ok = '"total_errors": 0' in out
    print(f"{tag:26s} {wb['Flights']['AT15'].value!r:32s} ok={ok} bad={bad[:3]}")


def cell(cid, col, v):
    return lambda wb, rows: wb["Checks"].__setitem__(f"{col}{rows[cid]}", v)


if __name__ == "__main__" and len(sys.argv) == 1:
  run("baseline_F11", None)
  run("onboard_before_carry", cell("F11-UPL-01", "AD", datetime(2026, 10, 12, 10, 0)))
  run("ca_open_space", lambda wb, rows: [wb["Checks"].__setitem__(f"AA{rows['F11-T7-01']}", "Re-plate"),
                                         wb["Checks"].__setitem__(f"AB{rows['F11-T7-01']}", "Open ")])
  run("carry_clar_NA", lambda wb, rows: [wb["Checks"].__setitem__(f"U{rows['F11-T7-10']}", "N/A"),
                                         wb["Checks"].__setitem__(f"AC{rows['F11-T7-10']}", "n/a")])
  run("ax_after_std", lambda wb, rows: wb["Flights"].__setitem__("AX15", datetime(2026, 10, 12, 9, 0)))
  run("result_blank", cell("F11-T24-01", "V", None))
  run("t7_stale", cell("F11-T7-01", "AD", datetime(2026, 7, 2, 10, 0)))
  run("pic_is_verifier", cell("F11-T7-01", "AE", "A. Rahman"))
  run("ca_text_no_status", cell("F11-T7-02", "AA", "Reprinted cards"))


def more():
    run("placeholder_evidence", cell("F11-T7-01", "Z", "TBC"))
    run("placeholder_verifier", cell("F11-T7-01", "AE", "?"))
    run("prep_after_carrier", cell("F11-T12-01", "AD", datetime(2026, 10, 12, 12, 0)))
    run("onboard_before_arrival", cell("F11-UPL-01", "AD", datetime(2026, 10, 12, 12, 30)))
    run("pass_after_ca_short", lambda wb, rows: [wb["Checks"].__setitem__(f"{k}{rows['F11-UPL-01']}", v) for k, v in
                                                 (("X", 20), ("Y", 15), ("U", "Pass after CA"), ("AA", "Topped up 5 meals from spare"), ("AB", "Closed"))])
    run("na_short_just", lambda wb, rows: [wb["Checks"].__setitem__(f"U{rows['F11-T7-02']}", "N/A"),
                                           wb["Checks"].__setitem__(f"AC{rows['F11-T7-02']}", "x")])
    run("uplift_window_24", lambda wb, rows: wb["Settings"].__setitem__("B7", 24))


if __name__ == "__main__" and len(sys.argv) > 1:
    more()
