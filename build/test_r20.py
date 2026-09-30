"""Round-20 cases: N/A timing, KUL-sourced follow-on N/A (F05), A350 tail-specific quantity (F01)."""
import os, sys, shutil, subprocess
from datetime import datetime, timedelta
import openpyxl
sys.path.insert(0, "/home/user/skytrax-checklist/build")
import test_ready as T

F05_STD = datetime(2026, 10, 14, 10, 30)  # placeholder, replaced from Flights I9


def states(tag, prep):
    p = T.TMP.format("r20_" + tag); shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p); wb["Settings"]["B4"] = datetime(2026, 11, 2)
    ck = wb["Checks"]; rows = {ck.cell(r, 1).value: r for r in range(5, ck.max_row + 1)}
    ids = prep(wb, ck, rows)
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "300"], capture_output=True)
    ck = openpyxl.load_workbook(p, data_only=True)["Checks"]
    for cid in ids:
        print(f"{tag:26s} {cid:12s} -> {ck['AH' + str(rows[cid])].value[:110]}")


def na(ck, r, when, outcome="Confirmed – not carried on this sector"):
    ck[f"U{r}"] = "N/A"; ck[f"AC{r}"] = outcome; ck[f"S{r}"] = "Email"; ck[f"Z{r}"] = f"EM-26-{1000 + r}"
    ck[f"T{r}"] = "A. Rahman"; ck[f"AE{r}"] = "N. Ismail"; ck[f"AD{r}"] = when


def f05(all_na):
    def prep(wb, ck, rows):
        std = wb["Flights"]["I9"].value
        for cid in ("F05-T12-08", "F05-UPL-04", "F05-T7-10", "F05-UPL-03"):
            na(ck, rows[cid], std - timedelta(days=2))
        if not all_na:
            r = rows["F05-UPL-04"]; ck[f"U{r}"] = None
        return ["F05-T12-08", "F05-UPL-04", "F05-T7-10", "F05-UPL-03"]
    return prep


def timing(when_fn):
    def prep(wb, ck, rows):
        std = wb["Flights"]["I9"].value
        na(ck, rows["F05-T12-08"], when_fn(std))
        return ["F05-T12-08"]
    return prep


def tailqty(tail, x):
    def prep(wb, ck, rows):
        wb["Flights"]["N5"] = tail
        r = rows["F01-T12-10"]
        std = wb["Flights"]["I5"].value
        for k, v in (("U", "Pass"), ("T", "A. Rahman"), ("AE", "N. Ismail"), ("V", "Blankets counted against GLD, all bundled"),
                     ("S", "Form / checklist"), ("Z", "QF-2210"), ("X", x), ("Y", x), ("AD", std - timedelta(hours=13))):
            ck[f"{k}{r}"] = v
        return ["F01-T12-10"]
    return prep


if __name__ == "__main__":
    states("carry_na_all_items_na", f05(True))
    states("carry_na_item_open", f05(False))
    states("na_after_departure", timing(lambda std: std + timedelta(days=2)))
    states("na_no_time", timing(lambda std: None))
    states("na_before_departure", timing(lambda std: std - timedelta(days=2)))
    states("tail_mac_280", tailqty("9M-MAC", 280))
    states("tail_mac_260", tailqty("9M-MAC", 260))
    states("tail_mah_280", tailqty("9M-MAH", 280))
