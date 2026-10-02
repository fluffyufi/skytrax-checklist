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
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "900"], capture_output=True)
    ck = openpyxl.load_workbook(p, data_only=True)["Checks"]
    for cid in ids:
        print(f"{tag:26s} {cid:12s} -> {ck['AH' + str(rows[cid])].value[:110]}")


def na(ck, r, when, outcome="Confirmed – not carried on this sector"):
    ck[f"U{r}"] = "N/A"; ck[f"AC{r}"] = outcome
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
        import test_logic2 as L
        L.fill_f11(wb)
        std = datetime(2026, 10, 12, 15, 50)
        na(ck, rows["F11-T7-02"], when_fn(std), "Refreshment service – no printed menu card")
        return ["F11-T7-02"]
    return prep


def tailqty(tail, x):
    def prep(wb, ck, rows):
        wb["Flights"]["N5"] = tail
        r = rows["F01-T12-09"]  # Blanket (EY)
        std = wb["Flights"]["I5"].value
        for k, v in (("U", "Pass"), ("T", "A. Rahman"), ("AE", "N. Ismail"), ("V", "Blankets counted against GLD, all bundled"),
                     ("X", x), ("Y", x), ("AD", std - timedelta(hours=13))):
            ck[f"{k}{r}"] = v
        return ["F01-T12-09"]
    return prep


def r22_gld_revised(wb, ck, rows):
    T.STD = datetime(2026, 10, 9, 21, 50); T.fill(ck, wb["Documents"])
    wb["Documents"]["H6"] = datetime(2026, 10, 9, 12, 0)  # revised after the T-12H prep checks, before STD
    return ["F02-T12-03", "F02-T12-07", "F02-UPL-02"]


def r22_kul_prep_after_loading(wb, ck, rows):
    import test_logic2 as L
    L.fill_f11(wb)
    ck[f"AD{rows['F11-T12-01']}"] = datetime(2026, 10, 12, 10, 30)  # KUL, after the KUL loading line (10:00)
    return ["F11-T12-01", "F11-UPL-03"]


def r22_f05_not_carried_no_ax(wb, ck, rows):
    T.STD = wb["Flights"]["I9"].value; T.fill(ck, wb["Documents"], fid="F05")
    r = rows["F05-UPL-03"]
    na(ck, r, T.STD - timedelta(days=2))
    return ["F05-T7-10", "F05-UPL-03", "F05-UPL-01", "F05-UPL-02"]


def r22_onboard_clar_photo(wb, ck, rows):
    T.STD = datetime(2026, 10, 9, 21, 50); T.fill(ck, wb["Documents"])
    for c in ("F02-T12-12", "F02-UPL-07"):
        r = rows[c]; ck[f"U{r}"] = "Pass"; ck[f"AC{r}"] = "Confirmed – applies / carried as listed"
        ck[f"V{r}"] = "Table cloth confirmed and counted"
    return ["F02-T12-12", "F02-UPL-07"]


def r22_asof_review(wb, ck, rows):
    T.STD = datetime(2026, 10, 9, 21, 50); T.fill(ck, wb["Documents"])
    wb["Settings"]["B4"] = datetime(2026, 10, 9, 5, 0)  # before the T-12H prep (08-Oct 20:50 local = 12:50Z) ... and uplift
    return ["F02-T12-01", "F02-UPL-01"]


def passline(ck, r, when, result="Confirmed with MAGCS, item prepared and counted"):
    for k, v in (("U", "Pass"), ("AC", "Confirmed – applies / carried as listed"),
                 ("T", "A. Rahman"), ("AE", "N. Ismail"), ("V", result), ("AD", when)):
        ck[f"{k}{r}"] = v


def sq_case(conf, carrier, prep_local, upl_local=None, flight="F06"):
    def prep(wb, ck, rows):
        st = wb["Settings"]
        r = next(r for r in range(27, 40) if str(st[f"A{r}"].value or "").split("  ")[-1].startswith(flight))
        st[f"F{r}"] = conf; st[f"G{r}"] = carrier
        pid = f"{flight}-T12-08" if flight == "F06" else f"{flight}-T12-10"
        uid = f"{flight}-UPL-03" if flight == "F06" else f"{flight}-UPL-05"
        passline(ck, rows[pid], prep_local)
        ids = [pid]
        if upl_local:
            passline(ck, rows[uid], upl_local, "Toiletry kits counted on board, all to GLD"); ids.append(uid)
        return ids
    return prep


if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "r22":
    states("r22_kul_prep_after_loading", r22_kul_prep_after_loading)
    states("r22_asof_review", r22_asof_review)

if __name__ == "__main__" and len(sys.argv) == 1:
    states("carry_na_item_open", f05(False))
    states("na_after_departure", timing(lambda std: std + timedelta(days=2)))
    states("na_no_time", timing(lambda std: None))
    states("na_before_departure", timing(lambda std: std - timedelta(days=2)))
    states("tail_mad_260", tailqty("9M-MAD", 260))  # A359 blanket qty (all Skytrax A350 tails are A359)
