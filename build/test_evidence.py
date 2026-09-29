"""Structured evidence (type in S + ID in Z): positive and negative cases on F11."""
import sys, shutil, subprocess
from datetime import datetime
import openpyxl
sys.argv = [sys.argv[0]]
sys.path.insert(0, "/home/user/skytrax-checklist/build")
import test_logic2 as L, test_ready as T

GOOD = [("Form / checklist", "SF-2210"), ("Photo", "IMG_2231"), ("Photo", "2231-2236"), ("Seal no.", "88213-88220"),
        ("Delivery note / receipt", "DN88213"), ("Email", "PASB-4471"), ("Load / uplift sheet", "PCS-0727"),
        ("Document (GLD / menu / ISOP)", "GLD-MH0727-A"), ("Link / file path", "\\\\fs01\\QA\\F11\\SF-2210.pdf"),
        ("System record", "QP-5521"), ("Memo / letter", "QA/26/118"), ("Form / checklist", "QF-44")]
BAD = [("Form / checklist", "Verified by CSM 1400"), ("Form / checklist", "Checked 10/10/26"), ("Photo", "1415"),
       ("Form / checklist", "MH1149"), ("Email", "QA/QC"), ("Link / file path", "SF-2210"), ("", "SF-2210"),
       ("Form / checklist", "9OCT26"), ("Seal no.", "Rev3"), ("Form / checklist", "SF-2210 to follow"),
       ("Photo", "MH88"), ("Form / checklist", "QA form QF-1234")]
IDS = ["T7-01", "T7-02", "T7-04", "T7-05", "T7-06", "T12-01", "T12-02", "T12-03", "T12-04", "T12-05", "T12-06", "T12-07"]


def run(tag, cases, batch=None):
    p = T.TMP.format("ev_" + tag); shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p); wb["Settings"]["B4"] = datetime(2026, 10, 12, 7, 49); L.fill_f11(wb)
    ck = wb["Checks"]; rows = {ck.cell(r, 1).value: r for r in range(5, ck.max_row + 1)}
    for sid, (typ, idv) in zip(IDS, cases):
        r = rows["F11-" + sid]; ck[f"S{r}"] = typ or None; ck[f"Z{r}"] = idv
    if batch:
        ck[f"W{rows['F11-T24-01']}"] = batch
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "300"], capture_output=True)
    wb = openpyxl.load_workbook(p, data_only=True); ck = wb["Checks"]
    for sid, (typ, idv) in zip(IDS, cases):
        st = ck["AH" + str(rows["F11-" + sid])].value
        print(f"{tag:5s} {typ[:22]:22s} {idv[:30]:30s} -> {st[:60]}")
    print(tag, "batch", batch, "->", ck["AH" + str(rows["F11-T24-01"])].value[:60], "| F11:", wb["Flights"]["AT15"].value)


run("good", GOOD, "261008-017")
run("bad", BAD, "Morning batch 0600")


def run_verifiers():
    cases = ["Siti Aminah", "QA", "CSM", "Duty CSM", "A. Rahman", "Rahman", "PASB supervisor"]
    p = T.TMP.format("ev_verifier"); shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p); wb["Settings"]["B4"] = datetime(2026, 10, 12, 7, 49); L.fill_f11(wb)
    ck = wb["Checks"]; rows = {ck.cell(r, 1).value: r for r in range(5, ck.max_row + 1)}
    for sid, v in zip(IDS, cases):
        ck[f"AE{rows['F11-' + sid]}"] = v
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "300"], capture_output=True)
    ck = openpyxl.load_workbook(p, data_only=True)["Checks"]
    for sid, v in zip(IDS, cases):
        print(f"verif {v:16s} (PIC {ck['T' + str(rows['F11-' + sid])].value}) -> {ck['AH' + str(rows['F11-' + sid])].value[:70]}")


run_verifiers()
