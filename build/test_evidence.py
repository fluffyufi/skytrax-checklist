"""Structured evidence (type in S + ID in Z): positive and negative cases on F11."""
import sys, shutil, subprocess
from datetime import datetime
import openpyxl
sys.argv = [sys.argv[0]]
sys.path.insert(0, "/home/user/skytrax-checklist/build")
import test_logic2 as L, test_ready as T

GOOD = [("Form / checklist", "SF-2210"), ("Photo", "IMG_2231"), ("Photo", "IMG_2231-2236"), ("Seal no.", "88213-88220"),
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
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "900"], capture_output=True)
    wb = openpyxl.load_workbook(p, data_only=True); ck = wb["Checks"]
    for sid, (typ, idv) in zip(IDS, cases):
        st = ck["AH" + str(rows["F11-" + sid])].value
        print(f"{tag:5s} {typ[:22]:22s} {idv[:30]:30s} -> {st[:60]}")
    print(tag, "batch", batch, "->", ck["AH" + str(rows["F11-T24-01"])].value[:60], "| F11:", wb["Flights"]["AT15"].value)


GOOD2 = [("Load / uplift sheet", "LS/OCT/0118"), ("Delivery note / receipt", "DN/OCT/2210"), ("Form / checklist", "SF/DEC/2210"),
         ("Form / checklist", "CAT/MAR/0045"), ("Form / checklist", "HR-2210"), ("Memo / letter", "PM-2210"),
         ("Form / checklist", "AM-2210"), ("Seal no.", "1045521"), ("Email", "EM-26-0441"), ("Photo", "IMG_4410"),
         ("System record", "QP-5521"), ("Form / checklist", "QF-44")]
BAD2 = [("Form / checklist", "9OCT2026"), ("Form / checklist", "10-Oct-2026"), ("Form / checklist", "2026/10/09"),
        ("Form / checklist", "20261009"), ("Form / checklist", "1415MYT"), ("Form / checklist", "14h15"),
        ("Photo", "1415-1430"), ("Form / checklist", "ETD2150"), ("Form / checklist", "Revision12"),
        ("Form / checklist", "v12"), ("Form / checklist", "A350-900"), ("Form / checklist", "Seat14C")]
BAD3 = [("Photo", "IMG_2231.jpg"), ("Photo", "https://magcs.sharepoint.com/qa/SF-2210"), ("Form / checklist", "SF-2210,IMG_2231"),
        ("Form / checklist", "SF-2210/IMG_2231"), ("Form / checklist", "00000"), ("Form / checklist", "12345"),
        ("Form / checklist", "XX-0000"), ("Form / checklist", "ABC-123"), ("Form / checklist", "NA-12345"),
        ("Form / checklist", "11111"), ("Form / checklist", "SF-00"), ("Form / checklist", "10OCT26-1415")]

GOOD3 = [("Form / checklist", "SF-2222"), ("Form / checklist", "2210A"), ("Form / checklist", "B45213"),
         ("Form / checklist", "No.45213"), ("Link / file path", "https://magcs.sharepoint.com/qa/F11/SF-2210"),
         ("Link / file path", "S:\\QA\\F11\\SF-2210.pdf"), ("Seal no.", "1045521"), ("Photo", "IMG_4410"),
         ("Email", "EM-26-0441"), ("System record", "QP-5521"), ("Form / checklist", "QF-44"), ("Photo", "DSC04410")]
BAD4 = [("Form / checklist", "2026-10-08T10:15"), ("Form / checklist", "8.10.26"), ("Form / checklist", "081026"),
        ("Form / checklist", "261008"), ("Form / checklist", "20261008101530"), ("Form / checklist", "20261008-017"),
        ("Form / checklist", "Q4-2026"), ("Form / checklist", "SF-2210/2211"), ("Link / file path", "www.google.com"),
        ("Link / file path", "\\\\fs01\\QA\\F11"), ("Link / file path", "C:\\Users\\ali\\Desktop\\photo.jpg"),
        ("Form / checklist", "A350")]

if __name__ == "__main__":
    run("good3", GOOD3, "PASB-261008-BC-017")
    run("bad4", BAD4, "261009")
    run("good", GOOD, "261008-017")
    run("bad", BAD, "Morning batch 0600")
    run("good2", GOOD2, "PASB-261008-BC-017")
    run("bad2", BAD2, "9OCT2026")
    run("bad3", BAD3, "Batch-1")


def run_verifiers():
    cases = ["Siti Aminah", "QA", "CSM", "Duty CSM", "A. Rahman", "Rahman", "PASB supervisor"]
    p = T.TMP.format("ev_verifier"); shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p); wb["Settings"]["B4"] = datetime(2026, 10, 12, 7, 49); L.fill_f11(wb)
    ck = wb["Checks"]; rows = {ck.cell(r, 1).value: r for r in range(5, ck.max_row + 1)}
    for sid, v in zip(IDS, cases):
        ck[f"AE{rows['F11-' + sid]}"] = v
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "900"], capture_output=True)
    ck = openpyxl.load_workbook(p, data_only=True)["Checks"]
    for sid, v in zip(IDS, cases):
        print(f"verif {v:16s} (PIC {ck['T' + str(rows['F11-' + sid])].value}) -> {ck['AH' + str(rows['F11-' + sid])].value[:70]}")



def run_people():
    """(PIC, verifier, expected ok) on 12 rows."""
    cases = [("A. Rahman", "Siti Aminah", True), ("Toby Lim", "Ruby Tan", True), ("Esra Yilmaz", "Isra Omar", True),
             ("Abby Wong", "Ali", True), ("A. Rahman", "Shift Supervisor", False), ("A. Rahman", "Station Manager KUL", False),
             ("A. Rahman", "Purser", False), ("A. Rahman", "Quality Assurance", False), ("A. Rahman", "Cabin Services Manager", False),
             ("Aisyah Rahman", "Rahman, Aisyah", False), ("Station Manager KUL", "Siti Aminah", False), ("A. Rahman", "SATS QA", False)]
    p = T.TMP.format("ev_people"); shutil.copy(T.SRC, p)
    wb = openpyxl.load_workbook(p); wb["Settings"]["B4"] = datetime(2026, 10, 12, 7, 49); L.fill_f11(wb)
    ck = wb["Checks"]; rows = {ck.cell(r, 1).value: r for r in range(5, ck.max_row + 1)}
    for sid, (pic, ver, _) in zip(IDS, cases):
        ck[f"T{rows['F11-' + sid]}"] = pic; ck[f"AE{rows['F11-' + sid]}"] = ver
    wb.save(p); subprocess.run([sys.executable, T.RECALC, p, "900"], capture_output=True)
    ck = openpyxl.load_workbook(p, data_only=True)["Checks"]
    for sid, (pic, ver, exp) in zip(IDS, cases):
        st = ck['AH' + str(rows['F11-' + sid])].value
        flag = "ok " if st.startswith("COMPLETE") == exp else "BAD"
        print(f"people {flag} {pic:20s} / {ver:22s} -> {st[:70]}")


if __name__ == "__main__":
    run_verifiers()
    run_people()
