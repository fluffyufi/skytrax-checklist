"""Single source of truth for the Skytrax 2026 catering readiness workbook.

Transcribes the 22 legs of the "Skytrax Agenda 2026" slide and derives the
per-flight uplift requirements from reference/STD_UPLIFT_INFORMATION.xlsx,
citing the exact reference cell for every requirement. Nothing here is
invented: where the reference is silent or contradictory the item is marked
"Clarification required" with the open question.

Run:  python build/data.py   -> writes build/data.json
"""
import json
import os
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, "..", "reference", "STD_UPLIFT_INFORMATION.xlsx")
REFNAME = "STD UPLIFT INFORMATION.xlsx"

# ---------------------------------------------------------------- schedule
# (itin, date, day, flt, dep, std, arr, sta, class, fleet, seat, transit, remark)
SCHEDULE = [
    (1, "2026-10-08", "THU", "MH0001", "LHR", "2135", "KUL", "1750+1", "EY", "A350", "14C", "-", ""),
    (1, "2026-10-09", "FRI", "MH0727", "KUL", "2150", "CGK", "2305", "BC", "B738MAX", "1C", "4H 00M", "Overnight CGK"),
    (1, "2026-10-10", "SAT", "MH0720", "CGK", "1605", "KUL", "1925", "EY", "A333", "11C", "-", ""),
    (1, "2026-10-10", "SAT", "MH0088", "KUL", "2330", "NRT", "0740+1", "BC", "A350", "1A", "4H 05M", "Overnight NRT"),
    (1, "2026-10-15", "THU", "MH0089", "NRT", "1005", "KUL", "1645", "EY", "A350", "9K", "-", ""),
    (1, "2026-10-15", "THU", "MH0002", "KUL", "2320", "LHR", "0555+1", "BC", "A350", "2K", "6H 35M", ""),
    (2, "2026-10-09", "FRI", "MH0003", "LHR", "1100", "KUL", "0645+1", "BC", "A350", "2A", "-", ""),
    (2, "2026-10-10", "SAT", "MH0752", "KUL", "0935", "HAN", "1215", "EY", "B738MAX", "4A", "2H 50M", "Overnight HAN"),
    (2, "2026-10-11", "SUN", "MH0753", "HAN", "1305", "KUL", "1730", "BC", "B738MAX", "1F", "-", "Overnight KUL"),
    (2, "2026-10-12", "MON", "MH1140", "KUL", "1145", "PEN", "1245", "BC", "B738MAX", "1F", "-", ""),
    (2, "2026-10-12", "MON", "MH1149", "PEN", "1550", "KUL", "1655", "EY", "B738MAX", "4A", "3H 05M", ""),
    (2, "2026-10-12", "MON", "MH0139", "KUL", "2225", "ADL", "0800+1", "BC", "A339", "1K", "5H 30M", "Overnight ADL"),
    (2, "2026-10-14", "WED", "MH0138", "ADL", "1045", "KUL", "1600", "EY", "A339", "11K", "-", "Overnight KUL"),
    (2, "2026-10-15", "THU", "MH0004", "KUL", "0950", "LHR", "1635", "EY", "A350", "14C", "-", ""),
    (3, "2026-10-25", "SUN", "MH0003", "LHR", "1100", "KUL", "0645+1", "EY", "A350", "12H", "-", ""),
    (3, "2026-10-26", "MON", "MH0072", "KUL", "0910", "HKG", "1315", "EY", "B738MAX", "4C", "2H 25M", "Overnight HKG"),
    (3, "2026-10-27", "TUE", "MH0073", "HKG", "1450", "KUL", "1855", "BC", "B738MAX", "1C", "-", ""),
    (3, "2026-10-27", "TUE", "MH0052", "KUL", "2225", "KIX", "0555+1", "BC", "A333", "1D", "3H 30M", "Overnight KIX"),
    (3, "2026-10-29", "THU", "MH0053", "KIX", "0940", "KUL", "1645", "EY", "A333", "9C", "-", ""),
    (3, "2026-10-29", "THU", "MH1450", "KUL", "1830", "LGK", "2000", "EY", "B738MAX", "4A", "1H 45M", "Overnight LGK"),
    (3, "2026-10-30", "FRI", "MH1437", "LGK", "1445", "KUL", "1605", "BC", "B738MAX", "1C", "-", ""),
    (3, "2026-10-30", "FRI", "MH0002", "KUL", "2320", "LHR", "0555+1", "BC", "A350", "5A", "7H 15M", ""),
]

# Station time zones. DST windows are the ones that fall inside the Oct-2026
# programme (IANA tz database rules).
STATIONS = {
    #       IANA zone             std   dst   dst_start_utc        dst_end_utc
    "KUL": ("Asia/Kuala_Lumpur", 8.0, None, None, None),
    "PEN": ("Asia/Kuala_Lumpur", 8.0, None, None, None),
    "LGK": ("Asia/Kuala_Lumpur", 8.0, None, None, None),
    "LHR": ("Europe/London", 0.0, 1.0, "2026-03-29 01:00", "2026-10-25 01:00"),
    "CGK": ("Asia/Jakarta", 7.0, None, None, None),
    "HAN": ("Asia/Ho_Chi_Minh", 7.0, None, None, None),
    "NRT": ("Asia/Tokyo", 9.0, None, None, None),
    "KIX": ("Asia/Tokyo", 9.0, None, None, None),
    "HKG": ("Asia/Hong_Kong", 8.0, None, None, None),
    "ADL": ("Australia/Adelaide", 9.5, 10.5, "2026-10-03 16:30", "2027-04-03 16:30"),
}

# Schedule fleet -> reference AIRCRAFT TYPE row(s)
FLEET_REF = {
    "A350": ("A350-900: A359 (9M-MAB..MAG) or A350 (9M-MAH) - tail to confirm",
             "AIRCRAFT TYPE!C6:R7"),
    "A333": ("A330-300 (A333)", "AIRCRAFT TYPE!C9:R9"),
    "A339": ("A330-900 (A339)", "AIRCRAFT TYPE!C8:R8"),
    "B738MAX": ("B737-8 (B7M8)", "AIRCRAFT TYPE!C16:R16"),
}
WIDEBODY = {"A350", "A333", "A339"}

# Round-trip catered return legs: the reference names KUL as the only uplift
# station for these sectors, so the return meals travel from KUL on the
# outbound leg of the same rotation (aircraft rotation still to be confirmed).
ROUND_TRIP_LOADED_ON = {"F11": "F10", "F21": "F20"}
# Pair (inbound KUL-outstation) flight of the aircraft that operates each return leg. MAGCS: toiletry kits, sales
# carts, compendium and meal cart covers are double-loaded at KUL on the pair flight. Timings from the published
# schedule; pairing confirmed by matching registrations on planemapper.com (MH721/MH720 9M-MTM 30-Sep, MH1148/MH1149
# 9M-MXE 28-Sep, MH1436/MH1437 9M-MXC/MXH, MH88/MH89 9M-MTG) or by the published turn (MH4->MH1, MH2->MH3).
# (pair flight, KUL departure local MYT, block hours)
PAIR = {
    "F01": ("MH0004", "2026-10-08 09:50", 13.75),
    "F03": ("MH0721", "2026-10-10 13:45", 2.33),
    "F05": ("MH0088", "2026-10-14 23:30", 7.17),
    "F07": ("MH0002", "2026-10-08 23:20", 13.58),
    "F09": ("MH0752", "2026-10-11 09:35", 3.67),
    "F11": ("MH1148", "2026-10-12 14:05", 1.00),
    "F13": ("MH0139", "2026-10-13 22:25", 7.08),
    "F15": ("MH0002", "2026-10-24 23:20", 13.58),
    "F17": ("MH0072", "2026-10-27 09:10", 4.08),
    "F19": ("MH0052", "2026-10-28 22:25", 6.50),
    "F21": ("MH1436", "2026-10-30 13:10", 1.08),
}
# MAGCS requirements given 01-Oct-2026 (not in the reference workbook); double-loaded from KUL
MAGCS_SRC = "MAGCS instruction 01-Oct-2026"
# Skytrax tail groups (Engineering 'best aircraft'; tail confirmed 48 h prior)
TAIL_GROUPS = {"B738MAX": ["9M-MVO", "9M-MVP", "9M-MVQ", "9M-MVR"], "A333": ["9M-MTJ", "9M-MTM", "9M-MTG"],
               "A350": [f"9M-MA{c}" for c in "BCDEFGH"]}
# caterer per flight (MAGCS, 01-Oct-2026)
CATERER = {"F01": "dnata", "F02": "MAGCS", "F03": "Purantara", "F04": "PASB", "F05": "TFK", "F06": "PASB", "F07": "dnata",
           "F08": "MAGCS", "F09": "NCS", "F10": "MAGCS", "F11": "MAGCS", "F12": "PASB", "F13": "dnata", "F14": "PASB",
           "F15": "dnata", "F16": "MAGCS", "F17": "CPCS", "F18": "PASB", "F19": "AAS", "F20": "MAGCS", "F21": "MAGCS",
           "F22": "PASB"}

# ---------------------------------------------------------------- reference
wb = openpyxl.load_workbook(REF)


def _merged_value(ws, row, col):
    for r in ws.merged_cells.ranges:
        if r.min_row <= row <= r.max_row and r.min_col <= col <= r.max_col:
            return ws.cell(r.min_row, r.min_col).value
    return ws.cell(row, col).value


def cell(sheet, addr):
    ws = wb[sheet]
    c = ws[addr]
    v = _merged_value(ws, c.row, c.column)
    return "" if v is None else str(v).strip()


def find_row(sheet, col, text):
    ws = wb[sheet]
    ci = openpyxl.utils.column_index_from_string(col)
    hits = [r for r in range(1, ws.max_row + 1)
            if str(_merged_value(ws, r, ci) or "").strip() == text]
    return hits


def tick(v, widebody=True):
    v = (v or "").strip()
    if v == "_/":
        return True
    if v == "_/*":  # SEAT LINEN!B72:C72 - wide body aircraft only
        return widebody
    return False


def cite(sheet, addr):
    """Cell citation; names the whole merged range when the cell sits inside one."""
    ws = wb[sheet]
    c = ws[addr]
    for r in ws.merged_cells.ranges:
        if r.min_row <= c.row <= r.max_row and r.min_col <= c.column <= r.max_col:
            return str(r.coord)
    return addr


def outstation(dep, arr):
    return arr if dep == "KUL" else dep


def local_dt(date, hhmm, plus=0):
    d = datetime.strptime(date, "%Y-%m-%d") + timedelta(days=plus)
    return d.replace(hour=int(hhmm[:2]), minute=int(hhmm[2:4]))


def galley_info(fleet):
    sh = "AIRCRAFT TYPE"
    rows = {"A350": [6, 7], "A333": [9], "A339": [8], "B738MAX": [16]}[fleet]
    parts = []
    for r in rows:
        v = {c: cell(sh, f"{c}{r}") for c in "CIJKLMNOPQR"}
        parts.append(f"{v['C']}: carts {v['I']}, SU {v['J']}, ovens {v['K']}, warming ovens {v['L']}, fridges {v['M']}, "
                     f"water heaters {v['N']}, beverage makers {v['O']}, espresso {v['P']}, folding trolleys {v['Q']}, "
                     f"headsets {v['R']} ({sh}!C{r}:R{r})")
    return " | ".join(parts)


def build():
    flights = []
    for i, (itin, date, day, flt, dep, std, arr, sta, cls, fleet, seat, transit, remark) in enumerate(SCHEDULE, 1):
        fid = f"F{i:02d}"
        plus = 1 if sta.endswith("+1") else 0
        std_l = local_dt(date, std)
        sta_l = local_dt(date, sta[:4], plus)
        out = outstation(dep, arr)
        sector = f"KUL/{out}/KUL"
        # CATERING UPLIFT STN
        r = find_row("CATERING UPLIFT STN", "D", sector)
        assert len(r) == 1, sector
        r = r[0]
        region = cell("CATERING UPLIFT STN", f"C{r}")
        upl = cell("CATERING UPLIFT STN", f"E{r}")
        svc_f = cell("CATERING UPLIFT STN", f"F{r}")
        svc_g = cell("CATERING UPLIFT STN", f"G{r}")
        upl_stns = [s.strip() for s in upl.split("/")]
        meal_upl = dep if dep in upl_stns else "KUL"
        tz = ZoneInfo(STATIONS[dep][0])
        std_utc = std_l.replace(tzinfo=tz).astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
        sta_utc = sta_l.replace(tzinfo=ZoneInfo(STATIONS[arr][0])).astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
        block_h = (sta_utc - std_utc).total_seconds() / 3600
        flights.append(dict(
            id=fid, itin=itin, seq=i, date=date, day=day, flt=flt, dep=dep, arr=arr,
            std_local=std_l.strftime("%Y-%m-%d %H:%M"), sta_local=sta_l.strftime("%Y-%m-%d %H:%M"),
            sta_text=sta, std_text=std,
            cls=cls, fleet=fleet, fleet_ref=FLEET_REF[fleet][0], fleet_ref_src=FLEET_REF[fleet][1],
            seat=seat, transit=transit, remark=remark,
            sector=sector, region=region, service=f"{svc_g} ({svc_f})" if svc_f != svc_g else svc_g,
            service_src=f"CATERING UPLIFT STN!C{r}:G{r}", ref_uplift_stns=upl,
            meal_uplift_stn=meal_upl,
            round_trip=fid in ROUND_TRIP_LOADED_ON,
            loaded_on=ROUND_TRIP_LOADED_ON.get(fid, ""),
            std_utc=std_utc.strftime("%Y-%m-%d %H:%M"), block_h=round(block_h, 3),
            widebody=fleet in WIDEBODY,
            galley_info=galley_info(fleet),
            caterer=CATERER.get(fid, ""), tail_group=TAIL_GROUPS.get(fleet, []),
        ))
        if fid in PAIR:
            pf, pl, pb = PAIR[fid]
            pdt = datetime.strptime(pl, "%Y-%m-%d %H:%M")
            flights[-1].update(pair_flt=pf, pair_kul_local=pl, pair_block=pb,
                               pair_kul_utc=(pdt - timedelta(hours=8)).strftime("%Y-%m-%d %H:%M"))
    return flights


# ---------------------------------------------------------------- requirements
def requirements_for(f):
    """Return list of dicts: category, item, cls, uplift_stn, expected, src,
    applic ('Required' | 'Clarification required' | 'Not provided (ref X)' | 'Not applicable to assessed class'),
    note, qty (expected quantity text or '')."""
    reqs = []
    out = outstation(f["dep"], f["arr"])
    sector = f["sector"]
    cls = f["cls"]
    bc = cls == "BC"

    def add(category, item, applic, src, expected="", uplift="", note="", qty="", cls_scope=cls):
        reqs.append(dict(category=category, item=item, applic=applic, src=f"{REFNAME} > {src}",
                         expected=expected, uplift_stn=uplift, note=note, qty=qty, cls_scope=cls_scope))

    # ---- AMENITIES (slippers explicitly BS/BC, pajamas BS/BC; toiletry class not stated)
    sh = "AMENITIES"
    if out == "LHR":
        tk_rows = find_row(sh, "C", "KUL/LHR/KUL (A350)")  # rows 12, 13
        r_tk, r_pj = tk_rows[0], tk_rows[1]
        r_sl = r_pj
    else:
        rr = find_row(sh, "C", sector)
        assert len(rr) == 1, (sh, sector)
        r_tk = r_pj = r_sl = rr[0]
    C = lambda a: cite(sh, a)
    tk = cell(sh, f"D{r_tk}")
    tk_stn = cell(sh, f"H{r_tk}")
    if tick(tk) and bc:  # MAGCS: toiletry kits are BC only (EY receives none); double-loaded from KUL
        src = f"{sh}!{C(f'D{r_tk}')} / {C(f'H{r_tk}')}; {MAGCS_SRC}"
        add("Amenities", "Toiletry kits (BC)", "Required", src, "Available (ticked in reference)", "KUL",
            "Toiletry kits are BC only and are double-loaded from KUL (MAGCS)." +
            ("" if f["dep"] == "KUL" else " This leg's kits travel on the pair flight from KUL."))
    # pajamas (BC red eyes only) - column F
    if bc:
        pj = cell(sh, f"F{r_pj}")
        if tick(pj):
            dep_h = int(f["std_text"][:2])
            src = f"{sh}!{C(f'F{r_pj}')} / {C(f'H{r_pj}')}"
            if f["dep"] == "KUL" and dep_h >= 21:
                add("Amenities", "Pajamas (BC)", "Required", src, "Available - BC red-eye only", cell(sh, f"H{r_pj}"),
                    f"Red-eye departure {f['std_text']} local.")
            # otherwise not a red-eye (MAGCS confirmed MH0003 LHR 11:00 is not): no pajamas
        sl = cell(sh, f"G{r_sl}")
        if tick(sl):
            rem = cell(sh, f"I{r_sl}")
            add("Amenities", "Slippers (BS/BC)" + (f" - {rem}" if rem else ""), "Required",
                f"{sh}!{C(f'G{r_sl}')} / {C(f'H{r_sl}')}" + (f" / {C(f'I{r_sl}')}" if rem else ""),
                "Available (ticked in reference)", cell(sh, f"H{r_sl}"))
        elif sl == "X":
            add("Amenities", "Disposable slippers in DAM cart (BC, on request)", "Required",
                f"{sh}!{C(f'G{r_sl}')} / C83", "Available in DAM cart on request basis", "Not stated in reference",
                "Sector has no BC slippers; note C83: disposable slipper available in DAM cart (on request basis).")

    # ---- F&B LINEN (BS & BC) - BC assessed flights only
    sh = "F&B LINEN"
    C = lambda a: cite(sh, a)
    if bc:
        r = find_row(sh, "C", sector)
        assert len(r) == 1
        r = r[0]
        cols = [("D", "Towel"), ("E", "Tray cloth"), ("F", "Napkin"), ("G", "Table cloth"),
                ("H", "Bread linen (using table cloth)"), ("I", "Trolley cloth"), ("J", "Linen bag")]
        region = f["region"]
        for c, name in cols:
            v = cell(sh, f"{c}{r}")
            if not tick(v, f["widebody"]):
                continue
            exp = "Available (ticked in reference)"
            applic = "Required"
            note = ""
            if name == "Trolley cloth":
                exp = ("Wide body: Top/Middle & Bottom" if f["widebody"] else "Narrow body: Top only")
                note = f"Per note {sh}!B79:C80."
            if name in ("Table cloth", "Bread linen (using table cloth)"):
                if "Refreshment" in f["service"] or f["service"].startswith("0.5"):
                    continue  # MAGCS: not applicable on refreshment / 0.5 meal sectors
                note = "MAGCS: table cloth applies on all sectors except refreshment / 0.5 meal."
            add("F&B Linen", name, applic, f"{sh}!{C(f'{c}{r}')}", exp, "Not stated in reference", note)

    # ---- SEAT LINEN
    sh = "SEAT LINEN"
    C = lambda a: cite(sh, a)
    r = find_row(sh, "C", sector)
    assert len(r) == 1
    r = r[0]
    if bc:
        cols = [("E", "Seat pillow (BS & BC)"), ("F", "Duvet (BS & BC)"), ("G", "Mattress (BS & BC)"),
                ("H", "Blanket (BS & BC)")]
    else:
        cols = [("I", "Seat pillow (EY)"), ("J", "Blanket (EY)")]
    for c, name in cols:
        v = cell(sh, f"{c}{r}")
        if not tick(v, f["widebody"]):
            continue
        qty = ""
        note = ""
        applic = "Required"
        if name == "Blanket (EY)":
            note = f"EY blanket: flight above 3 hours excl. Regional ({sh}!B74:C74); block {f['block_h']:.2f} h."
            if f["block_h"] <= 3:
                applic = "Clarification required"
                note += " Matrix shows _/ but block time is not above 3 h - confirm."
            if f["region"] in ("ORIENTAL", "ASEAN", "DOMESTIC"):
                note += f" MAGCS: blanket included on this {f['region']} sector at 100% (one per EY seat)."
                if not qty and f["fleet"] not in ("A333", "A339", "A350"):
                    qty = "100% (one per EY seat)"
            if f["fleet"] in ("A333", "A339"):
                qty = "280"
                note += f" Qty 280 pcs (14 bundles) for A332/A333/A339 ({sh}!B75:C75)."
            elif f["fleet"] == "A350":
                qty = "280 (A350 9M-MAH) / 260 (A359) - per tail"
                note += f" Qty per tail: A350 280 pcs ({sh}!B76:C76), A359 260 pcs ({sh}!B77:C77)."
            else:
                note += " Quantity for this aircraft not stated in reference."
        add("Seat Linen", name, applic, f"{sh}!{C(f'{c}{r}')}", "Available 100% (ticked in reference)", "Not stated in reference", note, qty)

    # ---- SIGNATURE DRINKS (BSCL/BCL) - BC only
    sh = "Signature Drinks"
    C = lambda a: cite(sh, a)
    if bc:
        r = find_row(sh, "D", sector)
        assert len(r) == 1
        r = r[0]
        if tick(cell(sh, f"F{r}")):
            add("Signature Drinks", "Signature drink & garnish (BC)", "Required", f"{sh}!{C(f'F{r}')} / {C(f'G{r}')} / D79",
                "Available (ticked in reference)", "KUL", "Signature drinks & garnish uplifted from KUL (note D79).")

    # ---- SALES CART (all classes, flights 4 h & above)
    sh = "SALES CART"
    C = lambda a: cite(sh, a)
    rr = find_row(sh, "D", sector)
    if rr:
        r = rr[0]
        caterer = cell(sh, f"C{r}")
        stn_note = ""
        applic = "Required"
        cart_stn = "KUL"  # MAGCS: double-loaded from KUL (return legs: on the pair flight)
        if f["dep"] != "KUL":
            stn_note = f" Double-loaded from KUL on the pair flight (MAGCS)."
        if f["block_h"] >= 4:
            if f["fleet"] == "B738MAX":
                loc = cell(sh, f"F{r}")
                add("Sales Cart", f"Sales cart (caterer {caterer})", applic,
                    f"{sh}!{C(f'F{r}')} / {C(f'C{r}')} / {C(f'M{r}')} / C40",
                    f"Location {loc} (B7M8)", cart_stn,
                    f"Block {f['block_h']:.2f} h >= 4 h (note {sh}!C40).{stn_note}", cls_scope="All")
            elif f["fleet"] == "A333":
                add("Sales Cart", f"Sales cart (caterer {caterer})", applic, f"{sh}!{C(f'I{r}')} / {C(f'C{r}')} / C40",
                    f"Location {cell(sh, f'I{r}')} (A333, full cart)", cart_stn,
                    f"Block {f['block_h']:.2f} h >= 4 h.{stn_note}", cls_scope="All")
            elif f["fleet"] == "A339":
                add("Sales Cart", f"Sales cart (caterer {caterer})", applic, f"{sh}!{C(f'J{r}')} / {C(f'C{r}')} / C40",
                    f"Location {cell(sh, f'J{r}')} (A339)", cart_stn,
                    f"Block {f['block_h']:.2f} h >= 4 h.{stn_note}", cls_scope="All")
            elif f["fleet"] == "A350" and out != "LHR":
                k = cell(sh, f"K{r}")
                l = cell(sh, f"L{r}")
                add("Sales Cart", f"Sales cart (caterer {caterer})", applic,
                    f"{sh}!{C(f'K{r}')} / {C(f'L{r}')} / {C(f'C{r}')} / {C(f'O{r}')}",
                    f"A359: {k} (full cart); 9M-MAH: {l} (half cart) - per tail", cart_stn,
                    f"Block {f['block_h']:.2f} h >= 4 h. Location depends on tail.{stn_note}", cls_scope="All")
    # ---- MAGCS additional requirements (01-Oct-2026): every flight, double-loaded from KUL
    wb_ = f["widebody"]
    via = "" if f["dep"] == "KUL" else " This leg's set travels on the pair flight from KUL."
    add("Cabin items", "Compendium (in DAM cart)", "Required", MAGCS_SRC, "Location: DAM cart",
        "KUL", "Double-loaded from KUL (MAGCS)." + via, "2" if wb_ else "1", cls_scope="All")
    add("Cabin items", "Meal cart cover", "Required", MAGCS_SRC,
        "Wide body 8 (3 BC + 5 EY)" if wb_ else "Narrow body 5 (2 BC + 3 EY)",
        "KUL", "Double-loaded from KUL (MAGCS)." + via, "8" if wb_ else "5", cls_scope="All")
    return reqs


# Informational list of what the reference says is NOT provided on this flight
def not_provided(f):
    rows = []
    out = outstation(f["dep"], f["arr"])
    sector = f["sector"]
    bc = f["cls"] == "BC"
    sh = "AMENITIES"
    if out != "LHR":
        r = find_row(sh, "C", sector)[0]
        for c, name, need_bc in (("D", "Toiletry kits", False), ("F", "Pajamas (BC red-eye)", True), ("G", "Slippers (BS/BC)", True)):
            if need_bc and not bc:
                continue
            if cell(sh, f"{c}{r}") == "X":
                rows.append(("Amenities", name, f"{sh}!{c}{r}"))
    if bc:
        sh = "F&B LINEN"
        r = find_row(sh, "C", sector)[0]
        for c, name in (("G", "Table cloth"), ("H", "Bread linen")):
            if cell(sh, f"{c}{r}") == "X":
                rows.append(("F&B Linen", name, f"{sh}!{c}{r}"))
        sh = "Signature Drinks"
        r = find_row(sh, "D", sector)[0]
        if cell(sh, f"F{r}") == "X":
            rows.append(("Signature Drinks", "Signature drink (BCL)", f"{sh}!F{r}"))
    sh = "SEAT LINEN"
    r = find_row(sh, "C", sector)[0]
    cols = (("E", "Seat pillow"), ("F", "Duvet"), ("G", "Mattress"), ("H", "Blanket")) if bc else (("I", "Seat pillow (EY)"), ("J", "Blanket (EY)"))
    for c, name in cols:
        if not tick(cell(sh, f"{c}{r}"), f["widebody"]):
            rows.append(("Seat Linen", name, f"{sh}!{cite(sh, f'{c}{r}')} (blank)"))
    if out == "LHR":
        rows.append(("Sales Cart", "No sales cart location on LHR sector (A359 'not applicable for LHR'; 9M-MAH '-')",
                     "SALES CART!K7:K14 / L7"))
    elif not find_row("SALES CART", "D", sector):
        rows.append(("Sales Cart", "Sector not listed", "SALES CART!D7:D37"))
    elif f["block_h"] < 4:
        rows.append(("Sales Cart", f"Block {f['block_h']:.2f} h < 4 h", "SALES CART!C40"))
    return [dict(category=a, item=b, src=f"{REFNAME} > {c}") for a, b, c in rows]


# ---------------------------------------------------------------- check lines
STD_T7 = [
    ("Meal", "Meal sensory testing completed for this flight's menu cycle (assessed class)", "Record panel result / score in Result", 0, 0),
    ("Menu", "Printed menu cards confirmed (content, cycle, language, quantity)", "Menu card matches approved menu for this flight", 0, 0),
    ("IFE", "Digital menu in IFE confirmed (A339 only)", "Applies when fleet = A339", 0, 0),
    ("Staffing", "Experienced catering officer assigned to this flight", "Record officer name in Result", 0, 0),
    ("ISOP", "All latest applicable ISOP revisions communicated to caterer", "List revision numbers in Result (see ISOP register)", 0, 0),
    ("ISOP", "Caterer acknowledgement of ISOP revisions recorded", "Record acknowledgement ref in Evidence", 0, 0),
    ("Documents", "Galley loading diagram (GLD) received and its link recorded", "Link on the Documents sheet (ON FILE)", 0, 1),
    ("Documents", "Menu checklist received and its link recorded", "Link on the Documents sheet (ON FILE)", 0, 1),
    ("Uplift plan", "Uplift plan confirmed with caterer vs STD Uplift Information", "", 0, 0),
]
STD_T24 = [
    ("Meal", "Sensory test repeated on ACTUAL production batch for this flight", "Batch ID, result, CA and acceptance mandatory", 1, 0),
]
STD_T12 = [
    ("Meal", "Meal quality vs menu checklist", "", 0, 0),
    ("Meal", "Meal quantities vs menu checklist", "Enter expected & actual", 0, 0, 1),
    ("Equipment", "Equipment cleanliness vs GLD", "", 0, 0),
    ("Equipment", "Equipment presentation vs GLD", "", 0, 0),
    ("Equipment", "Equipment serviceability vs GLD", "", 0, 0),
    ("Equipment", "Equipment quantities vs GLD", "Enter expected & actual", 0, 0, 1),
    ("Equipment", "Loading positions planned vs GLD", "", 0, 0),
]
STD_UPL = [
    ("Meal", "Meals physically uplifted on board - quantity vs menu checklist", "Enter expected & actual", 0, 0, 1),
    ("Equipment", "Equipment physically loaded at GLD positions - quantity & positions", "Enter expected & actual", 0, 0, 1),
]


def check_lines(f, reqs, flights):
    lines = []
    inbound = [x for x in flights if x["dep"] == "KUL" and x["arr"] == f["dep"]]

    def add(cp, ctype, cat, item, expected, src, applic, note, station, batch=0, doc=0, qty=0,
            exp_qty="", uplift_stn="", due_bound=None, due_rule="", carry=0):
        n = sum(1 for l in lines if l["checkpoint"] == cp) + 1
        code = {"T-7D": "T7", "T-24H": "T24", "T-12H PREP": "T12", "UPLIFT": "UPL"}[cp]
        lines.append(dict(check_id=f"{f['id']}-{code}-{n:02d}", flight_id=f["id"], checkpoint=cp, check_type=ctype,
                          category=cat, item=item, expected=expected, src=src, applic=applic, note=note,
                          station=station, req_batch=batch, req_doc=doc, req_qty=qty, exp_qty=exp_qty,
                          uplift_stn=uplift_stn, due_bound=due_bound or [], due_rule=due_rule, req_carry=carry))

    ms = f["meal_uplift_stn"]
    for cat, item, exp, batch, doc in STD_T7:
        applic, note, src = "Required", "", "User brief (T-7 days)"
        if cat == "IFE":
            applic = "RULE:A339"
            note = "Digital IFE menu applies to A339 only (user brief). Applicability follows the Fleet cell on Flights."
        if cat == "Uplift plan":
            exp = f"Service: {f['service']}; meal uplift stn {ms} (ref lists {f['ref_uplift_stns']})"
            src = f"{REFNAME} > {f['service_src']}"
            if f["round_trip"]:
                note = "Return catering uplifted at KUL on the outbound rotation (ref uplift stn = KUL). Confirm aircraft rotation."
        if cat == "Menu" and "Refreshment" in f["service"]:
            note = "Refreshment service (peanut & water). If no printed menu card is used, mark N/A with the confirmation as justification."
        if cat == "Meal" and "Refreshment" in f["service"]:
            note = "Refreshment service: sensory test covers the refreshment items."
        add("T-7D", "Preparation", cat, item, exp, src, applic, note, ms, batch, doc)
    if f["round_trip"]:
        lf = next(x for x in flights if x["id"] == f["loaded_on"])
        add("T-7D", "Preparation", "Uplift plan",
            f"Pair flight {f['pair_flt']} (KUL {f['pair_kul_local'][8:10]}-Oct {f['pair_kul_local'][11:]} local) confirmed as the "
            f"carrier of this leg's return catering; actual KUL departure on Flights AX",
            f"All catering for {f['flt']} {f['dep']}-KUL is uplifted at KUL (ref uplift stn KUL only)", f"{REFNAME} > {f['service_src']}",
            "Required",
            f"{f['dep']} does not cater. Pair flight {f['pair_flt']} (aircraft rotation; tail confirmed 48 h prior) - "
            f"Flights AX is pre-filled with its scheduled KUL departure; correct it if the rotation changes.",
            "KUL", carry=1)
    kul_items = [r for r in reqs if r["uplift_stn"] == "KUL"]
    carry = f["dep"] != "KUL" and not f["round_trip"] and bool(kul_items)
    rt_bound = [x["id"] for x in inbound] if f["round_trip"] else None
    names = ", ".join(r["item"] for r in kul_items)
    if carry:
        lo = (datetime.strptime(f["std_utc"], "%Y-%m-%d %H:%M") - timedelta(days=3)).strftime("%Y-%m-%d %H:%M")
        prior = [x for x in inbound if lo <= x["std_utc"] < f["std_utc"]]  # a carrier more than 3 days early is not valid
        cand = (f"agenda candidate {prior[-1]['flt']} {prior[-1]['date'][8:]}-Oct ({prior[-1]['id']})" if prior else
                "no KUL-" + f["dep"] + " flight before this leg in the agenda")
        pair = (f"pair flight {f['pair_flt']} (KUL {f['pair_kul_local'][8:10]}-Oct {f['pair_kul_local'][11:]} local)"
                if f.get("pair_flt") else cand)
        add("T-7D", "Preparation", "Uplift plan",
            f"Pair flight carrying the KUL double-loaded items confirmed: {pair}; actual KUL departure on Flights AX",
            f"KUL-sourced items ({names}) travel on the inbound KUL-{f['dep']} pair flight",
            f"{REFNAME} > item uplift stn columns; {MAGCS_SRC}", "Required",
            f"Flights AX is pre-filled with the pair flight's scheduled KUL departure (aircraft rotation; tail confirmed 48 h "
            f"prior); correct it if the rotation changes. Prep and KUL-loading dues follow it.",
            "KUL", carry=1)
    for cat, item, exp, batch, doc in STD_T24:
        add("T-24H", "Preparation", cat, item, exp, "User brief (T-24 hours)", "Required",
            "Round-trip catered from KUL: due capped at the KUL loading deadline (Flights AY)." if rt_bound else "", ms, batch, doc,
            due_bound=rt_bound, due_rule="CARRY" if rt_bound else "")
    for t in STD_T12:
        cat, item, exp, batch, doc = t[:5]
        qty = t[5] if len(t) > 5 else 0
        if cat == "Equipment" and "quantities" in item:
            exp = f"{exp}. Aircraft galley data: {f['galley_info']}"
        add("T-12H PREP", "Preparation", cat, item, exp, "User brief (T-12 hours)", "Required",
            "Preparation check at caterer - does NOT confirm loading." +
            (" Round-trip catered from KUL: due capped at the KUL loading deadline (Flights AY)." if rt_bound else ""),
            ms, batch, doc, qty, due_bound=rt_bound, due_rule="CARRY" if rt_bound else "")
    for r in reqs:
        stn = r["uplift_stn"] if r["uplift_stn"] in STATIONS else ms
        exp = r["expected"] + (f"; qty {r['qty']}" if r["qty"] else "")
        rule, note = "", r["note"]
        if r["uplift_stn"].startswith("Not stated"):
            note = (note + " " if note else "") + (f"Reference gives no uplift stn for this item: preparation checked at the "
                                                   f"meal uplift stn ({ms}) - confirm with caterer.")
        if rt_bound:
            rule = "CARRY"
            note = (note + " " if note else "") + "Round-trip catered from KUL: due capped at the KUL loading deadline (Flights AY)."
        elif stn == "KUL" and carry:
            rule = "CARRY"
            note = (note + " " if note else "") + (f"Loaded at KUL on the inbound KUL-{f['dep']} flight: due = that flight's KUL "
                                                   "departure (Flights AX) minus the uplift window.")
        add("T-12H PREP", "Preparation", r["category"], r["item"] + " - prepared", exp, r["src"], r["applic"],
            note, stn, 0, 0, 1 if r["qty"] else 0, r["qty"], r["uplift_stn"],
            due_bound=(rt_bound or [x["id"] for x in inbound]) if rule else None, due_rule=rule)
    for t in STD_UPL:
        cat, item, exp, batch, doc, qty = t
        add("UPLIFT", "Physical uplift", cat, item, exp, "User brief (T-12 hours - physical uplift)", "Required",
            "Confirm on board before departure; completion time must fall inside the uplift window.", f["dep"], 0, 0, qty)
    if f["round_trip"]:
        allnames = ", ".join(["meals / refreshment", "equipment"] + [r["item"] for r in reqs])
        add("UPLIFT", "Physical uplift", "Uplift plan",
            f"Return catering for {f['flt']} physically loaded at KUL on the carrying flight (meals, equipment and all items below)",
            f"Items: {allnames}. Quantities as per menu checklist, GLD and Requirements", "User brief (physical uplift) + CATERING UPLIFT STN (KUL only)",
            "Required",
            "Loading happens at KUL, so it is confirmed at KUL inside the uplift window before the carrying flight's KUL "
            "departure (Flights AX). The on-board lines below are then re-checked at the departure station.", "KUL",
            due_bound=rt_bound, due_rule="CARRY")
    if carry:
        add("UPLIFT", "Physical uplift", "Uplift plan",
            f"KUL-sourced items physically loaded at KUL on the pair flight {f.get('pair_flt', '')} (KUL-{f['dep']})",
            f"Items: {names}. Quantities as per Requirements", f"{REFNAME} > item uplift stn columns", "Required",
            "Confirmed at KUL inside the uplift window before the carrying flight's KUL departure (Flights AX).", "KUL",
            due_bound=[x["id"] for x in inbound], due_rule="CARRY")
    for r in reqs:
        exp = r["expected"] + (f"; qty {r['qty']}" if r["qty"] else "")
        add("UPLIFT", "Physical uplift", r["category"], r["item"] + " - on board", exp, r["src"], r["applic"],
            r["note"], f["dep"], 0, 0, 1 if r["qty"] else 0, r["qty"], r["uplift_stn"])
    # link on-board rows to their preparation row (same item) for quantity / N/A consistency
    prep = {}
    for l in lines:
        if l["checkpoint"] == "T-12H PREP":
            key = l["item"][:-len(" - prepared")] if l["item"].endswith(" - prepared") else l["item"]
            prep[key] = l["check_id"]
    std_map = {"Meals physically uplifted on board - quantity vs menu checklist": "Meal quantities vs menu checklist",
               "Equipment physically loaded at GLD positions - quantity & positions": "Equipment quantities vs GLD"}
    for l in lines:
        l["prep_link"] = ""
        if l["checkpoint"] != "UPLIFT":
            continue
        if l["item"].endswith(" - on board"):
            l["prep_link"] = prep.get(l["item"][:-len(" - on board")], "")
        elif l["item"] in std_map:
            l["prep_link"] = prep.get(std_map[l["item"]], "")
    return lines


def expected_dues(f, window_h=6):
    """Independent zoneinfo computation of due instants (for critics / tests)."""
    std_utc = datetime.strptime(f["std_utc"], "%Y-%m-%d %H:%M")
    out = {}
    for k, h in (("T-7D", 168), ("T-24H", 24), ("T-12H PREP", 12), ("UPLIFT", 0)):
        out[k] = (std_utc - timedelta(hours=h)).strftime("%Y-%m-%d %H:%M")
    return out


if __name__ == "__main__":
    flights = build()
    byid = {f["id"]: f for f in flights}
    all_lines, all_reqs, all_np = [], [], []
    for f in flights:
        reqs = requirements_for(f)
        for r in reqs:
            r["flight_id"] = f["id"]
        all_reqs += reqs
        for n in not_provided(f):
            n["flight_id"] = f["id"]
            all_np.append(n)
        all_lines += check_lines(f, reqs, flights)
        f["expected_due_utc"] = expected_dues(f)
    # DST note where a checkpoint's local time differs from STD's offset (e.g. LHR BST->GMT on 25 Oct)
    seen = set()
    for ln in all_lines:
        f = byid[ln["flight_id"]]
        stn = ln["station"]
        if ln["checkpoint"] == "UPLIFT" or stn not in STATIONS or (f["id"], ln["checkpoint"]) in seen:
            continue
        hrs = {"T-7D": 168, "T-24H": 24, "T-12H PREP": 12}[ln["checkpoint"]]
        std = datetime.strptime(f["std_utc"], "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo("UTC"))
        tz = ZoneInfo(STATIONS[stn][0])
        if (std - timedelta(hours=hrs)).astimezone(tz).utcoffset() != std.astimezone(tz).utcoffset():
            ln["note"] = ((ln["note"] + " ") if ln["note"] else "") + (
                f"Clock change: {stn} local time changes between this due time and departure, so the due time "
                f"shows a different clock hour than STD ({hrs} h before STD is still correct).")
            seen.add((f["id"], ln["checkpoint"]))
    for ln in all_lines:
        ln["note"] = ln["note"].replace("shows _/", "shows a tick").replace("_/", "tick")
        ln["expected"] = ln["expected"].replace("_/", "tick")
    for r in all_reqs:
        r["note"] = r["note"].replace("shows _/", "shows a tick").replace("_/", "tick")
    for fid, lid in ROUND_TRIP_LOADED_ON.items():
        byid[fid]["loaded_on_flt"] = byid[lid]["flt"]
        byid[fid]["loaded_on_std_local"] = byid[lid]["std_local"]
    stations = [dict(code=k, zone=v[0], std=v[1], dst=v[2], dst_start_utc=v[3], dst_end_utc=v[4])
                for k, v in STATIONS.items()]
    json.dump(dict(flights=flights, stations=stations, requirements=all_reqs, not_provided=all_np,
                   checks=all_lines, ref_revision=cell("AIRCRAFT TYPE", "C22")),
              open(os.path.join(HERE, "data.json"), "w"), indent=1)
    print(len(flights), "flights", len(all_reqs), "reqs", len(all_lines), "check lines", len(all_np), "not-provided")
    for f in flights:
        n = sum(1 for l in all_lines if l["flight_id"] == f["id"])
        print(f["id"], f["flt"], f["dep"], f["arr"], f["cls"], f["fleet"], f["std_local"], "UTC", f["std_utc"],
              f"blk {f['block_h']:.2f}", f["service"], "upl", f["meal_uplift_stn"], n)
