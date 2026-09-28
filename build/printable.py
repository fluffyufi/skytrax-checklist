"""Printable per-flight checklists P01..P22 (owner: printable builder).

Each sheet is a formula VIEW over Flights / Checks / Documents / Settings; no
result is stored here. Static text from data.json is used only to size row
heights (LibreOffice does not auto-fit rows of openpyxl-written files).

Paper-first layout: 8 wide columns (~150 width units) so A4 landscape
fit-to-width prints at ~95-100 % with a 9 pt body; every check row leaves room
for hand-written or typed entries of ~140 characters in the combined cells.
"""
from datetime import datetime

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "1F3864"
FONT = "Arial"
FS = 9  # body font size

# ------------------------------------------------------------------ layout
COLS = [
    ("Check ID / item", 22),
    ("Requirement / expected", 25),
    ("Applicability / due (local @ stn)", 14),
    ("Status (CA status)", 11),
    ("Result  |  batch  |  exp / act qty", 34),
    ("Evidence  |  corrective action", 24),
    ("Completed (local) / verifier", 14),
    ("Record state", 11),
]
NCOL = len(COLS)
LASTCOL = get_column_letter(NCOL)
C_REC = get_column_letter(8)

CHECKPOINTS = [
    ("T-7D", "7 days before departure", "AG", "AJ"),
    ("T-24H", "Production batch sensory", "AH", "AK"),
    ("T-12H PREP", "Preparation check at caterer (does NOT confirm loading)", "AI", "AL"),
    ("UPLIFT", "Physical uplift – on-board confirmation before departure", "I", "AM"),
]
FIRST_CHECK_ROW = 5
FIRST_FLIGHT_ROW = 5
NFL = 22
TFMT = "DD-MMM-YY HH:MM"  # one due-time format everywhere

# ------------------------------------------------------------------ styles
thin = Side(style="thin", color="808080")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_NAVY = PatternFill("solid", fgColor=NAVY)
F_BAND = PatternFill("solid", fgColor="D9E1F2")
F_LABEL = PatternFill("solid", fgColor="F2F2F2")

RED_FILL = PatternFill("solid", fgColor="FFC7CE", bgColor="FFC7CE")
RED_STRONG = PatternFill("solid", fgColor="C00000", bgColor="C00000")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE", bgColor="C6EFCE")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C", bgColor="FFEB9C")
GREY_FILL = PatternFill("solid", fgColor="D9D9D9", bgColor="D9D9D9")

AL_TOP = Alignment(wrap_text=True, vertical="top", horizontal="left")
AL_MID = Alignment(wrap_text=True, vertical="center", horizontal="left")
AL_CEN = Alignment(wrap_text=True, vertical="center", horizontal="center")

LINE_PT = {8: 10.5, 9: 11.5, 10: 13.0, 11: 14.5, 12: 15.5, 13: 17.0}
ENTRY_LINES = 5  # room for ~140 chars of result/CA + batch + qty lines


def _width(c1, c2=None):
    return sum(COLS[i - 1][1] for i in range(c1, (c2 or c1) + 1))


def _lines(text, width, size=FS, bold=False):
    """Conservative wrapped line count for Arial text in `width` column units."""
    if text is None or text == "":
        return 1
    per_char = 1.5 * (8.0 / size)  # LibreOffice fits ~1.9 Arial-8 chars/unit; margin for Excel
    if bold:
        per_char *= 0.9
    cap = max(1, int(width * per_char) - 1)
    n = 0
    for para in str(text).split("\n"):
        cnt, line = 1, 0
        for w in para.split(" "):
            wl = len(w)
            if wl > cap:
                cnt += (1 if line else 0) + (wl - 1) // cap
                line = wl % cap or cap
                continue
            add = wl if line == 0 else wl + 1
            if line + add > cap:
                cnt += 1
                line = wl
            else:
                line += add
        n += cnt
    return n


def _height(nlines, size=FS, pad=4.0, minimum=0):
    return max(minimum, round(nlines * LINE_PT.get(size, size * 1.3) + pad, 1))


def _font(size=FS, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def _merge(ws, r1, c1, r2, c2):
    if r1 != r2 or c1 != c2:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _box(ws, r1, c1, r2, c2, side=thin):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(
                left=side if c == c1 else None, right=side if c == c2 else None,
                top=side if r == r1 else None, bottom=side if r == r2 else None)


def _put(ws, r, c1, value, c2=None, font=None, fill=None, align=AL_MID, fmt=None, border=True):
    c2 = c2 or c1
    cell = ws.cell(r, c1, value)
    cell.font = font or _font()
    cell.alignment = align
    if fill:
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill
    if fmt:
        cell.number_format = fmt
    _merge(ws, r, c1, r, c2)
    if border:
        _box(ws, r, c1, r, c2)
    return cell


def _band(ws, r, text, size=9, fill=F_NAVY, color="FFFFFF", lines=1):
    _put(ws, r, 1, text, NCOL, font=_font(size, True, color), fill=fill, align=AL_MID)
    ws.row_dimensions[r].height = _height(lines, size, 4)


def _blank(src):
    return f'=IF({src}="","",{src})'


def _t(ref):
    """Datetime -> 'dd-mmm-yy hh:mm' text, blank if not a number."""
    return f'IF(ISNUMBER({ref}),TEXT({ref},"{TFMT}"),"")'


def sheet_title(n, f):
    d = datetime.strptime(f["date"], "%Y-%m-%d")
    t = f"P{n:02d} {f['flt']} {d.strftime('%d%b').upper()}"
    for ch in '[]:*?/\\':
        t = t.replace(ch, "")
    return t[:31]


# ------------------------------------------------------------------ builder
def _build_one(wb, n, f, checks_idx):
    fr = FIRST_FLIGHT_ROW + n - 1
    FL = lambda col: f"Flights!${col}${fr}"  # noqa: E731
    DC = lambda col: f"Documents!${col}${fr}"  # noqa: E731
    ws = wb.create_sheet(sheet_title(n, f))
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for i, (_, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    lab = dict(font=_font(8, True), fill=F_LABEL)
    d_text = datetime.strptime(f["date"], "%Y-%m-%d").strftime("%d-%b-%y")
    flight_tag = f"{f['id']} {f['flt']} {d_text} {f['dep']}-{f['arr']} Class {f['cls']}"

    # ---------------- title (one row, formula-driven flight line)
    r = 1
    _put(ws, r, 1,
         f'="SKYTRAX 2026 CATERING READINESS CHECKLIST  |  {f["id"]}  "&{FL("F")}&"  "&TEXT({FL("D")},"DD-MMM-YY")'
         f'&" ("&{FL("E")}&")  "&{FL("G")}&" → "&{FL("H")}&"  |  Class "&{FL("K")}',
         NCOL, font=_font(12, True, "FFFFFF"), fill=F_NAVY)
    ws.row_dimensions[r].height = 21

    # ---------------- flight details
    V = dict(font=_font(FS))
    tail = f'=IF({FL("N")}="","(enter in Flights!N)",{FL("N")})'
    loaded = (f'=IF(TRIM({FL("W")})="","–",{FL("W")}&IFERROR(" "&INDEX(Flights!$F${FIRST_FLIGHT_ROW}:$F$'
              f'{FIRST_FLIGHT_ROW + NFL - 1},MATCH(TRIM({FL("W")}),Flights!$A${FIRST_FLIGHT_ROW}:$A$'
              f'{FIRST_FLIGHT_ROW + NFL - 1},0)),""))')
    # each item: (label, lc1, lc2, formula, vc1, vc2, sizing text)
    grid = [
        [("Flight no", 1, 1, f"={FL('F')}", 2, 2, f["flt"]),
         ("Date / day", 3, 3, f'=TEXT({FL("D")},"DD-MMM-YY")&" ("&{FL("E")}&")"', 4, 4, "08-Oct-26 (THU)"),
         ("Sector (DEP-ARR)", 5, 5, f'={FL("G")}&"-"&{FL("H")}', 6, 6, ""),
         ("Class assessed", 7, 7, f"={FL('K')}", 8, 8, "")],
        [("STD local (dep)", 1, 1, f'={_t(FL("I"))}&" "&{FL("G")}', 2, 2, "08-Oct-26 21:35 LHR"),
         ("STA local (arr)", 3, 3, f'={_t(FL("J"))}&" "&{FL("H")}', 4, 4, "09-Oct-26 17:50 KUL"),
         ("Seat", 5, 5, f"={FL('O')}", 6, 6, ""),
         ("Tail / Reg", 7, 7, tail, 8, 8, "(enter in Flights!N)")],
        [("Aircraft ref type", 1, 1, f"={FL('M')}", 2, 4, f["fleet_ref"]),
         ("Fleet (schedule)", 5, 5, f"={FL('L')}", 6, 6, f["fleet"]),
         ("Block (h)", 7, 7, f"={FL('AB')}", 8, 8, "")],
        [("Service type", 1, 1, f"={FL('T')}", 2, 2, f["service"]),
         ("Meal uplift stn", 3, 3, f"={FL('V')}", 4, 4, f["meal_uplift_stn"]),
         ("Ref uplift stns", 5, 5, f"={FL('U')}", 6, 6, f["ref_uplift_stns"]),
         ("Transit / remark", 7, 7, f'=TRIM({FL("P")}&"  "&{FL("Q")})', 8, 8, f"{f['transit']}  {f['remark']}")],
        [("Round-trip: loaded on", 1, 1, loaded, 2, 2, "F10 MH1140 (KUL 1145)"),
         ("Caterer", 3, 3, f'=IF({FL("X")}="","(enter in Flights!X)",{FL("X")})', 4, 5, "x" * 40),
         ("Flight PIC", 6, 6, f'=IF({FL("Y")}="","(enter in Flights!Y)",{FL("Y")})', 7, 8, "x" * 36)],
    ]
    for row in grid:
        r += 1
        nl = 1
        for label, l1, l2, formula, v1, v2, txt in row:
            _put(ws, r, l1, label, l2, **lab)
            _put(ws, r, v1, formula, v2, fmt="0.00" if label.startswith("Block") else None, **V)
            nl = max(nl, _lines(label, _width(l1, l2), 8, True), _lines(txt, _width(v1, v2)))
        ws.row_dimensions[r].height = _height(nl, FS, 4, 16)

    # ---------------- due times / readiness
    r += 1
    hdr = r
    for (c1, c2), t in zip([(1, 1), (2, 2), (3, 4), (5, 5)],
                           ["Checkpoint", "Purpose",
                            "Due (local) – T-7D / T-24H / T-12H PREP at meal uplift stn; UPLIFT = STD at DEPARTURE stn",
                            "% complete"]):
        _put(ws, r, c1, t, c2, font=_font(8, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    _put(ws, r, 6, "READINESS STATUS", 6, font=_font(8, True), fill=F_LABEL, align=AL_CEN)
    _put(ws, r, 7, f"={FL('AT')}", 8, font=_font(10, True), align=AL_CEN)
    ws.row_dimensions[r].height = _height(max(2, _lines(ws.cell(r, 3).value, _width(3, 4), 8, True)), 8, 4)
    readiness = f"$G${r}"
    ws.conditional_formatting.add(f"G{r}:H{r}", FormulaRule(formula=[f'LEFT({readiness},5)="READY"'],
                                                           fill=GREEN_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(f"G{r}:H{r}", FormulaRule(formula=[f'LEFT({readiness},9)="NOT READY"'],
                                                           fill=RED_FILL, font=Font(color="9C0006", bold=True),
                                                           stopIfTrue=True))
    ws.conditional_formatting.add(f"G{r}:H{r}", FormulaRule(formula=[f'{readiness}<>""'], fill=AMBER_FILL))

    counts = [("Overdue checks", "AP", True), ("Open discrepancies", "AQ", True),
              ("Open required checks", "AO", False), ("Invalid entries", "AU", True),
              ("Docs outstanding (0-2)", "AR", True)]
    due_txt = {
        "T-7D": f'={_t(FL("AG"))}&" "&{FL("V")}',
        "T-24H": f'={_t(FL("AH"))}&" "&{FL("V")}',
        "T-12H PREP": f'={_t(FL("AI"))}&" "&{FL("V")}&IF({FL("AI")}={FL("AH")},"  (same as T-24H – see note)","")',
        "UPLIFT": f'="STD "&{_t(FL("I"))}&" "&{FL("G")}&" (departure stn)"',
        "OVERALL": f'="Last due: STD "&{_t(FL("I"))}&" "&{FL("G")}',
    }
    rows = CHECKPOINTS + [("OVERALL", "All checkpoints", None, "AN")]
    for (cp, desc, _due, pct), (clab, ccol, red) in zip(rows, counts):
        r += 1
        _put(ws, r, 1, cp, 1, **lab)
        _put(ws, r, 2, desc, 2, font=_font(8, cp == "OVERALL"))
        _put(ws, r, 3, due_txt[cp], 4, font=_font(FS, cp != "OVERALL"), align=AL_CEN)
        _put(ws, r, 5, f"={FL(pct)}", 5, font=_font(FS, True), fmt="0%", align=AL_CEN)
        _put(ws, r, 6, clab, 6, **lab)
        cell = _put(ws, r, 7, f"={FL(ccol)}", 8, font=_font(10, True), align=AL_CEN)
        if red:
            cc = cell.coordinate
            ws.conditional_formatting.add(f"G{r}:H{r}", FormulaRule(
                formula=[f"AND(ISNUMBER(${cc}),${cc}>0)"], fill=RED_FILL, font=Font(color="9C0006", bold=True)))
        ws.row_dimensions[r].height = _height(max(_lines(desc, _width(2), 8), 1), 8, 4, 15)
    # T-12H warning + clarifications
    r += 1
    warn = ("T-12H PREP is a preparation check at the caterer only – loading is confirmed solely by the "
            "UPLIFT physical on-board check.")
    _put(ws, r, 1, warn, 5, font=_font(8, True, "C00000"))
    _put(ws, r, 6, "Clarifications open", 6, **lab)
    _put(ws, r, 7, f"={FL('AS')}", 8, font=_font(10, True), align=AL_CEN)
    ws.row_dimensions[r].height = _height(_lines(warn, _width(1, 5), 8, True), 8, 4, 15)
    # legend + completed
    r += 1
    _put(ws, r, 1, "Record state legend", 1, **lab)
    for c, (t, fill) in zip(range(2, 6), [("COMPLETE", GREEN_FILL), ("OPEN", AMBER_FILL), ("N/A", GREY_FILL),
                                          ("OVERDUE / FAIL / INVALID", RED_FILL)]):
        _put(ws, r, c, t, c, font=_font(8, True), fill=fill, align=AL_CEN)
    _put(ws, r, 6, "Completed / in scope", 6, **lab)
    _put(ws, r, 7, f'={FL("AV")}&" / "&{FL("AW")}', 8, font=_font(10, True), align=AL_CEN)
    ws.row_dimensions[r].height = 15
    # round-trip / completion note + as-of
    r += 1
    rt_note = (f'=IF(TRIM({FL("W")})<>"","ROUND-TRIP LEG: return catering is loaded at "&{FL("V")}&" on "&'
               f'{FL("W")}&" before that flight departs, so T-24H and T-12H PREP are pulled forward to before that '
               f'loading window (they may share one due time). UPLIFT is still confirmed on board at "&{FL("G")}&".",'
               f'"Completion % = completed / in-scope checks (n/a when nothing is in scope).")')
    _put(ws, r, 1, rt_note, 5, font=_font(8, bool(f["round_trip"]), "7F4F00" if f["round_trip"] else "595959",
                                         italic=not f["round_trip"]))
    if f["round_trip"]:
        ws.cell(r, 1).fill = AMBER_FILL
    _put(ws, r, 6, "Status as of (UTC)", 6, **lab)
    _put(ws, r, 7, f"={_t('Settings!$B$6')}", 8, font=_font(FS), align=AL_CEN)
    rt_sz = ("ROUND-TRIP LEG: return catering is loaded at KUL on F10 MH1140 before that flight departs, so T-24H "
             "and T-12H PREP are pulled forward to before that loading window (they may share one due time). "
             "UPLIFT is still confirmed on board at PEN.")
    ws.row_dimensions[r].height = _height(_lines(rt_sz if f["round_trip"] else "x" * 80, _width(1, 5), 8, True),
                                          8, 4, 15)

    # ---------------- PIC reference documents
    r += 1
    _band(ws, r, "PIC REFERENCE DOCUMENTS (register: Documents sheet)")
    r += 1
    for (c1, c2), t in zip([(1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 8)],
                           ["Document", "Doc no", "Revision / rev date", "Status", "Attachment (link / location)",
                            "Flag"]):
        _put(ws, r, c1, t, c2, font=_font(8, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    ws.row_dimensions[r].height = _height(2, 8, 3)
    doc_rows = {}
    for name, (no, rev, rdate, att, st) in [("Galley Loading Diagram", ("F", "G", "H", "I", "J")),
                                            ("Menu Checklist", ("K", "L", "M", "N", "O"))]:
        r += 1
        _put(ws, r, 1, name, 1, **lab)
        _put(ws, r, 2, _blank(DC(no)), 2)
        _put(ws, r, 3, f'=IF({DC(rev)}="","",{DC(rev)})&IF(ISNUMBER({DC(rdate)}),CHAR(10)&TEXT({DC(rdate)},"DD-MMM-YY"),"")',
             3, align=AL_CEN)
        _put(ws, r, 4, _blank(DC(st)), 4, font=_font(8, True), align=AL_CEN)
        _put(ws, r, 5, _blank(DC(att)), 5, font=_font(8))
        _put(ws, r, 6, f'=IF({DC(st)}="OUTSTANDING","OUTSTANDING – attach before T-7D (due "&'
                       f'{_t(FL("AG"))}&" "&{FL("V")}&")",IF({DC(st)}="","Status not recorded in Documents sheet",'
                       f'"On file – see attachment location"))', 8, font=_font(FS, True))
        ws.row_dimensions[r].height = _height(3, 8, 4)
        doc_rows[name] = r
        for rng in (f"D{r}", f"F{r}:H{r}"):
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$D${r}="OUTSTANDING"'], fill=RED_STRONG,
                                                           font=Font(color="FFFFFF", bold=True), stopIfTrue=True))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$D${r}="ON FILE"'], fill=GREEN_FILL))
    r += 1
    _put(ws, r, 1, "Notes", 1, **lab)
    _put(ws, r, 2, _blank(DC("Q")), NCOL, font=_font(8))
    ws.row_dimensions[r].height = _height(2, 8, 3)

    # ---------------- attachment spaces (cols 1-2, 3-5) + sign-off (cols 6-8)
    r += 1
    top = r
    box_rows = 7
    for name, c1, c2 in [("Galley Loading Diagram", 1, 2), ("Menu Checklist", 3, 5)]:
        dr = doc_rows[name]
        _put(ws, r, c1, f'=IF($D${dr}="OUTSTANDING","⚠ OUTSTANDING – attach before T-7D | ","")'
                        f'&"ATTACHMENT SPACE – attach / insert {name} here (Insert > Picture or Object)"',
             c2, font=_font(8, True, NAVY), fill=F_BAND, border=False)
        ws.conditional_formatting.add(f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}",
                                      FormulaRule(formula=[f'$D${dr}="OUTSTANDING"'], fill=RED_STRONG,
                                                  font=Font(color="FFFFFF", bold=True)))
        body = r + 1
        ws.cell(body, c1, f'=IF($D${dr}="OUTSTANDING","Placeholder – document OUTSTANDING. Insert the current '
                          f'revision here before T-7D.","On file: "&$B${dr}&" "&$C${dr})')
        ws.cell(body, c1).font = _font(8, False, "7F7F7F", italic=True)
        ws.cell(body, c1).alignment = AL_CEN
        _merge(ws, body, c1, top + box_rows - 1, c2)
        _box(ws, top, c1, top + box_rows - 1, c2, side=Side(style="medium", color=NAVY))
    # sign-off block stacked in cols 6-8
    _put(ws, top, 6, "SIGN-OFF", NCOL, font=_font(9, True, "FFFFFF"), fill=F_NAVY)
    so = [("Flight PIC", f'=IF({FL("Y")}="","",{FL("Y")})'), ("Verifier", None),
          ("Station catering officer", f'="Station: "&{FL("V")}')]
    rr = top + 1
    for role, name in so:
        _put(ws, rr, 6, role, 6, **lab)
        _put(ws, rr, 7, name if name else "Name:", 8,
             font=_font(8, False, "000000" if name else "7F7F7F", italic=not name))
        _put(ws, rr + 1, 6, "Signature / date-time", 6, font=_font(8, False, "7F7F7F", italic=True))
        _put(ws, rr + 1, 7, None, 8)
        rr += 2
    for i in range(box_rows):
        ws.row_dimensions[top + i].height = 18 if i % 2 == 0 else 24
    ws.row_dimensions[top].height = _height(2, 8, 3)
    r = top + box_rows - 1

    # ---------------- checklist table
    r += 1
    head_row = r
    for i, (h, w) in enumerate(COLS, 1):
        _put(ws, r, i, h, font=_font(8, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    ws.row_dimensions[r].height = _height(max(_lines(h, w, 8, True) for h, w in COLS), 8, 3)

    by_cp = {cp: [] for cp, *_ in CHECKPOINTS}
    for gi, c in checks_idx:
        by_cp.setdefault(c["checkpoint"], []).append((gi, c))
    first_data = None
    for cp, desc, due_col, pct in CHECKPOINTS:
        r += 1
        if cp == "UPLIFT":
            due = f'"STD "&{_t(FL("I"))}&" "&{FL("G")}&" (departure stn)"'
        else:
            due = f'{_t(FL(due_col))}&" "&{FL("V")}&" (meal uplift stn)"'
        pct_txt = f'IF(ISNUMBER({FL(pct)}),TEXT({FL(pct)},"0%"),{FL(pct)})'
        _band(ws, r, f'="{cp} — {desc}   |   Checkpoint due "&{due}&"; rows with a different due show their own'
                     f'   |   Complete: "&{pct_txt}&"   |   {len(by_cp[cp])} line(s)"', 9, fill=F_BAND, color=NAVY,
              lines=2 if len(desc) > 40 else 1)
        ws.row_dimensions[r].height = _height(_lines(
            f"{cp} — {desc}   |   Checkpoint due STD 08-Oct-26 21:35 LHR (meal uplift stn); rows with a different due "
            f"show their own   |   Complete: 100%   |   49 line(s)", _width(1, NCOL), 9, True), 9, 5)
        if not by_cp[cp]:
            r += 1
            _put(ws, r, 1, "No check lines at this checkpoint for this flight.", NCOL, font=_font(8, italic=True))
            ws.row_dimensions[r].height = 16
            continue
        for gi, c in by_cp[cp]:
            r += 1
            first_data = first_data or r
            k = FIRST_CHECK_ROW + gi
            C = lambda col: f"Checks!${col}${k}"  # noqa: E731
            vals = [
                f'={C("A")}&CHAR(10)&IF({C("I")}="","",{C("I")}&": ")&{C("J")}',
                f'={C("K")}&IF({C("O")}="","",IF({C("K")}="","",CHAR(10))&"Note: "&{C("O")})',
                f'={C("N")}&CHAR(10)&{_t(C("R"))}&" "&{C("P")}',
                f'=IF({C("U")}="","",{C("U")})&IF({C("AB")}="","",CHAR(10)&"CA: "&{C("AB")})',
                (f'=IF({C("V")}="","",{C("V")})'
                 f'&IF({C("AC")}="","",IF({C("V")}="","",CHAR(10))&"N/A just.: "&{C("AC")})'
                 f'&IF({C("W")}="","",CHAR(10)&"Batch: "&{C("W")})'
                 f'&IF(AND({C("X")}="",{C("Y")}=""),"",CHAR(10)&"Exp: "&{C("X")}&"  /  Act: "&{C("Y")})'),
                (f'=IF({C("Z")}="","","Evidence: "&{C("Z")})'
                 f'&IF({C("AA")}="","",IF({C("Z")}="","",CHAR(10))&"CA: "&{C("AA")})'),
                (f'=IF(ISNUMBER({C("AD")}),TEXT({C("AD")},"{TFMT}")&" "&{C("P")},"")'
                 f'&IF({C("AE")}="","",CHAR(10)&{C("AE")})'),
                _blank(C("AH")),
            ]
            for i, v in enumerate(vals, 1):
                cell = ws.cell(r, i, v)
                cell.font = _font(8 if i == 8 else FS, bold=(i == 8))
                cell.alignment = AL_TOP
                cell.border = BORDER
            item_txt = c["check_id"] + "\n" + (c["category"] + ": " if c["category"] else "") + c["item"]
            req_txt = "\n".join(t for t in (c["expected"], "Note: " + c["note"] if c["note"] else "") if t)
            app = "Clarification required" if c["applic"] != "Required" else "Required"
            exp_line = f"Exp: {c.get('exp_qty') or ''}  /  Act: 9999" if c.get("exp_qty") else ""
            nl = max(
                _lines(item_txt, COLS[0][1]),
                _lines(req_txt, COLS[1][1]),
                _lines(app + "\n08-Oct-26 21:35 LHR", COLS[2][1]),
                ENTRY_LINES + (_lines(exp_line, COLS[4][1]) - 1 if exp_line else 0),
            )
            ws.row_dimensions[r].height = _height(nl, FS, 5)
    last_row = r

    if first_data:
        rs = f"{C_REC}{first_data}:{C_REC}{last_row}"
        tl = f"${C_REC}{first_data}"
        for formula, fill, color in [
            (f'LEFT({tl},8)="COMPLETE"', GREEN_FILL, "006100"),
            (f'OR(LEFT({tl},7)="OVERDUE",LEFT({tl},4)="FAIL",LEFT({tl},7)="INVALID")', RED_FILL, "9C0006"),
            (f'LEFT({tl},4)="OPEN"', AMBER_FILL, "9C5700"),
            (f'LEFT({tl},3)="N/A"', GREY_FILL, "404040"),
        ]:
            ws.conditional_formatting.add(rs, FormulaRule(formula=[formula], fill=fill,
                                                          font=Font(color=color, bold=True), stopIfTrue=True))
        ws.conditional_formatting.add(f"A{first_data}:G{last_row}",
                                      FormulaRule(formula=[f'LEFT({tl},3)="N/A"'], font=Font(color="808080")))
        ws.conditional_formatting.add(f"D{first_data}:D{last_row}",
                                      FormulaRule(formula=[f'LEFT($D{first_data},4)="Fail"'], fill=RED_FILL,
                                                  font=Font(color="9C0006", bold=True)))

    # ---------------- print setup
    ws.print_area = f"A1:{LASTCOL}{last_row}"
    ws.print_title_rows = f"{head_row}:{head_row}"
    ps = ws.page_setup
    ps.orientation = "landscape"
    ps.paperSize = ws.PAPERSIZE_A4
    ps.fitToWidth = 1
    ps.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.page_margins.header = ws.page_margins.footer = 0.2
    for part, text, font in [(ws.oddHeader.left, "MAGCS – Skytrax 2026 catering readiness", "Arial"),
                             (ws.oddHeader.right, flight_tag, "Arial,Bold"),
                             (ws.oddFooter.left, flight_tag, "Arial,Bold"),
                             (ws.oddFooter.center, "Printed: ____________________", "Arial"),
                             (ws.oddFooter.right, f"{f['id']}  Page &P of &N", "Arial")]:
        part.text, part.size, part.font = text, 8, font
    return ws


def build(wb, data):
    idx = {}
    for gi, c in enumerate(data["checks"]):
        idx.setdefault(c["flight_id"], []).append((gi, c))
    for n, f in enumerate(data["flights"], 1):
        _build_one(wb, n, f, idx.get(f["id"], []))
