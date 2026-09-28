"""Printable per-flight checklists P01..P22 (owner: printable builder).

Each sheet is a formula VIEW over Flights / Checks / Documents / Settings; no
result is stored here. Static text from data.json is used only to size row
heights (LibreOffice does not auto-fit rows of openpyxl-written files).
"""
import math
from datetime import datetime

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break, RowBreak

NAVY = "1F3864"
FONT = "Arial"
FS = 8  # table font size

# ------------------------------------------------------------------ layout
# (header, width, Checks column or None)
COLS = [
    ("Check ID", 10),
    ("Check item", 27),
    ("Requirement / expected", 33),
    ("Applicability", 12),
    ("Due (local @ stn)", 12),
    ("Status (CA status)", 10),
    ("Result / assessment", 18),
    ("Batch ID", 9),
    ("Exp qty", 10),
    ("Act qty", 7),
    ("Evidence ref", 12),
    ("Corrective action", 18),
    ("Completion time (local)", 12),
    ("Verifier", 10),
    ("Record state", 15),
]
NCOL = len(COLS)
LASTCOL = get_column_letter(NCOL)

CHECKPOINTS = [
    ("T-7D", "7 days before departure", "AG", "AJ"),
    ("T-24H", "Production batch sensory", "AH", "AK"),
    ("T-12H PREP", "Preparation check at caterer (does NOT confirm loading)", "AI", "AL"),
    ("UPLIFT", "Physical uplift – on-board confirmation before departure", "I", "AM"),
]

FIRST_CHECK_ROW = 5
FIRST_FLIGHT_ROW = 5

# ------------------------------------------------------------------ styles
thin = Side(style="thin", color="808080")
med = Side(style="medium", color=NAVY)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_NAVY = PatternFill("solid", fgColor=NAVY)
F_BAND = PatternFill("solid", fgColor="D9E1F2")
F_LABEL = PatternFill("solid", fgColor="F2F2F2")
F_BOX = PatternFill("solid", fgColor="FBFBFB")

RED_FILL = PatternFill("solid", fgColor="FFC7CE", bgColor="FFC7CE")
RED_STRONG = PatternFill("solid", fgColor="C00000", bgColor="C00000")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE", bgColor="C6EFCE")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C", bgColor="FFEB9C")
GREY_FILL = PatternFill("solid", fgColor="D9D9D9", bgColor="D9D9D9")

DT_FMT = "dd-mmm-yy hh:mm"
DATE_FMT = "dd-mmm-yyyy (ddd)"
PCT_FMT = "0%"

AL_WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
AL_WRAP_C = Alignment(wrap_text=True, vertical="center", horizontal="left")
AL_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

LINE_PT = {8: 10.5, 9: 12.0, 10: 13.0, 11: 14.5, 14: 18.5}


def _width(c1, c2=None):
    c2 = c2 or c1
    return sum(COLS[i - 1][1] for i in range(c1, c2 + 1))


def _lines(text, width, size=FS, bold=False):
    """Conservative wrapped line count for Arial text in a column of `width` chars."""
    if text is None or text == "":
        return 1
    per_char = 1.5 * (8.0 / size)  # LibreOffice fits ~1.9 Arial-8 chars per width unit; keep margin for Excel
    if bold:
        per_char *= 0.9
    cap = max(1, int(width * per_char) - 1)
    n = 0
    for para in str(text).split("\n"):
        cnt, line = 1, 0
        for w in para.split(" "):
            wl = len(w)
            if wl > cap:  # long token is hard-broken
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
    """Outline + inner grid for a (merged) range."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(
                left=side if c == c1 else None,
                right=side if c == c2 else None,
                top=side if r == r1 else None,
                bottom=side if r == r2 else None,
            )


def _put(ws, r, c1, value, c2=None, font=None, fill=None, align=AL_WRAP_C, fmt=None, border=True):
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


def _band(ws, r, text, size=9, fill=F_NAVY, color="FFFFFF", height=None):
    _put(ws, r, 1, text, NCOL, font=_font(size, True, color), fill=fill,
         align=Alignment(vertical="center", horizontal="left", wrap_text=True))
    ws.row_dimensions[r].height = height or _height(1, size, 6)


def _blank(src):
    return f'=IF({src}="","",{src})'


def sheet_title(n, f):
    d = datetime.strptime(f["date"], "%Y-%m-%d")
    t = f"P{n:02d} {f['flt']} {d.strftime('%d%b').upper()}"
    for ch in '[]:*?/\\':
        t = t.replace(ch, "")
    return t[:31]


# ------------------------------------------------------------------ builder
def _build_one(wb, data, n, f, checks_idx):
    fr = FIRST_FLIGHT_ROW + n - 1
    FL = lambda col: f"Flights!${col}${fr}"  # noqa: E731
    DC = lambda col: f"Documents!${col}${fr}"  # noqa: E731
    ws = wb.create_sheet(sheet_title(n, f))
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for i, (_, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    lab = dict(font=_font(FS, True), fill=F_LABEL)
    val = dict(font=_font(9))

    # ---------------- title
    r = 1
    _put(ws, r, 1, "MAGCS INFLIGHT CATERING – SKYTRAX 2026 PER-FLIGHT READINESS CHECKLIST"
         f"  |  {f['id']}", NCOL, font=_font(14, True, "FFFFFF"), fill=F_NAVY,
         align=Alignment(vertical="center", horizontal="left"))
    ws.row_dimensions[r].height = 26
    r = 2
    _put(ws, r, 1,
         f'={FL("F")}&"   "&TEXT({FL("D")},"DD-MMM-YYYY")&" ("&{FL("E")}&")   "&{FL("G")}&" → "&{FL("H")}'
         f'&"   |   Class assessed: "&{FL("K")}&"   |   Seat "&{FL("O")}&"   |   Fleet "&{FL("L")}',
         NCOL, font=_font(11, True, NAVY), fill=F_BAND,
         align=Alignment(vertical="center", horizontal="left"))
    ws.row_dimensions[r].height = 20
    r = 3
    _put(ws, r, 1,
         "View only – all values are live formulas from Flights / Checks / Documents. Enter data in the "
         "Checks sheet (inputs) or print this sheet and transcribe hand-written entries back into Checks.",
         NCOL, font=_font(FS, False, "595959", italic=True), border=False,
         align=Alignment(vertical="center", horizontal="left", wrap_text=True))
    ws.row_dimensions[r].height = _height(_lines(ws.cell(r, 1).value, _width(1, NCOL)), FS, 3)

    # ---------------- flight details (label/value grid)
    r = 4
    _band(ws, r, "FLIGHT DETAILS")
    grid = [
        [("Flight no", f"={FL('F')}", None), ("Date / day", f"={FL('D')}", DATE_FMT),
         ("Sector (DEP-ARR)", f'={FL("G")}&"-"&{FL("H")}', None), ("Class assessed", f"={FL('K')}", None)],
        [("STD local (dep)", f"={FL('I')}", DT_FMT), ("STA local (arr)", f"={FL('J')}", DT_FMT),
         ("Seat", f"={FL('O')}", None), ("Service type", f"={FL('T')}", None)],
        [("Fleet (schedule)", f"={FL('L')}", None), ("Aircraft ref type", f"={FL('M')}", None),
         ("Transit / remark", f'=TRIM({FL("P")}&"  "&{FL("Q")})', None),
         ("Tail / Reg", f'=IF({FL("N")}="","(enter tail in Flights!N)",{FL("N")})', None)],
        [("Meal uplift station", f"={FL('V')}", None),
         ("Round-trip: loaded on flight", f'=IF({FL("W")}="","– (catered at meal uplift stn)",{FL("W")})', None),
         ("Ref uplift stations", f"={FL('U')}", None),
         ("Caterer", f'=IF({FL("X")}="","(enter in Flights!X)",{FL("X")})', None)],
        [("Flight PIC", f'=IF({FL("Y")}="","(enter in Flights!Y)",{FL("Y")})', None),
         ("Sector key / region", f'={FL("R")}&"  "&{FL("S")}', None),
         ("Status as of (UTC)", "=Settings!$B$6", DT_FMT),
         ("Block time (h)", f"={FL('AB')}", "0.00")],
    ]
    # column spans: (label c1,c2),(value c1,c2) x4
    spans = [((1, 1), (2, 2)), ((3, 3), (4, 6)), ((7, 7), (8, 10)), ((11, 11), (12, 15))]
    static = {  # static text used for height sizing (longest plausible value)
        "Aircraft ref type": f["fleet_ref"], "Service type": f["service"],
        "Ref uplift stations": f["ref_uplift_stns"], "Transit / remark": f"{f['transit']}  {f['remark']}",
        "Sector key / region": f"{f['sector']}  {f['region']}",
        "Round-trip: loaded on flight": "– (catered at meal uplift stn)",
        "Caterer": "(enter in Flights!X) " + "x" * 20, "Flight PIC": "x" * 30,
        "Tail / Reg": "(enter tail in Flights!N)",
    }
    for row in grid:
        r += 1
        nl = 1
        for (label, formula, fmt), ((l1, l2), (v1, v2)) in zip(row, spans):
            _put(ws, r, l1, label, l2, **lab)
            _put(ws, r, v1, formula, v2, fmt=fmt, **val)
            nl = max(nl, _lines(label, _width(l1, l2), FS, True),
                     _lines(static.get(label, "x" * 12), _width(v1, v2), 9))
        ws.row_dimensions[r].height = _height(nl, 9, 5, 18)

    # ---------------- due times / readiness
    r += 1
    _band(ws, r, "DUE TIMES (local @ meal uplift station; uplift = STD local @ departure) & READINESS")
    r += 1
    hdr_r = r
    for (c1, c2), t in zip([(1, 1), (2, 3), (4, 5), (6, 6)],
                           ["Checkpoint", "Purpose", "Due (local)", "% complete"]):
        _put(ws, r, c1, t, c2, font=_font(FS, True, "FFFFFF"), fill=F_NAVY, align=AL_CENTER)
    _put(ws, r, 7, "READINESS STATUS", 7, font=_font(FS, True), fill=F_LABEL, align=AL_CENTER)
    _put(ws, r, 8, f"={FL('AT')}", NCOL, font=_font(11, True), align=AL_CENTER)
    ws.row_dimensions[r].height = 22
    readiness_cell = f"H{r}"
    counts = [
        [("Overdue", "AP"), ("Open discrepancies", "AQ"), ("Open required checks", "AO"),
         ("Invalid entries", "AU")],
        [("Docs outstanding (0-2)", "AR"), ("Clarifications open", "AS"), ("Checks completed", "AV"),
         ("Checks in scope", "AW")],
    ]
    cspans = [((7, 7), (8, 8)), ((9, 10), (11, 11)), ((12, 12), (13, 13)), ((14, 14), (15, 15))]
    count_cells = []
    for i, (cp, desc, due_col, pct_col) in enumerate(CHECKPOINTS + [("OVERALL", "All checkpoints", None, "AN")]):
        r += 1
        bold = cp == "OVERALL"
        _put(ws, r, 1, cp, 1, font=_font(FS, True), fill=F_LABEL)
        _put(ws, r, 2, desc, 3, font=_font(FS, bold))
        if due_col:
            _put(ws, r, 4, f"={FL(due_col)}", 5, font=_font(9, True), fmt=DT_FMT, align=AL_CENTER)
        else:
            _put(ws, r, 4, f"={FL('I')}", 5, font=_font(9), fmt='"STD "' + DT_FMT, align=AL_CENTER)
        _put(ws, r, 6, f"={FL(pct_col)}", 6, font=_font(9, True), fmt=PCT_FMT, align=AL_CENTER)
        if i < 2:
            for (label, col), ((l1, l2), (v1, v2)) in zip(counts[i], cspans):
                _put(ws, r, l1, label, l2, **lab)
                cell = _put(ws, r, v1, f"={FL(col)}", v2, font=_font(10, True), align=AL_CENTER)
                if col in ("AP", "AQ", "AU", "AR"):
                    count_cells.append(cell.coordinate)
        elif i == 2:
            _put(ws, r, 7, "Legend (record state)", 8, **lab)
            _put(ws, r, 9, "COMPLETE", 10, font=_font(FS, True), fill=GREEN_FILL, align=AL_CENTER)
            _put(ws, r, 11, "OPEN", 11, font=_font(FS, True), fill=AMBER_FILL, align=AL_CENTER)
            _put(ws, r, 12, "OVERDUE / FAIL / INVALID", 13, font=_font(FS, True), fill=RED_FILL, align=AL_CENTER)
            _put(ws, r, 14, "N/A", 15, font=_font(FS, True), fill=GREY_FILL, align=AL_CENTER)
        elif i == 3:
            _put(ws, r, 7, "T-12H PREP is a preparation check only – loading is confirmed solely by the "
                           "UPLIFT physical on-board check.", NCOL,
                 font=_font(FS, True, "C00000"), align=AL_WRAP_C)
        ws.row_dimensions[r].height = _height(max(_lines(desc, _width(2, 3)), 1), FS, 6, 18)
    r += 1
    # overall row filled above (i == 4) -> fill remaining cells on that row
    last_due_r = r - 1
    _put(ws, last_due_r, 7, "Completion % per checkpoint = completed / in-scope checks "
                            "(n/a when nothing is in scope).", NCOL,
         font=_font(FS, False, "595959", italic=True), align=AL_WRAP_C)

    # conditional formatting: readiness + counts
    rng = f"H{hdr_r}:{LASTCOL}{hdr_r}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${readiness_cell},5)="READY"'],
                                                   fill=GREEN_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${readiness_cell},9)="NOT READY"'],
                                                   fill=RED_FILL, font=Font(color="9C0006", bold=True),
                                                   stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${readiness_cell}<>""'], fill=AMBER_FILL))
    for cc in count_cells:
        ws.conditional_formatting.add(cc, FormulaRule(formula=[f"AND(ISNUMBER({cc}),{cc}>0)"],
                                                      fill=RED_FILL, font=Font(color="9C0006", bold=True)))

    # ---------------- PIC reference documents
    _band(ws, r, "PIC REFERENCE DOCUMENTS (register: Documents sheet)")
    r += 1
    heads = [((1, 1), "Document"), ((2, 2), "Doc no"), ((3, 3), "Attachment (link / location)"),
             ((4, 4), "Revision"), ((5, 5), "Rev date"), ((6, 6), "Status"), ((7, NCOL), "Flag")]
    for (c1, c2), t in heads:
        _put(ws, r, c1, t, c2, font=_font(FS, True, "FFFFFF"), fill=F_NAVY, align=AL_CENTER)
    ws.row_dimensions[r].height = _height(2, FS, 4)
    doc_rows = {}
    for name, cols in [("Galley Loading Diagram", ("F", "G", "H", "I", "J")),
                       ("Menu Checklist", ("K", "L", "M", "N", "O"))]:
        r += 1
        no, rev, rdate, att, st = cols
        _put(ws, r, 1, name, 1, **lab)
        _put(ws, r, 2, _blank(DC(no)), 2, font=_font(9))
        _put(ws, r, 3, _blank(DC(att)), 3, font=_font(FS))
        _put(ws, r, 4, _blank(DC(rev)), 4, font=_font(9), align=AL_CENTER)
        _put(ws, r, 5, _blank(DC(rdate)), 5, font=_font(9), fmt="dd-mmm-yyyy", align=AL_CENTER)
        _put(ws, r, 6, _blank(DC(st)), 6, font=_font(FS, True), align=AL_CENTER)
        _put(ws, r, 7, f'=IF({DC(st)}="OUTSTANDING","OUTSTANDING – attach before T-7D (due "'
                       f'&TEXT({FL("AG")},"DD-MMM-YYYY HH:MM")&" local)",'
                       f'IF({DC(st)}="","Status not recorded in Documents sheet","On file – see attachment '
                       f'location / space below"))', NCOL, font=_font(9, True), align=AL_WRAP_C)
        ws.row_dimensions[r].height = _height(3, FS, 4)  # attachment path may be long
        doc_rows[name] = (r, f"Documents!${st}${fr}")
        for rng_ in (f"F{r}", f"G{r}:{LASTCOL}{r}"):
            ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'$F{r}="OUTSTANDING"'], fill=RED_STRONG,
                                                            font=Font(color="FFFFFF", bold=True), stopIfTrue=True))
            ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'$F{r}="ON FILE"'], fill=GREEN_FILL))
    r += 1
    _put(ws, r, 1, "Notes", 1, **lab)
    _put(ws, r, 2, _blank(DC("Q")), NCOL, font=_font(FS), align=AL_WRAP_C)
    ws.row_dimensions[r].height = _height(2, FS, 4)

    # attachment spaces (side by side)
    r += 1
    ws.row_dimensions[r].height = 6
    r += 1
    box_top = r
    boxes = [("Galley Loading Diagram", 1, 6), ("Menu Checklist", 7, NCOL)]
    for name, c1, c2 in boxes:
        dr, st_ref = doc_rows[name]
        _put(ws, r, c1,
             f'=IF({st_ref}="OUTSTANDING","⚠ OUTSTANDING – attach before T-7D  |  ","")'
             f'&"ATTACHMENT SPACE – Attach / insert {name} here (Insert > Picture or Object)"',
             c2, font=_font(9, True, NAVY), fill=F_BAND, border=False,
             align=Alignment(wrap_text=True, vertical="center", horizontal="left"))
        ws.conditional_formatting.add(
            f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}",
            FormulaRule(formula=[f'$F${dr}="OUTSTANDING"'], fill=RED_STRONG, font=Font(color="FFFFFF", bold=True)))
        body = r + 1
        ws.cell(body, c1, f'=IF({st_ref}="OUTSTANDING","Document OUTSTANDING – placeholder. '
                          f'Insert the current revision here before T-7D.",'
                          f'"Document on file: "&{DC("F" if name.startswith("Galley") else "K")}&'
                          f'" rev "&{DC("G" if name.startswith("Galley") else "L")})')
        ws.cell(body, c1).font = _font(9, False, "7F7F7F", italic=True)
        ws.cell(body, c1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _merge(ws, body, c1, body + 7, c2)
        for rr in range(r, body + 8):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = Border()
        _box(ws, r, c1, body + 7, c2, side=Side(style="medium", color=NAVY))
    ws.row_dimensions[r].height = _height(2, 9, 4)
    for rr in range(box_top + 1, box_top + 9):
        ws.row_dimensions[rr].height = 20
    r = box_top + 8

    # ---------------- checklist table (new page)
    r += 1
    ws.row_breaks.append(Break(id=r))
    r += 1
    head_row = r
    for i, (h, w) in enumerate(COLS, 1):
        _put(ws, r, i, h, font=_font(FS, True, "FFFFFF"), fill=F_NAVY, align=AL_CENTER)
    ws.row_dimensions[r].height = _height(max(_lines(h, w, FS, True) for h, w in COLS), FS, 4)

    by_cp = {cp: [] for cp, *_ in CHECKPOINTS}
    for gi, c in checks_idx:
        by_cp.setdefault(c["checkpoint"], []).append((gi, c))
    first_data = None
    for cp, desc, due_col, pct_col in CHECKPOINTS:
        r += 1
        due = FL(due_col)
        band_formula = (f'="{cp} — {desc}   |   Due (local): "&IF(ISNUMBER({due}),TEXT({due},"DD-MMM-YYYY HH:MM"),"")'
                        f'&"   |   Complete: "&IF(ISNUMBER({FL(pct_col)}),TEXT({FL(pct_col)},"0%"),{FL(pct_col)})'
                        f'&"   |   {len(by_cp[cp])} check line(s)"')
        _band(ws, r, band_formula, 9, fill=F_BAND, color=NAVY)
        if not by_cp[cp]:
            r += 1
            _put(ws, r, 1, "No check lines at this checkpoint for this flight.", NCOL, font=_font(FS, italic=True))
            ws.row_dimensions[r].height = 16
            continue
        for gi, c in by_cp[cp]:
            r += 1
            first_data = first_data or r
            k = FIRST_CHECK_ROW + gi
            C = lambda col: f"Checks!${col}${k}"  # noqa: E731
            vals = [
                (f"={C('A')}", None, AL_WRAP),
                (f'=IF({C("I")}="","",{C("I")}&": ")&{C("J")}', None, AL_WRAP),
                (f'={C("K")}&IF({C("O")}="","",IF({C("K")}="","",CHAR(10))&"Note: "&{C("O")})', None, AL_WRAP),
                (_blank(C("N")), None, AL_WRAP),
                (f'=IF(ISNUMBER({C("R")}),TEXT({C("R")},"DD-MMM HH:MM")&" "&{C("P")},"")', None, AL_WRAP),
                (f'=IF({C("U")}="","",{C("U")})&IF({C("AB")}="","",CHAR(10)&"CA: "&{C("AB")})', None, AL_WRAP),
                (f'=IF({C("V")}="","",{C("V")})&IF({C("AC")}="","",IF({C("V")}="","",CHAR(10))&"N/A just.: "&{C("AC")})',
                 None, AL_WRAP),
                (_blank(C("W")), None, AL_WRAP),
                (_blank(C("X")), None, AL_WRAP),
                (_blank(C("Y")), None, AL_WRAP),
                (_blank(C("Z")), None, AL_WRAP),
                (_blank(C("AA")), None, AL_WRAP),
                (_blank(C("AD")), "dd-mmm hh:mm", AL_WRAP),
                (_blank(C("AE")), None, AL_WRAP),
                (_blank(C("AH")), None, AL_WRAP),
            ]
            for i, (v, fmt, al) in enumerate(vals, 1):
                cell = ws.cell(r, i, v)
                cell.font = _font(FS, bold=(i in (1, 15)))
                cell.alignment = al
                cell.border = BORDER
                if fmt:
                    cell.number_format = fmt
            item_txt = (c["category"] + ": " if c["category"] else "") + c["item"]
            req_txt = "\n".join(t for t in (c["expected"], "Note: " + c["note"] if c["note"] else "") if t)
            nl = max(
                _lines(c["check_id"], COLS[0][1]),
                _lines(item_txt, COLS[1][1]),
                _lines(req_txt, COLS[2][1]),
                _lines("Clarification required" if c["applic"] != "Required" else "Required", COLS[3][1]),
                _lines("08-Oct 20:35 LHR", COLS[4][1]),
                _lines(str(c.get("exp_qty") or ""), COLS[8][1]),
                _lines("OPEN – CLARIFICATION", COLS[14][1], FS, True),
                2,  # leave room for hand-written entries
            )
            ws.row_dimensions[r].height = _height(nl, FS, 5)
    last_data = r

    # record-state / status conditional formatting
    if first_data:
        rs = f"O{first_data}:O{last_data}"
        tl = f"$O{first_data}"
        for formula, fill, color in [
            (f'LEFT({tl},8)="COMPLETE"', GREEN_FILL, "006100"),
            (f'OR(LEFT({tl},7)="OVERDUE",LEFT({tl},4)="FAIL",LEFT({tl},7)="INVALID")', RED_FILL, "9C0006"),
            (f'LEFT({tl},4)="OPEN"', AMBER_FILL, "9C5700"),
            (f'LEFT({tl},3)="N/A"', GREY_FILL, "404040"),
        ]:
            ws.conditional_formatting.add(rs, FormulaRule(formula=[formula], fill=fill,
                                                          font=Font(color=color, bold=True), stopIfTrue=True))
        # grey the whole line for N/A rows, red text for Fail status
        ws.conditional_formatting.add(f"A{first_data}:N{last_data}",
                                      FormulaRule(formula=[f'LEFT({tl},3)="N/A"'], font=Font(color="808080")))
        ws.conditional_formatting.add(f"F{first_data}:F{last_data}",
                                      FormulaRule(formula=[f'$F{first_data}="Fail"'], fill=RED_FILL,
                                                  font=Font(color="9C0006", bold=True)))

    # ---------------- sign-off footer
    r += 2
    _band(ws, r, "SIGN-OFF")
    r += 1
    stmt = ("I confirm the checks above were performed and recorded at the stated times. T-12H PREP confirms "
            "preparation at the caterer only; loading is confirmed solely by the UPLIFT on-board check "
            "before departure. Any FAIL / discrepancy has a corrective action recorded.")
    _put(ws, r, 1, stmt, NCOL, font=_font(FS, italic=True), align=AL_WRAP_C)
    ws.row_dimensions[r].height = _height(_lines(stmt, _width(1, NCOL)), FS, 5)
    r += 1
    for (c1, c2), t in zip([(1, 2), (3, 5), (6, 10), (11, 12), (13, 15)],
                           ["Role", "Name", "Signature", "Date / time (local)", "Station / remarks"]):
        _put(ws, r, c1, t, c2, font=_font(FS, True, "FFFFFF"), fill=F_NAVY, align=AL_CENTER)
    ws.row_dimensions[r].height = 16
    for role, name in [("Flight PIC", f'=IF({FL("Y")}="","",{FL("Y")})'),
                       ("Verifier", ""),
                       ("Station catering officer", "")]:
        r += 1
        _put(ws, r, 1, role, 2, **lab)
        _put(ws, r, 3, name, 5, font=_font(9))
        _put(ws, r, 6, None, 10)
        _put(ws, r, 11, None, 12)
        _put(ws, r, 13, f'={FL("V")}' if role.startswith("Station") else None, 15, font=_font(9))
        ws.row_dimensions[r].height = 32
    last_row = r

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
    ws.page_margins.top = ws.page_margins.bottom = 0.55
    ws.page_margins.header = ws.page_margins.footer = 0.25
    hdr = f"{f['id']}  {f['flt']}  {f['date']}  {f['dep']}-{f['arr']}  Class {f['cls']}"
    ws.oddHeader.left.text = "MAGCS – Skytrax 2026 catering readiness"
    ws.oddHeader.left.size = 8
    ws.oddHeader.left.font = "Arial"
    ws.oddHeader.right.text = hdr
    ws.oddHeader.right.size = 8
    ws.oddHeader.right.font = "Arial,Bold"
    ws.oddFooter.left.text = f"{f['id']} {f['flt']} {f['date']} – printed &D &T"
    ws.oddFooter.left.size = 8
    ws.oddFooter.left.font = "Arial"
    ws.oddFooter.right.text = "Page &P of &N"
    ws.oddFooter.right.size = 8
    ws.oddFooter.right.font = "Arial"
    return ws


def build(wb, data):
    flights = data["flights"]
    idx = {}
    for gi, c in enumerate(data["checks"]):
        idx.setdefault(c["flight_id"], []).append((gi, c))
    for n, f in enumerate(flights, 1):
        _build_one(wb, data, n, f, idx.get(f["id"], []))
