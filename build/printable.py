"""Printable per-flight checklists P01..P22 (owner: printable builder).

Each sheet is a formula VIEW over Flights / Checks / Documents / Settings; no
result is stored here. Static text from data.json is used only to size row
heights (LibreOffice does not auto-fit rows of openpyxl-written files).

Paper-first layout: 8 wide columns (~150 width units) so A4 landscape
fit-to-width prints at ~95-100 % with a 9 pt body; every check row leaves room
for hand-written or typed entries of ~140 characters in the combined cells.
"""
import re
from datetime import datetime

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

from core import station_questions

NAVY = "1F3864"
FONT = "Arial"
FS = 9  # body font size

# ------------------------------------------------------------------ layout
# Total 143 width units: A4 landscape, 0.25" margins -> fit-to-width scale ~97 % (measured in LibreOffice).
COLS = [
    ("Check ID / item", 18),
    ("Requirement / expected", 19),
    ("Applic. / due (local)", 12),
    ("Status (CA)", 10),
    ("Result  |  batch  |  exp / act qty", 31),
    ("Corrective action", 22),
    ("Done (local) / verifier / PIC", 18),
    ("Record state", 13)  # fits "CLARIFICATION" in bold 8.5 pt without a mid-word break,
]
NCOL = len(COLS)
LASTCOL = get_column_letter(NCOL)
C_REC = get_column_letter(8)
REF_SHEETS = ["AIRCRAFT TYPE", "CATERING UPLIFT STN", "AMENITIES", "F&B LINEN", "SEAT LINEN", "SALES CART",
              "Signature Drinks"]

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

LINE_PT = {8: 10.0, 8.5: 10.5, 9: 11.0, 10: 13.0, 11: 14.5, 12: 15.5, 13: 17.0}
LS = 8.5  # label size
BLANK = "__________________"
AX_WARN = ("CARRYING FLIGHT NOT YET ENTERED IN THE WORKBOOK: its timing is assumed, so on-board (UPLIFT) checks "
           "before the assumed arrival show INVALID – before uplift window.")
CA_BLANK = "__________________"
W_BLANK = "________________"
Q_BLANK = "______"


def _width(c1, c2=None):
    return sum(COLS[i - 1][1] for i in range(c1, (c2 or c1) + 1))


def _lines(text, width, size=FS, bold=False):
    """Conservative wrapped line count for Arial text in `width` column units."""
    if text is None or text == "":
        return 1
    per_char = 1.15 * (9.0 / size)  # measured: Arial 9pt ~1.24 chars per width unit in print; small margin
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




# sample entries used to reserve room for hand-written / typed inputs
_RESULT_SAMPLE = ("Panel of 4 tasted chicken rendang and nasi lemak; texture, seasoning and temperature within spec; "
                  "plating matched the menu photo ok.")  # 140 chars
_CA_SAMPLE = ("Sauce re-seasoned and re-tasted; batch re-labelled; caterer supervisor informed; re-check logged "
              "at 20:40 by QA.")  # ~115 chars


_SHEET = r"(?:AIRCRAFT TYPE|CATERING UPLIFT STN|AMENITIES|F&B LINEN|SEAT LINEN|SALES CART|Signature Drinks)"
_RNG = r"![A-Z]{1,2}\d+(?::[A-Z]{1,2}\d+)?"
# Paper wording: spreadsheet cell addresses of the reference workbook mean nothing to a PIC on paper.
_PAPER = [
    (r"\b([a-z]+) -(,| \(| \||$)", r"\1 n/a\2"),  # galley data "-" = none; anchored so other dashes survive
    (r"Per note (" + _SHEET + ")" + _RNG + r"\.", r"Per the \1 reference note."),
    (r"note (" + _SHEET + ")" + _RNG, r"the \1 reference note"),
    (r" \((?:note )?(" + _SHEET + ")" + _RNG + r"\)", ""),
    (r" \(note [A-Z]{1,2}\d+\)", ""),
    (r"note [A-Z]{1,2}\d+: ", "reference note: "),
    (r"A350 row AMENITIES![A-Z]{1,2}\d+", "The A350 row of the AMENITIES reference"),
    (r"\(general row [A-Z]{1,2}\d+: uplift stn [A-Z]{1,2}\d+ = ", "(general row: uplift stn "),
    (r"but [A-Z]{1,2}\d+ excludes", "but the reference excludes"),
    (r" \(Flights A[XY]\)", ""),
    (r"entered on Flights col AX", "recorded in the workbook (Flights sheet)"),
    (r"until Flights AX holds", "until the workbook holds"),
    (r" Applicability follows the Fleet cell on Flights\.", ""),
    (r"in Flights AX", "in the workbook (Flights sheet)"),
    (r"^Documents sheet row complete$", "Recorded on the Documents sheet; copy stapled as A1 / A2"),
]


def _paper_pairs(texts):
    """Literal (old, new) substitutions that turn reference-workbook cell addresses into plain words."""
    pairs = []
    for t in texts:
        for pat, rep in _PAPER:
            for m in re.finditer(pat, t or ""):
                pr = (m.group(0), m.expand(rep))
                if pr not in pairs:
                    pairs.append(pr)
    return pairs


def _paperify(expr, texts):
    for old, new in _paper_pairs(texts):
        expr = f'SUBSTITUTE({expr},"{old.replace(chr(34), chr(34) * 2)}","{new}")'
    return expr


def _paper_static(t):
    for old, new in _paper_pairs([t]):
        t = t.replace(old, new)
    return t


# Printable body height in sheet points. Measured LibreOffice break points put the real capacity at
# ~530-540 pt (row-height rounding varies); every page break is set manually at a conservative 522 pt, so the
# engine's own breaks never fire and the simulation cannot drift (Excel's capacity is larger still).
PAGE_BODY_PT = 522.0


def _keep_bands_with_rows(ws, head_row, band_rows, last_row):
    """Paginate explicitly: cover (page 1) ends before the table header; each later page is filled up to
    PAGE_BODY_PT (including the repeated header) and a checkpoint band is always kept with its first row."""
    h = lambda r: ws.row_dimensions[r].height or 15  # noqa: E731
    title_h = h(head_row)
    bands = set(band_rows)
    used = 0.0
    for r in range(1, last_row + 1):
        if r < head_row:
            continue  # cover sheet: fixed content, checked by render
        if r == head_row:
            ws.row_breaks.append(Break(id=r - 1))
            used = h(r)
            continue
        need = h(r) + (h(r + 1) if r in bands and r + 1 <= last_row else 0)
        if used + need > PAGE_BODY_PT and used > title_h + 1:
            ws.row_breaks.append(Break(id=r - 1))
            used = title_h
        used += h(r)


# ------------------------------------------------------------------ builder
def _build_one(wb, n, f, checks_idx, carry_ids=()):
    fr = FIRST_FLIGHT_ROW + n - 1
    FL = lambda col: f"Flights!${col}${fr}"  # noqa: E731
    DC = lambda col: f"Documents!${col}${fr}"  # noqa: E731
    ws = wb.create_sheet(sheet_title(n, f))
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for i, (_, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    lab = dict(font=_font(LS, True), fill=F_LABEL)
    d_text = datetime.strptime(f["date"], "%Y-%m-%d").strftime("%d-%b-%y")
    flight_tag = f"{f['id']} {f['flt']} {d_text} {f['dep']}-{f['arr']} Class {f['cls']}"
    fill_in = lambda ref: f'=IF(TRIM({ref})="","{BLANK}",{ref})'  # noqa: E731

    # ---------------- title
    r = 1
    _put(ws, r, 1,
         f'="SKYTRAX 2026 CATERING READINESS CHECKLIST  |  {f["id"]}  "&{FL("F")}&"  "&TEXT({FL("D")},"DD-MMM-YY")'
         f'&" ("&{FL("E")}&")  "&{FL("G")}&" → "&{FL("H")}&"  |  Class "&{FL("K")}',
         NCOL, font=_font(12, True, "FFFFFF"), fill=F_NAVY)
    ws.row_dimensions[r].height = 20

    # ---------------- two blocks sharing rows: flight key/value (A:D) | due & readiness (E:H)
    left = [
        ("Flight / date", f'={FL("F")}&"   "&TEXT({FL("D")},"DD-MMM-YY")&" ("&{FL("E")}&")"', "MH0001 08-Oct-26 (THU)"),
        ("Sector / class / seat", f'={FL("G")}&" → "&{FL("H")}&"   Class "&{FL("K")}&"   Seat "&{FL("O")}', "x" * 30),
        ("STD → STA (local)", f'={_t(FL("I"))}&" "&{FL("G")}&"  →  "&{_t(FL("J"))}&" "&{FL("H")}',
         "08-Oct-26 21:35 LHR  →  09-Oct-26 17:50 KUL"),
        ("Fleet / aircraft ref", f'={FL("L")}&" – "&{FL("M")}', f"{f['fleet']} – {f['fleet_ref']}"),
        ("Tail / Reg", fill_in(FL("N")), BLANK),
        ("Service type", f"={FL('T')}", f["service"]),
        ("Meal uplift stn", f'={FL("V")}&"   (reference uplift stns: "&{FL("U")}&")"',
         f"{f['meal_uplift_stn']}   (reference uplift stns: {f['ref_uplift_stns']})"),
        ("Transit / remark", f'=TRIM({FL("P")}&"  "&{FL("Q")})', f"{f['transit']}  {f['remark']}"),
        ("Caterer", fill_in(FL("X")), "x" * 40),
        ("Flight PIC", fill_in(FL("Y")), "x" * 40),
    ]
    top = r + 1
    for i, (label, formula, txt) in enumerate(left):
        rr = top + i
        _put(ws, rr, 1, label, 1, **lab)
        _put(ws, rr, 2, formula, 4, font=_font(FS))
    # right block
    rr = top
    _put(ws, rr, 5, "READINESS STATUS", 5, font=_font(LS, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    _put(ws, rr, 6, f"={FL('AT')}", 8, font=_font(10, True), align=AL_CEN)
    rd = f"$F${rr}"
    for formula, fill, font in [(f'LEFT({rd},5)="READY"', GREEN_FILL, None),
                                (f'LEFT({rd},9)="NOT READY"', RED_FILL, Font(color="9C0006", bold=True)),
                                (f'{rd}<>""', AMBER_FILL, None)]:
        ws.conditional_formatting.add(f"F{rr}:H{rr}", FormulaRule(formula=[formula], fill=fill, font=font,
                                                                 stopIfTrue=True))
    rr += 1
    for c1, c2, t in [(5, 5, "Checkpoint"), (6, 6, "Due (local)"), (7, 8, "% complete")]:
        _put(ws, rr, c1, t, c2, font=_font(LS, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    due_txt = {
        "T-7D": f'={_t(FL("AG"))}&" "&{FL("V")}',
        "T-24H": f'={_t(FL("AH"))}&" "&{FL("V")}',
        "T-12H PREP": f'={_t(FL("AI"))}&" "&{FL("V")}&IF({FL("AI")}={FL("AH")},"  (= T-24H due)","")',
        "UPLIFT": f'="STD "&{_t(FL("I"))}&" "&{FL("G")}&" (dep stn)"',
    }
    cp_label = {"T-7D": "T-7D – 7 days before departure",
                "T-24H": "T-24H – production batch sensory",
                "T-12H PREP": "T-12H PREP – caterer prep only",
                "UPLIFT": "UPLIFT – on-board confirmation"}
    right_h = {}
    for cp, desc, _due, pct in CHECKPOINTS + [("OVERALL", "", None, "AN")]:
        rr += 1
        _put(ws, rr, 5, cp_label.get(cp, "OVERALL (all checkpoints)"), 5, **lab)
        _put(ws, rr, 6, due_txt.get(cp, "–"), 6, font=_font(FS, cp != "OVERALL"), align=AL_CEN)
        _put(ws, rr, 7, f"={FL(pct)}", 8, font=_font(FS, True), fmt="0%", align=AL_CEN)
        due_sample = {"UPLIFT": "STD 08-Oct-26 21:35 LHR (dep stn)",
                      "T-12H PREP": "08-Oct-26 09:35 LHR  (= T-24H due)"}.get(cp, "08-Oct-26 09:35 LHR")
        right_h[rr] = max(_lines(cp_label.get(cp, ""), COLS[4][1], LS, True), _lines(due_sample, COLS[5][1], FS, True))
    # counts: three per row, each value in its own cell for local conditional formatting
    for label, cols, red, names in [("Overdue / discrep. / invalid", ("AP", "AQ", "AU"), True,
                                     ("Overdue", "Discrep.", "Invalid")),
                                    ("Open reqd / clarif. / docs out", ("AO", "AS", "AR"), None,
                                     ("Open reqd", "Clarif.", "Docs out"))]:
        rr += 1
        right_h[rr] = _lines(label, COLS[4][1], LS, True)
        _put(ws, rr, 5, label, 5, **lab)
        for c, col, nm in zip((6, 7, 8), cols, names):
            cell = _put(ws, rr, c, f"={FL(col)}", c, font=_font(10, True), align=AL_CEN, fmt=f'"{nm} "0')
            if red or col == "AR":
                cc = cell.coordinate
                ws.conditional_formatting.add(cc, FormulaRule(formula=[f"AND(ISNUMBER({cc}),{cc}>0)"],
                                                              fill=RED_FILL, font=Font(color="9C0006", bold=True)))
    rr += 1
    _put(ws, rr, 5, "Done / in scope | as of UTC", 5, **lab)
    _put(ws, rr, 6, f'={FL("AV")}&" / "&{FL("AW")}', 6, font=_font(10, True), align=AL_CEN)
    _put(ws, rr, 7, f"={_t('Settings!$B$6')}", 8, font=_font(8), align=AL_CEN)
    assert rr == top + len(left) - 1
    for i, (label, formula, txt) in enumerate(left):
        row = top + i
        nl = max(_lines(txt, _width(2, 4)), _lines(label, COLS[0][1], LS, True), right_h.get(row, 1))
        write_in = label in ("Tail / Reg", "Caterer", "Flight PIC")  # hand-written on paper when blank
        ws.row_dimensions[row].height = _height(nl, FS, 2, 21 if write_in else 13)
    r = rr

    # ---------------- notes: T-12H warning / round-trip
    r += 1
    warn = ("T-12H PREP = preparation at the caterer only; loading is confirmed solely by the UPLIFT "
            "on-board check.  Record-state legend →")
    _put(ws, r, 1, warn, 4, font=_font(LS, True, "C00000"))
    for c, (t, fill) in zip(range(5, 9), [("COMPLETE", GREEN_FILL), ("OPEN", AMBER_FILL), ("N/A", GREY_FILL),
                                          ("OVERDUE / FAIL / INVALID", RED_FILL)]):
        _put(ws, r, c, t, c, font=_font(8, True), fill=fill, align=AL_CEN)
    ws.row_dimensions[r].height = _height(max(_lines(warn, _width(1, 4), LS, True),
                                              _lines("OVERDUE / FAIL / INVALID", COLS[7][1], 8, True)), LS, 3)
    carry = f["round_trip"] or f["id"] in carry_ids
    if carry:
        r += 1
        kul = "IFERROR(INDEX(Settings!$C$13:$C$22,MATCH(\"KUL\",Settings!$A$13:$A$22,0)),8)"
        ay = FL("AY")
        est = (f'IF(ISNUMBER({FL("AX")}),"",IF(TRIM({FL("W")})<>""," (agenda candidate – actual not yet entered)",'
               f'" (latest possible – actual not yet entered)"))&IF(ISNUMBER({FL("AX")}),"",CHAR(10)&"{AX_WARN}")')
        times = (f'"KUL loading window opens: "&IF(ISNUMBER({ay}),TEXT({ay}+{kul}/24,"{TFMT}")'
                 f'&" KUL","(not set)")&"   |   Carrying flight leaves KUL: "&'
                 f'IF(ISNUMBER({ay}),TEXT({ay}+UpliftWindowH/24+{kul}/24,"{TFMT}")&" KUL","(not set)")&{est}')
        if f["round_trip"]:
            cand = (f'{FL("W")}&IFERROR(" "&INDEX(Flights!$F${FIRST_FLIGHT_ROW}:$F${FIRST_FLIGHT_ROW + NFL - 1},'
                    f'MATCH(TRIM({FL("W")}),Flights!$A${FIRST_FLIGHT_ROW}:$A${FIRST_FLIGHT_ROW + NFL - 1},0)),"")')
            note = (f'="ROUND-TRIP LEG: all catering is loaded at KUL on the pair flight "&{cand}'
                    f'&"; prep dues are capped at the KUL loading window; on-board checks at "'
                    f'&{FL("G")}&". Each row shows its own due."&CHAR(10)&{times}')
            head = ("ROUND-TRIP LEG: all catering is loaded at KUL on the pair flight MH1148; prep dues are capped at the KUL loading window; on-board checks at PEN. Each row shows its own due.")
        else:
            note = (f'="KUL-SOURCED ITEMS: some items are uplifted at KUL on the inbound carrying flight; '
                    f'their preparation is capped at the KUL loading window. Each row shows its own due."'
                    f'&CHAR(10)&{times}')
            head = ("KUL-SOURCED ITEMS: some items are uplifted at KUL on the inbound carrying flight; "
                    "their preparation is capped at the KUL loading window. Each row shows its own due.")
        _put(ws, r, 1, note, NCOL, font=_font(LS, True, "7F4F00"), fill=AMBER_FILL)
        sz = head + "\n" + ("KUL loading window opens: 12-Oct-26 05:45 KUL   |   Carrying flight "
                            "leaves KUL: 12-Oct-26 11:45 KUL (agenda candidate – actual not yet "
                            "entered)\n" + AX_WARN)
        ws.row_dimensions[r].height = _height(_lines(sz, _width(1, NCOL), LS, True), LS, 4)

    # ---------------- attachments register (paper pack)
    r += 1
    for (c1, c2), t in zip([(1, 1), (2, 4), (5, 5), (6, 8)],
                           ["Attachments", "Link (Documents sheet)", "Status", "Paper pack instruction"]):
        _put(ws, r, c1, t, c2, font=_font(LS, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    ws.row_dimensions[r].height = 14
    for code, name, (att, st) in [("A1", "Galley Loading Diagram", ("F", "G")),
                                  ("A2", "Menu Checklist", ("H", "I"))]:
        r += 1
        _put(ws, r, 1, f"{code}  {name}", 1, **lab)
        _put(ws, r, 2, _blank(DC(att)), 4, font=_font(LS))
        _put(ws, r, 5, f'=IF({DC(st)}="OUTSTANDING","OUTSTANDING – attach before T-7D (due "&{_t(FL("AG"))}'
                       f'&" "&{FL("V")}&")",IF({DC(st)}="","Status not recorded",{DC(st)}))', 5,
             font=_font(FS, True))
        _put(ws, r, 6, f'="Print and staple behind this sheet as page {code}."', 8, font=_font(LS))
        ws.row_dimensions[r].height = _height(
            _lines("OUTSTANDING – attach before T-7D (due 01-Oct-26 21:35 LHR)", COLS[4][1], FS, True), FS, 3)
        ws.conditional_formatting.add(f"E{r}", FormulaRule(formula=[f'LEFT($E${r},11)="OUTSTANDING"'],
                                                           fill=RED_STRONG, font=Font(color="FFFFFF", bold=True),
                                                           stopIfTrue=True))
        ws.conditional_formatting.add(f"E{r}", FormulaRule(formula=[f'$E${r}="ON FILE"'], fill=GREEN_FILL))
    r += 1
    _put(ws, r, 1, "Document notes", 1, **lab)
    _put(ws, r, 2, _blank(DC("K")), NCOL, font=_font(8))
    ws.row_dimensions[r].height = _height(2, 8, 3)  # Documents notes (free text) up to ~2 lines at full width
    ws.cell(r, 1).comment = Comment("Excel only: insert the GLD / menu checklist image (Insert > Picture) in the "
                                    "insertion area after the checklist table; it is not printed.", "MAGCS")

    # ---------------- sign-off (horizontal)
    r += 1
    roles = [("Flight PIC", 1, 2, f'="Name: "&IF(TRIM({FL("Y")})="","{BLANK}",{FL("Y")})'),
             ("Verifier", 3, 5, f'="Name: {BLANK}"'),
             ("Station catering officer", 6, 8, f'="Name: {BLANK}"&"   Station: "&{FL("V")}')]
    for role, c1, c2, _ in roles:
        _put(ws, r, c1, "SIGN-OFF – " + role, c2, font=_font(LS, True, "FFFFFF"), fill=F_NAVY)
    ws.row_dimensions[r].height = 14
    r += 1
    for role, c1, c2, nm in roles:
        _put(ws, r, c1, nm, c2, font=_font(FS))
    ws.row_dimensions[r].height = 18
    r += 1
    for role, c1, c2, _ in roles:
        _put(ws, r, c1, "Signature / date-time:", c2, font=_font(8, False, "7F7F7F", italic=True), align=AL_TOP)
    ws.row_dimensions[r].height = 30

    # ---------------- checklist table
    r += 1
    head_row = r
    for i, (h, w) in enumerate(COLS, 1):
        _put(ws, r, i, h, font=_font(LS, True, "FFFFFF"), fill=F_NAVY, align=AL_CEN)
    ws.row_dimensions[r].height = _height(max(_lines(h, w, LS, True) for h, w in COLS), LS, 3)

    entry_lines = max(
        _lines(_RESULT_SAMPLE + "\nBatch: BATCH-LHR-20261001-JCL-0042\nExp: 280  /  Act: 280", COLS[4][1]),
        _lines("CA: " + _CA_SAMPLE, COLS[5][1]),
        _lines("Done: 01-Oct-26 20:35 LHR\nVerifier: Nurul Izzah Mohd Shahrizal (QA Lead)\n"
               "PIC: Capt. Ahmad Rahman bin Abdullah", COLS[6][1]),
        _lines("[ ] Pass\n[ ] Pass after CA\n[ ] Fail\n[ ] N/A\nCA: Open / Closed", COLS[3][1]),
    ) + 1  # spare line
    band_rows = []
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
        band_rows.append(r)
        pct_txt = f'IF(ISNUMBER({FL(pct)}),TEXT({FL(pct)},"0%"),{FL(pct)})'
        _band(ws, r, f'="{cp} — {desc}  |  Due "&{due}&" (rows with another due show their own)'
                     f'  |  Complete "&{pct_txt}&"  |  {len(by_cp[cp])} lines"', FS, fill=F_BAND, color=NAVY)
        ws.row_dimensions[r].height = _height(_lines(
            f"{cp} — {desc}  |  Due STD 08-Oct-26 21:35 LHR (departure stn) (rows with another due show their own)"
            f"  |  Complete 100%  |  49 lines", _width(1, NCOL), FS, True), FS, 4)
        if not by_cp[cp]:
            r += 1
            _put(ws, r, 1, "No check lines at this checkpoint for this flight.", NCOL, font=_font(LS, italic=True))
            ws.row_dimensions[r].height = 16
            continue
        for gi, c in by_cp[cp]:
            r += 1
            first_data = first_data or r
            k = FIRST_CHECK_ROW + gi
            C = lambda col: f"Checks!${col}${k}"  # noqa: E731
            sqq = sq_by_line.get(c["check_id"])
            req = _paperify(f'{C("K")}&IF({C("O")}="","",IF({C("K")}="","",CHAR(10))&"Note: "&{C("O")})',
                          (c["expected"], c["note"]))
            rna = f'LEFT({C("AH")},10)="N/A – RULE"'  # rule-N/A rows: nothing to write by hand
            vals = [
                "=" + _paperify(f'{C("A")}&CHAR(10)&IF({C("I")}="","",{C("I")}&": ")&{C("J")}', (c["item"],)),
                ("=" + req) if not sqq else
                (f'=IF(TRIM(Settings!$F${sqq["row"]})="",{req},"CONFIRMED uplift stn: "&TRIM(Settings!$F${sqq["row"]})'
                 f'&IF(ISNUMBER(Settings!$G${sqq["row"]}),", carried on the flight leaving "&TRIM(Settings!$F${sqq["row"]})&" "'
                 f'&TEXT(Settings!$G${sqq["row"]},"{TFMT}")&" UTC","")&CHAR(10)&{req})'),
                f'={C("N")}&CHAR(10)&{_t(C("R"))}&" "&{C("P")}',
                (f'=IF(TRIM({C("U")})<>"",{C("U")}&IF({C("AB")}="","",CHAR(10)&"CA: "&{C("AB")}),'
                 f'IF({rna},"","[ ] Pass"&CHAR(10)&"[ ] Pass after CA"&CHAR(10)&"[ ] Fail"'
                 f'&IF({C("AT")}>=1,CHAR(10)&"[ ] N/A","")&CHAR(10)&"CA: Open / Closed"))'),
                (f'=IF({C("V")}="","",{C("V")})'
                 f'&IF({C("AC")}="",IF(OR({rna},AND({C("N")}<>"Clarification required",{C("AT")}=0)),"",'
                 f'IF({C("V")}="","",CHAR(10))&IF({C("AT")}=2,"Outcome if N/A: [ ] refreshment – no menu card",'
                 f'"Outcome: [ ] as listed  [ ] not carried  [ ] n/a to aircraft")),'
                 f'IF({C("V")}="","",CHAR(10))&"Outcome: "&{C("AC")})'
                 f'&IF(OR({C("W")}<>"",AND({C("AQ")}=1,NOT({rna}))),CHAR(10)&"Batch: "&IF({C("W")}="","{W_BLANK}",{C("W")}),"")'
                 + (f'&IF(TRIM(Settings!$F${sqq["row"]})="",CHAR(10)&"Loaded at: [ ] {sqq["dep"]}  [ ] {sqq["alt"]}'
                    f'   if {sqq["alt"]}: flight MH____ dep (UTC) ___-___ __:__","")' if sqq and c["check_type"] == "Preparation" else "") +
                 f'&IF({C("AS")}=3,CHAR(10)&"Carrying flight: "&IF(ISNUMBER({FL("AX")}),"KUL dep "&TEXT({FL("AX")},"{TFMT}")&" UTC",'
                 f'"MH______  KUL dep (UTC): ___-___ __:__"),"")'
                 f'&IF(OR({C("X")}<>"",{C("Y")}<>"",AND({C("AR")}=1,NOT({rna}))),CHAR(10)&"Exp: "'
                 f'&IF({C("X")}="","{Q_BLANK}",{C("X")})&"  /  Act: "&IF({C("Y")}="","{Q_BLANK}",{C("Y")}),"")'),
                f'=IF({C("AA")}<>"","CA: "&{C("AA")},IF({rna},"","CA (if any): "&CHAR(10)&"{CA_BLANK}"))',
                (f'=IF(AND({rna},NOT(ISNUMBER({C("AD")})),{C("AE")}="",{C("T")}=""),"",'
                 f'"Done (dd-mmm hh:mm "&{C("P")}&"): "&IF(ISNUMBER({C("AD")}),TEXT({C("AD")},"{TFMT}"),"{W_BLANK}")'
                 f'&CHAR(10)&"Verifier: "&IF({C("AE")}="","{W_BLANK}",{C("AE")})'
                 f'&CHAR(10)&"PIC: "&IF({C("T")}="","{W_BLANK}",{C("T")}))'),
                _blank(C("AH")),
            ]
            for i, v in enumerate(vals, 1):
                cell = ws.cell(r, i, v)
                cell.font = _font(LS if i == 8 else FS, bold=(i == 8))
                cell.alignment = AL_TOP
                cell.border = BORDER
            item_txt = _paper_static(c["check_id"] + "\n" + (c["category"] + ": " if c["category"] else "") + c["item"])
            req_txt = _paper_static("\n".join(t for t in (c["expected"], "Note: " + c["note"] if c["note"] else "") if t))
            app = "Clarification required" if c["applic"] != "Required" else "Required"
            exp_extra = (_lines(f"Exp: {c['exp_qty']}  /  Act: 9999", COLS[4][1]) - 1) if c.get("exp_qty") else 0
            nl = max(
                _lines(item_txt, COLS[0][1]),
                _lines(req_txt, COLS[1][1]),
                _lines(app + "\n08-Oct-26 21:35 LHR", COLS[2][1]),
                _lines("OPEN – CLARIFICATION", COLS[7][1], LS, True),
                entry_lines + exp_extra,
            )
            ws.row_dimensions[r].height = _height(nl, FS, 4)
    last_table = r
    _keep_bands_with_rows(ws, head_row, band_rows, last_table)

    if first_data:
        rs = f"{C_REC}{first_data}:{C_REC}{last_table}"
        tl = f"${C_REC}{first_data}"
        for formula, fill, color in [
            (f'ISNUMBER(SEARCH("CHECK WORDING",{tl}))', AMBER_FILL, "7F4F00"),
            (f'LEFT({tl},8)="COMPLETE"', GREEN_FILL, "006100"),
            (f'OR(LEFT({tl},7)="OVERDUE",LEFT({tl},4)="FAIL",LEFT({tl},7)="INVALID")', RED_FILL, "9C0006"),
            (f'LEFT({tl},4)="OPEN"', AMBER_FILL, "9C5700"),
            (f'LEFT({tl},3)="N/A"', GREY_FILL, "404040"),
        ]:
            ws.conditional_formatting.add(rs, FormulaRule(formula=[formula], fill=fill,
                                                          font=Font(color=color, bold=True), stopIfTrue=True))
        ws.conditional_formatting.add(f"A{first_data}:G{last_table}",
                                      FormulaRule(formula=[f'LEFT({tl},3)="N/A"'], font=Font(color="808080")))
        ws.conditional_formatting.add(f"D{first_data}:D{last_table}",
                                      FormulaRule(formula=[f'LEFT($D{first_data},4)="Fail"'], fill=RED_FILL,
                                                  font=Font(color="9C0006", bold=True)))

    # ---------------- Excel insertion area (end of sheet, never pushes the table)
    r += 2
    _put(ws, r, 1, "EXCEL INSERTION AREA (on-screen only, not printed) – A1 Galley Loading Diagram: Insert > "
                   "Picture here. On paper: staple behind this sheet as page A1.", 4, font=_font(LS, True, NAVY),
         fill=F_BAND)
    _put(ws, r, 5, "EXCEL INSERTION AREA (on-screen only, not printed) – A2 Menu Checklist: Insert > Picture here. "
                   "On paper: staple behind this sheet as page A2.", 8, font=_font(LS, True, NAVY), fill=F_BAND)
    ws.row_dimensions[r].height = _height(2, LS, 3)
    body = r + 1
    for c1, c2 in ((1, 4), (5, 8)):
        _merge(ws, body, c1, body + 9, c2)
        ws.cell(body, c1).alignment = AL_CEN
        _box(ws, r, c1, body + 9, c2, side=Side(style="medium", color=NAVY))
    for i in range(10):
        ws.row_dimensions[body + i].height = 22

    # ---------------- print setup
    ws.print_area = f"A1:{LASTCOL}{last_table}"  # insertion area stays on screen only
    ws.print_title_rows = f"{head_row}:{head_row}"
    ps = ws.page_setup
    ps.orientation = "landscape"
    ps.paperSize = ws.PAPERSIZE_A4
    ps.fitToWidth = 1
    ps.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.page_margins.header = ws.page_margins.footer = 0.2
    for part, text, font in [(ws.oddHeader.left, "MAGCS – Skytrax 2026 catering readiness", "Arial"),
                             (ws.oddHeader.right, flight_tag, "Arial,Bold"),
                             (ws.oddFooter.left, flight_tag, "Arial,Bold"),
                             (ws.oddFooter.center, "Printed: ____________________", "Arial"),
                             (ws.oddFooter.right, f"{f['id']}  Page &P of &N", "Arial")]:
        part.text, part.size, part.font = text, 8, font
    return ws


def build(wb, data):
    global sq_by_line
    sq_by_line = {}
    for q in station_questions(data):
        sq_by_line[q["prep"]] = q
        sq_by_line[q["upl"]] = q
    idx = {}
    carry_ids = {c["flight_id"] for c in data["checks"] if c.get("due_rule") == "CARRY"}
    for gi, c in enumerate(data["checks"]):
        idx.setdefault(c["flight_id"], []).append((gi, c))
    for n, f in enumerate(data["flights"], 1):
        _build_one(wb, n, f, idx.get(f["id"], []), carry_ids)
