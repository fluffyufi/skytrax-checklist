"""Dashboard sheet for the MAGCS Skytrax 2026 catering readiness workbook.

Every figure on this sheet is a formula that reads Flights / Checks /
Documents / Settings (see build/CONTRACT.md). Nothing is computed in Python
except layout (row heights, the Checks last-row number).
"""
import math

from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

SHEET = "Dashboard"
FONT = "Arial"

NAVY = "1F3864"
NAVY_MID = "2F5597"
NAVY_LIGHT = "D9E1F2"
GREY = "F2F2F2"
GREY_MID = "D9D9D9"
GREY_TEXT = "595959"
RED = "C00000"
RED_FILL = "F8CBAD"
GREEN_FILL = "C6EFCE"
GREEN_TEXT = "006100"
AMBER_FILL = "FFE699"
AMBER_TEXT = "7F6000"
BAR = "8EA9DB"

FIRST_FLIGHT_ROW, LAST_FLIGHT_ROW = 5, 26
CHECKPOINTS = [
    ("T-7D", "T-7D", "Preparation checks due 7 days before STD"),
    ("T-24H", "T-24H", "Preparation checks due 24 hours before STD"),
    ("T-12H PREP", "T-12H prep", "Preparation checks due 12 hours before STD"),
    ("UPLIFT", "Uplift", "Physical uplift / on-board confirmation, due at STD"),
]
TOP_N = 40

# Column grid (A = margin / row #; B..R = content; T = hidden helper)
WIDTHS = {
    "A": 4, "B": 6, "C": 10, "D": 11, "E": 11, "F": 7, "G": 10, "H": 16,
    "I": 9, "J": 9, "K": 9, "L": 9, "M": 9, "N": 9, "O": 11, "P": 11,
    "Q": 11, "R": 30, "S": 2, "T": 8,
}
LAST_COL = "R"

thin = Side(style="thin", color=GREY_MID)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def _font(size=9, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def _fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def _width(cols):
    """Total width (char units) of a list of column letters."""
    return sum(WIDTHS[c] for c in cols)


def _span(a, b):
    ia, ib = _col_idx(a), _col_idx(b)
    return [get_column_letter(i) for i in range(ia, ib + 1)]


def _col_idx(c):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(c)


def _lines(text, width, size=9):
    """Estimated wrapped line count for text in a cell `width` chars wide."""
    usable = max(width - 1.5, 1) * 1.25 * 9.0 / size
    return sum(max(1, math.ceil(len(p) / usable)) for p in str(text).split("\n"))


def _height(lines, size=9):
    return round(lines * size * 1.32 + 5, 1)


class _Sheet:
    def __init__(self, ws):
        self.ws = ws

    def put(self, ref, value, *, size=9, bold=False, color="000000", italic=False,
            fill=None, h="left", v="center", wrap=False, fmt=None, border=False,
            merge_to=None):
        ws = self.ws
        c = ws[ref]
        c.value = value
        c.font = _font(size, bold, color, italic)
        c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap,
                                indent=(1 if h in ("left", "right") else 0))
        if fmt:
            c.number_format = fmt
        cells = [c]
        if merge_to:
            ws.merge_cells(f"{ref}:{merge_to}")
            cells = [x for row in ws[f"{ref}:{merge_to}"] for x in row]
        for x in cells:
            if fill:
                x.fill = _fill(fill)
            if border:
                x.border = BORDER
        return c

    def section(self, row, title, note=None):
        if note is None:
            self.put(f"B{row}", title, size=11, bold=True, color="FFFFFF", fill=NAVY_MID,
                     merge_to=f"{LAST_COL}{row}")
        else:  # title left, formula note right-aligned in the same band
            self.put(f"B{row}", title, size=11, bold=True, color="FFFFFF",
                     fill=NAVY_MID, merge_to=f"K{row}")
            self.put(f"L{row}", note, size=9, italic=True, color="FFFFFF",
                     fill=NAVY_MID, h="right", merge_to=f"{LAST_COL}{row}")
        self.ws.row_dimensions[row].height = 20

    def header(self, row, spans, height=None):
        """spans: list of (first_col, last_col, text)."""
        max_lines = 1
        for a, b, text in spans:
            self.put(f"{a}{row}", text, bold=True, color="FFFFFF", fill=NAVY,
                     h="center", wrap=True, border=True,
                     merge_to=(f"{b}{row}" if a != b else None))
            max_lines = max(max_lines, _lines(text, _width(_span(a, b)) - 1))
        self.ws.row_dimensions[row].height = height or _height(max_lines)


def build(wb, data):
    ws = wb.create_sheet(SHEET)
    S = _Sheet(ws)
    n_checks = len(data["checks"])
    CL = 5 + n_checks - 1  # last Checks data row

    def ck(col):
        return f"Checks!${col}$5:${col}${CL}"

    def fl(col):
        return f"Flights!${col}${FIRST_FLIGHT_ROW}:${col}${LAST_FLIGHT_ROW}"

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["T"].hidden = True

    # ------------------------------------------------------------ title
    S.put("B1", "MAGCS Inflight Catering – Skytrax 2026 Readiness Dashboard",
          size=16, bold=True, color="FFFFFF", fill=NAVY, merge_to=f"{LAST_COL}1")
    ws.row_dimensions[1].height = 32
    S.put("B2", "As of (UTC)", bold=True, color=GREY_TEXT, fill=GREY, merge_to="C2")
    S.put("D2", "=AsOfUTC", bold=True, color=NAVY, fill=GREY, fmt="dd-mmm-yyyy hh:mm",
          merge_to="E2")
    S.put("F2", "As of (MYT, UTC+8)", bold=True, color=GREY_TEXT, fill=GREY, merge_to="G2")
    S.put("H2", "=AsOfUTC+8/24", bold=True, color=NAVY, fill=GREY, fmt="dd-mmm-yyyy hh:mm")
    S.put("I2", '=IF(Settings!$B$4<>"","As-of time is FIXED by the override in Settings!B4",'
                '"Live clock (Settings!B4 blank) – recalculate (F9) to refresh")'
                '&"  •  All figures are formulas over Flights / Checks / Documents."',
          italic=True, color=GREY_TEXT, fill=GREY, merge_to=f"{LAST_COL}2")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 6

    # ------------------------------------------------------------ KPI tiles
    S.section(4, "Key indicators")
    tiles = [
        ("B", "C", "Flights in programme", f'=COUNTA({fl("A")})', "0", "legs on the Skytrax agenda", False),
        ("D", "E", "Flights READY", f'=COUNTIF({fl("AT")},"READY")', "0",
         "=\"of \"&COUNTA(" + fl("A") + ")&\" flights\"", False),
        ("F", "G", "Overall completion", f'=IF(SUM({ck("AI")})=0,"n/a",SUM({ck("AJ")})/SUM({ck("AI")}))',
         "0.0%", f'=SUM({ck("AJ")})&" of "&SUM({ck("AI")})&" in-scope checks"', False),
        ("H", "I", "Overdue actions", f'=SUM({ck("AK")})', "0", "past due, not complete", True),
        ("J", "K", "Open discrepancies", f'=SUM({ck("AL")})', "0", "fail / variance, CA not closed", True),
        ("L", "M", "Documents outstanding", f'=SUM({fl("AR")})', "0", "GLD + menu checklist", True),
        ("N", "O", "Clarifications open", f'=SUM({ck("AN")})', "0", "awaiting reference answer", True),
        ("P", "R", "Invalid entries", f'=SUM({ck("AM")})', "0", "entries failing validation rules", True),
    ]
    for a, b, label, formula, fmt, sub, alarm in tiles:
        S.put(f"{a}5", label, bold=True, color=NAVY, fill=NAVY_LIGHT, h="center", wrap=True,
              merge_to=f"{b}5")
        S.put(f"{a}6", formula, size=20, bold=True, color=NAVY, fill=GREY, h="center",
              fmt=fmt, merge_to=f"{b}6")
        S.put(f"{a}7", sub, size=8, italic=True, color=GREY_TEXT, fill=GREY, h="center",
              wrap=True, merge_to=f"{b}7")
        if alarm:
            ws.conditional_formatting.add(
                f"{a}6:{b}6", CellIsRule(operator="greaterThan", formula=["0"],
                                         font=Font(name=FONT, color=RED, bold=True)))
    ws.conditional_formatting.add(
        "F6:G6", FormulaRule(formula=['AND(ISNUMBER($F$6),$F$6>=1)'],
                             font=Font(name=FONT, color=GREEN_TEXT, bold=True)))
    ws.row_dimensions[5].height = _height(2)
    ws.row_dimensions[6].height = 32
    ws.row_dimensions[7].height = _height(2, 8)
    ws.row_dimensions[8].height = 8

    # ------------------------------------------------------------ checkpoint summary
    r = 9
    S.section(r, "Checkpoint summary")
    r += 1
    S.header(r, [("B", "C", "Checkpoint"), ("D", "D", "In scope"), ("E", "E", "Complete"),
                 ("F", "G", "Justified N/A (excluded)"), ("H", "H", "Open"),
                 ("I", "I", "Overdue"), ("J", "K", "Completion %"),
                 ("L", LAST_COL, "What is checked")])
    first_cp = r + 1
    for key, label, desc in CHECKPOINTS:
        r += 1
        crit = f'{ck("G")},"{key}"'
        S.put(f"B{r}", label, bold=True, color=NAVY, border=True, merge_to=f"C{r}")
        S.put(f"D{r}", f"=SUMIFS({ck('AI')},{crit})", h="center", fmt="0", border=True)
        S.put(f"E{r}", f"=SUMIFS({ck('AJ')},{crit})", h="center", fmt="0", border=True)
        S.put(f"F{r}", f'=COUNTIFS({crit},{ck("AH")},"N/A – JUSTIFIED*")'
                       f'+COUNTIFS({crit},{ck("AH")},"N/A – RULE*")',
              h="center", fmt="0", border=True, merge_to=f"G{r}")
        S.put(f"H{r}", f"=D{r}-E{r}", h="center", fmt="0", border=True)
        S.put(f"I{r}", f"=SUMIFS({ck('AK')},{crit})", h="center", fmt="0", border=True)
        S.put(f"J{r}", f'=IF(D{r}=0,"n/a",E{r}/D{r})', h="center", fmt="0.0%", border=True,
              merge_to=f"K{r}")
        S.put(f"L{r}", desc, color=GREY_TEXT, border=True, merge_to=f"{LAST_COL}{r}")
        ws.row_dimensions[r].height = 16
    last_cp = r
    r += 1
    S.put(f"B{r}", "All checkpoints", bold=True, color=NAVY, fill=NAVY_LIGHT, border=True,
          merge_to=f"C{r}")
    for col in ("D", "E", "F", "H", "I"):
        S.put(f"{col}{r}", f"=SUM({col}{first_cp}:{col}{last_cp})", bold=True, h="center",
              fmt="0", fill=NAVY_LIGHT, border=True, merge_to=(f"G{r}" if col == "F" else None))
    S.put(f"J{r}", f'=IF(D{r}=0,"n/a",E{r}/D{r})', bold=True, h="center", fmt="0.0%",
          fill=NAVY_LIGHT, border=True, merge_to=f"K{r}")
    S.put(f"L{r}", "Open = in scope − complete (includes overdue).", italic=True,
          color=GREY_TEXT, fill=NAVY_LIGHT, border=True, merge_to=f"{LAST_COL}{r}")
    ws.row_dimensions[r].height = 16
    ws.conditional_formatting.add(
        f"J{first_cp}:J{r}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                         end_value=1, color=BAR, showValue=True))
    ws.conditional_formatting.add(
        f"I{first_cp}:I{r}", CellIsRule(operator="greaterThan", formula=["0"],
                                        font=Font(name=FONT, color=RED, bold=True)))
    r += 1
    ws.row_dimensions[r].height = 8

    # ------------------------------------------------------------ per-flight readiness
    r += 1
    S.section(r, "Per-flight readiness")
    r += 1
    cols = [
        ("B", "Itin", "B", "0"), ("C", "Flight", "F", None), ("D", "Date", "D", "dd-mmm-yy"),
        ("E", "Sector", None, None), ("F", "Class", "K", None), ("G", "Fleet", "L", None),
        ("H", "STD local", "I", "dd-mmm-yy hh:mm"),
        ("I", "T-7D %", "AJ", "0%"), ("J", "T-24H %", "AK", "0%"),
        ("K", "T-12H prep %", "AL", "0%"), ("L", "Uplift %", "AM", "0%"),
        ("M", "Overall %", "AN", "0%"), ("N", "Overdue", "AP", "0"),
        ("O", "Open discrep.", "AQ", "0"), ("P", "Docs outstanding", "AR", "0"),
        ("Q", "Clarifications open", "AS", "0"), ("R", "Readiness", "AT", None),
    ]
    S.header(r, [(c, c, t) for c, t, _, _ in cols])
    f0 = r + 1
    for i in range(LAST_FLIGHT_ROW - FIRST_FLIGHT_ROW + 1):
        r += 1
        src = FIRST_FLIGHT_ROW + i
        band = GREY if i % 2 else None
        for c, _t, fcol, fmt in cols:
            if fcol is None:  # sector = Dep-Arr
                val = f'=IF(Flights!$G${src}="","",Flights!$G${src}&"-"&Flights!$H${src})'
            else:
                ref = f"Flights!${fcol}${src}"
                val = f'=IF({ref}="","",{ref})'
            S.put(f"{c}{r}", val, h=("left" if c in ("C", "E", "G", "R") else "center"),
                  bold=(c in ("C", "R")), fmt=fmt, fill=band, border=True)
        ws.row_dimensions[r].height = 16
    f1 = r
    pct = f"I{f0}:M{f1}"
    ws.conditional_formatting.add(
        pct, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                         color=BAR, showValue=True))
    ws.conditional_formatting.add(
        f"N{f0}:Q{f1}", CellIsRule(operator="greaterThan", formula=["0"],
                                   font=Font(name=FONT, color=RED, bold=True)))
    rd = f"R{f0}:R{f1}"
    ws.conditional_formatting.add(rd, FormulaRule(
        formula=[f'$R{f0}="READY"'], fill=_fill(GREEN_FILL),
        font=Font(name=FONT, color=GREEN_TEXT, bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(rd, FormulaRule(
        formula=[f'LEFT($R{f0},9)="NOT READY"'], fill=_fill(RED_FILL),
        font=Font(name=FONT, color=RED, bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(rd, FormulaRule(
        formula=[f'LEFT($R{f0},9)="PREP DONE"'], fill=_fill(AMBER_FILL),
        font=Font(name=FONT, color=AMBER_TEXT, bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(rd, FormulaRule(
        formula=[f'OR($R{f0}="IN PROGRESS",$R{f0}="NOT STARTED")'], fill=_fill(GREY_MID),
        font=Font(name=FONT, color=GREY_TEXT, bold=True), stopIfTrue=True))
    r += 1
    S.put(f"B{r}", "Percentages show “n/a” where a checkpoint has no in-scope checks. "
                   "Counts above zero are shown in red. Source: Flights sheet (rows 5–26).",
          size=8, italic=True, color=GREY_TEXT, merge_to=f"{LAST_COL}{r}")
    ws.row_dimensions[r].height = 14
    r += 1
    ws.row_dimensions[r].height = 8

    # ------------------------------------------------------------ top-N lists
    def top_list(r, title, total_expr, seq_col, spans, noun):
        """spans: (first_col, last_col, header, checks_col, kind) kind: text|num|date."""
        S.section(r, title,
                  note=f'="Showing "&MIN({TOP_N},{total_expr})&" of "&{total_expr}'
                       f'&" (ordered as in Checks)"')
        r += 1
        S.header(r, [("A", "A", "#")] + [(a, b, h) for a, b, h, _, _ in spans])
        ws[f"T{r}"].value = "Checks row"
        ws[f"T{r}"].font = _font(8, color=GREY_TEXT)
        # height: the widest-wrapping column decides (text lengths from data where known)
        max_lines = 2
        for a, b, _h, _c, kind in spans:
            w = _width(_span(a, b))
            if kind == "item":
                longest = max(len(c["item"]) for c in data["checks"])
                max_lines = max(max_lines, _lines("x" * longest, w))
            elif kind == "long":
                max_lines = max(max_lines, 3)
        max_lines = min(max_lines, 4)
        for k in range(1, TOP_N + 1):
            r += 1
            band = GREY if k % 2 == 0 else None
            ws[f"T{r}"].value = f"=IFERROR(MATCH({k},{ck(seq_col)},0),\"\")"
            ws[f"T{r}"].font = _font(8, color=GREY_TEXT)
            S.put(f"A{r}", f'=IF($T{r}="","",{k})', size=8, color=GREY_TEXT, h="center",
                  v="top", fill=band, border=True)
            for a, b, _h, ccol, kind in spans:
                rng = ck(ccol)
                if kind in ("num", "date"):
                    val = (f'=IF($T{r}="","",IF(INDEX({rng},$T{r})="","",'
                           f'INDEX({rng},$T{r})))')
                else:
                    val = f'=IF($T{r}="","",INDEX({rng},$T{r})&"")'
                fmt = {"date": "dd-mmm-yy hh:mm", "num": "#,##0;-#,##0;0"}.get(kind)
                S.put(f"{a}{r}", val, h=("center" if kind in ("num", "date", "short") else "left"),
                      v="top", wrap=True, fmt=fmt, fill=band, border=True,
                      merge_to=(f"{b}{r}" if a != b else None))
            ws.row_dimensions[r].height = _height(max_lines)
        r += 1
        S.put(f"B{r}", f'=IF({total_expr}=0,"No {noun} at the as-of time.",'
                       f'IF({total_expr}>{TOP_N},"+"&({total_expr}-{TOP_N})&" more {noun} '
                       f'not shown – see Checks sheet",""))',
              bold=True, italic=True, color=GREEN_TEXT, merge_to=f"{LAST_COL}{r}")
        ws.conditional_formatting.add(f"B{r}", FormulaRule(
            formula=[f'LEFT($B${r},1)="+"'], font=Font(name=FONT, color=RED, bold=True)))
        ws.row_dimensions[r].height = 15
        return r

    r += 1
    r = top_list(
        r, f"Overdue actions (top {TOP_N})", f"SUM({ck('AK')})", "AO",
        [("B", "C", "CheckID", "A", "short"), ("D", "D", "Flight No", "C", "short"),
         ("E", "E", "Checkpoint", "G", "short"), ("F", "K", "Check item", "J", "item"),
         ("L", "M", "Due local (check stn)", "R", "date"), ("N", "N", "Check stn", "P", "short"),
         ("O", "P", "PIC", "T", "text"), ("Q", "R", "Record state", "AH", "text")],
        "overdue actions")
    r += 1
    ws.row_dimensions[r].height = 8
    r += 1
    r = top_list(
        r, f"Outstanding discrepancies (top {TOP_N})", f"SUM({ck('AL')})", "AP",
        [("B", "C", "CheckID", "A", "short"), ("D", "D", "Flight No", "C", "short"),
         ("E", "E", "Checkpoint", "G", "short"), ("F", "H", "Check item", "J", "item"),
         ("I", "I", "Status", "U", "short"), ("J", "J", "Expected qty", "X", "num"),
         ("K", "K", "Actual qty", "Y", "num"), ("L", "L", "Variance", "AF", "num"),
         ("M", "P", "Corrective action", "AA", "long"), ("Q", "Q", "CA status", "AB", "short"),
         ("R", "R", "Record state", "AH", "text")],
        "outstanding discrepancies")
    r += 1
    ws.row_dimensions[r].height = 8

    # ------------------------------------------------------------ legend
    r += 1
    S.section(r, "How the figures are calculated")
    legend = [
        ("Blank checks", "A check line with no Status / completion entry never counts as complete. "
                         "It stays OPEN until its due time, then becomes OVERDUE."),
        ("N/A", "N/A only removes a check from the denominator when it is either rule-based "
                "(applicability “N/A – rule”, e.g. fleet-specific items) or JUSTIFIED: "
                "Status N/A with a written N/A justification AND a named verifier. An unjustified "
                "or unverified N/A is flagged INVALID and still counts as open."),
        ("Complete", "Status Pass or Pass after CA with a valid completion time and the fields the "
                     "item requires. Entries that break the validation rules (e.g. completed outside "
                     "the valid window, missing mandatory data) are shown INVALID and do not count."),
        ("Discrepancy", "Status Fail, or an actual quantity that differs from the expected quantity, "
                        "counts as an open discrepancy until a corrective action is recorded and its CA "
                        "status is Closed."),
        ("READY", "A flight is READY only when every in-scope check at all four checkpoints is "
                  "complete – including the physical uplift confirmation – with zero open "
                  "discrepancies, the GLD and menu checklist on file (Documents sheet) and no invalid "
                  "entries. PREP DONE – AWAITING UPLIFT means all preparation checks are complete "
                  "but the physical uplift is not yet confirmed."),
        ("Time basis", "Due times are computed in UTC from STD and shown in local time at the check "
                       "station. As-of time comes from Settings (override in B4, else the PC clock "
                       "adjusted by the offset in B5)."),
    ]
    lab_w = _width(["B", "C"])
    txt_w = _width(_span("D", LAST_COL))
    for label, text in legend:
        r += 1
        S.put(f"B{r}", label, bold=True, color=NAVY, fill=NAVY_LIGHT, v="top", border=True,
              wrap=True, merge_to=f"C{r}")
        S.put(f"D{r}", text, v="top", wrap=True, border=True, merge_to=f"{LAST_COL}{r}")
        ws.row_dimensions[r].height = _height(max(_lines(text, txt_w), _lines(label, lab_w)))
    last_row = r

    # ------------------------------------------------------------ view / print
    ws.freeze_panes = "A4"
    ws.print_area = f"A1:{LAST_COL}{last_row}"
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.left.text = "MAGCS – Skytrax 2026 catering readiness – Dashboard"
    ws.oddFooter.right.text = "Page &P of &N"
    ws.sheet_properties.tabColor = NAVY
    return ws
