# Workbook build contract (shared by all builder modules)

Output: `output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx`
Entry point: `python build/build.py` → loads `build/data.json` (from `build/data.py`), creates an
openpyxl Workbook, calls each module's `build(wb, data)` in order core → dashboard → printable,
reorders sheets, saves, then runs LibreOffice recalc.

Each module: `build/<name>.py` exposing `build(wb, data) -> None`. Modules only ADD their own sheets.
Font: Arial everywhere. Header fill navy `1F3864` with white bold text. Input cells: fill `FFF2CC`
(light yellow). Formula/locked cells: no fill or light grey `F2F2F2`. Never hardcode results —
every status / % / count / due time is a formula that references the sheets below.

Sheet order: Instructions, Dashboard, Flights, Checks, Documents, Requirements, Settings, P01 … P22.

Formula rules: Excel-2007 functions only (SUMIFS, COUNTIFS, INDEX, MATCH, IFERROR, SUMPRODUCT).
No XLOOKUP/FILTER/SORT/UNIQUE/LET/dynamic arrays. `_xlfn.` prefix needed for MAXIFS/MINIFS/TEXTJOIN/IFS.
Quote sheet names with spaces. Workbook is recalculated by LibreOffice, then opened in Excel.

## Settings (owner: core)
| Cell | Meaning | Named range |
|---|---|---|
| B4 | As-of override, UTC datetime (blank = live clock) — input | |
| B5 | This PC's clock offset from UTC in hours (default 8 = MYT) — input | |
| B6 | Effective as-of UTC `=IF(B4<>"",B4,NOW()-B5/24)` | `AsOfUTC` |
| B7 | Physical-uplift window: hours before STD when on-board confirmation may start (default 6) | `UpliftWindowH` |
| B8 | T-24H earliest valid completion: hours before due (default 24) | `T24EarlyH` |
| B9 | T-12H prep earliest valid completion: hours before due (default 12) | `T12EarlyH` |
Station table header row 12, data rows 13-22: A Station, B IANA zone, C Std offset h, D DST offset h,
E DST start UTC, F DST end UTC. Named: `TZ_STN`=A13:A22, `TZ_STD`=C13:C22, `TZ_DST`=D13:D22,
`TZ_DSTS`=E13:E22, `TZ_DSTE`=F13:F22.
Lists (H column onward): `L_Status` = Not started, In progress, Pass, Pass after CA, Fail, N/A;
`L_CA` = Open, Closed.

## Flights (owner: core) — header row 4, data rows 5-26 (F01 = row 5 … F22 = row 26)
A FlightID · B Itin · C Seq · D Date · E Day · F Flight No · G Dep · H Arr · I STD local (datetime) ·
J STA local (datetime) · K Class assessed · L Fleet (schedule; editable) · M Aircraft ref type ·
N Tail/Reg (input) · O Seat · P Transit · Q Remark · R Sector key · S Region · T Service (ref) ·
U Ref uplift stns · V Meal uplift stn · W Round-trip: loaded on flight · X Caterer (input) ·
Y Flight PIC (input) · Z Dep UTC offset · AA STD UTC · AB Block time (h) ·
AC T-7D due UTC · AD T-24H due UTC · AE T-12H prep due UTC · AF Uplift due UTC (=STD UTC) ·
AG T-7D due local · AH T-24H due local · AI T-12H prep due local (all local = meal uplift stn) ·
AJ T-7D % · AK T-24H % · AL T-12H prep % · AM Uplift % · AN Overall % (text "n/a" if denominator 0) ·
AO Open required checks · AP Overdue · AQ Open discrepancies · AR Docs outstanding (0-2) ·
AS Clarifications open · AT Readiness (text) · AU Invalid entries · AV Checks completed · AW Checks in scope ·
AX Inbound carrying flight KUL departure UTC (input; "n/a" where not applicable) · AY KUL loading deadline for KUL-sourced items (UTC)
Readiness values: `READY`, `NOT READY – OVERDUE`, `NOT READY – DISCREPANCY`,
`NOT READY – INVALID ENTRY`, `NOT READY – DOCUMENTS OUTSTANDING`, `NOT READY – CLARIFICATION OPEN`, `PREP DONE – AWAITING UPLIFT`, `IN PROGRESS`, `NOT STARTED`.

## Checks (owner: core) — header row 4, data rows 5 … (one row per check line, in data.json order)
A CheckID · B FlightID · C Flight No · D Date · E Sector (DEP-ARR) · F Class · G Checkpoint
(`T-7D`,`T-24H`,`T-12H PREP`,`UPLIFT`) · H Check type (Preparation / Physical uplift) · I Category ·
J Check item · K Requirement / expected · L Source reference · M Item uplift stn · N Applicability
(`Required`, `Clarification required`, `N/A – rule`) · O Rule note / clarification · P Check station ·
Q Due UTC · R Due local @ check station · S Earliest valid UTC ·
INPUTS: T PIC · U Status · V Result / assessment · W Batch ID · X Expected qty · Y Actual qty ·
Z Evidence ref · AA Corrective action · AB CA status · AC N/A justification · AD Completion time
(local @ check station) · AE Verifier ·
COMPUTED: AF Qty variance · AG Completion UTC · AH Record state (text) · AI In scope (1/0) ·
AJ Complete (1/0) · AK Overdue (1/0) · AL Open discrepancy (1/0) · AM Invalid (1/0) ·
AN Clarification open (1/0) · AO Overdue seq (1..n or "") · AP Discrepancy seq (1..n or "") ·
AQ Req batch · AR Req qty · AS Req doc (1 = GLD, 2 = menu checklist, 3 = carrying flight on Flights AX) · AT N/A permitted (hidden)
Record state values start with one of: `COMPLETE`, `COMPLETE – LATE`, `N/A – RULE`, `N/A – JUSTIFIED`,
`OPEN`, `OPEN – CLARIFICATION`, `OVERDUE`, `FAIL – DISCREPANCY`, `INVALID – <reason>`.
The last data row number is `5 + len(data["checks"]) - 1`; use `len(data["checks"])` — do not hardcode.
Rows of one flight are contiguous and ordered T-7D, T-24H, T-12H PREP, UPLIFT.

## Documents (owner: core) — header row 4, rows 5-26 per flight (same order as Flights)
A FlightID · B Flight No · C Date · D Sector · E Fleet · F GLD doc no · G GLD revision · H GLD rev date ·
I GLD attachment (link / location) · J GLD status (`ON FILE` / `OUTSTANDING`) · K Menu checklist doc no ·
L revision · M rev date · N attachment · O Menu checklist status · P Outstanding count · Q Notes
ISOP revision register below (rows 30+).

## Dashboard (owner: dashboard builder) — see prompt.
## P01..P22 printable flight checklists (owner: printable builder) — see prompt.
