# MAGCS Skytrax 2026 – Inflight Catering Readiness Workbook

**Deliverable:** `output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx`

It covers all 22 legs of the *Skytrax Agenda 2026*, across 3 itineraries. Requirements are applied from `STD UPLIFT INFORMATION.xlsx` by sector, aircraft, assessed cabin class and uplift station.

## Quick start
1. **Settings.** Leave B4 blank to use the live clock. Check that B5 is your PC's offset from UTC; it is 8 for Malaysia.
2. **Flights.** Caterers are pre-filled (column X). Enter the flight PIC (Y) and, once Engineering confirms it 48 h before departure, the tail (N, dropdown):
   - B737-8: 9M-MVO / MVP / MVQ / MVR · A333: 9M-MTJ / MTM / MTG · A350: 9M-MAB … MAH. No flight can be READY until its tail is entered and in the group.
   - Return legs get their double-loaded KUL items on the **pair flight**. Column AX is pre-filled with its scheduled KUL departure (UTC); correct it if the rotation changes.
3. **Documents.** Paste each flight's GLD link and menu checklist link (SharePoint / OneDrive URL, network path, file name, or a hyperlink whose text names the document). The status turns **ON FILE**; a flight cannot be READY while either is OUTSTANDING.
4. **Checks.** This is the only place to record results; use the yellow columns T–AE.
   - Record the PIC, a result, evidence, the completion time and a *different* verifier. On clarification lines also choose the Outcome (column AC). Enter the completion time in local time at the check station shown in column P.
   - **Evidence is two fields.** Pick the *Evidence type* from the list in column S, then enter only the record's own ID in column Z: one token such as `SF-2210`, `IMG_2231`, `DN88213` or a 5+ digit number such as a seal number. A link, network path or file name is accepted only for type *Link / file path*. Free text, flight numbers, dates, times and revisions ("Rev3") are rejected; describe what you found in Result instead.
   - Several IDs in one cell, dummy values (`00000`, `XX-0000`) and links under a non-link type are rejected too.
   - PIC and verifier must be named people, and different ones. A role or title alone ("QA", "PASB supervisor", "Station Manager KUL") is rejected.
   - A plain *Pass* whose result describes a problem ("2 meals short", "menu cards not delivered") is rejected: record *Fail*, then *Pass after CA* with the corrective action.
   - T-24H checks also need a batch ID. Quantity lines need expected and actual quantities.
   - Set the Status last. Column AH shows whether the record is accepted.
5. **Dashboard.** Shows completion for each checkpoint, overdue actions, outstanding discrepancies and each flight's readiness.
6. **P01–P22.** One printable A4 pack per flight: a cover sheet with due times, the list of evidence types, the attachments register and sign-off, then the checklist from page 2. Each check row has "Type:" and "ID:" lines to fill in by hand. Print each flight's sheet on its own so the page numbers run per flight.

## Rules built in
- **Due times.** T-7D, T-24H and T-12H are measured back from STD in UTC and shown in local time at the check station. This handles the LHR switch from BST to GMT on 25 Oct and Adelaide's summer time (ACDT) from 4 Oct.
- **What never counts:** a blank check, a placeholder or "not yet on file" phrase ('-', 'TBC', 'photo to follow', 'done', …), a completion time in the future, before its valid window, after loading or after departure, or a Pass with an unresolved quantity variance or an open corrective action.
- **T-12H is split in two.** The preparation check at the caterer (T-12H PREP) never confirms loading. The physical uplift check (UPLIFT) is valid only inside the uplift window before STD, and only after any carrying flight could have arrived.
- **N/A** is allowed only for printed menu cards on refreshment-only flights (Outcome "Refreshment service – no printed menu card", with a named PIC and a different named verifier). Every other line must be checked.
- **Double-loaded from KUL (MAGCS).** Toiletry kits (BC only), sales carts, compendium (DAM cart) and meal cart covers are loaded at KUL for both legs. On a return leg they travel on the pair flight, so their preparation must be finished before it leaves KUL, and a "KUL-sourced items loaded at KUL" line is confirmed at KUL.
- **Results** need at least two words saying what was checked and found; stock phrases ("All good", "As per menu") are rejected. A justified N/A is removed from both the numerator and the denominator of the completion rate.
- **READY** requires every in-scope check to be complete, including the physical uplift. It also needs zero discrepancies, zero invalid entries, both documents on file and every clarification resolved.

## Decisions applied (MAGCS, 1 Oct 2026)
- Toiletry kits: BC only (EY none), double-loaded from KUL.
- Table cloth and bread linen: all sectors except refreshment / 0.5 meal.
- EY blanket on MH0072 KUL-HKG: included, 100% (one per EY seat).
- MH0003 LHR 11:00 is not a red-eye: no pajamas.
- Sales carts on return legs: double-loaded from KUL on the pair flight.
- New on every flight: compendium in the DAM cart (wide body 2, narrow body 1) and meal cart covers (wide body 8 = 3 BC + 5 EY; narrow body 5 = 2 BC + 3 EY), double-loaded from KUL.
- Pair flights (aircraft rotation, KUL departure local): MH0004 08-Oct 09:50 → MH0001; MH0721 10-Oct 13:45 → MH0720; MH0088 14-Oct 23:30 → MH0089; MH0002 08-Oct 23:20 → MH0003; MH0752 11-Oct 09:35 → MH0753; MH1148 12-Oct 14:05 → MH1149; MH0139 13-Oct 22:25 → MH0138; MH0002 24-Oct 23:20 → MH0003; MH0072 27-Oct 09:10 → MH0073; MH0052 28-Oct 22:25 → MH0053; MH1436 30-Oct 13:10 → MH1437.

## Rebuild
```
python build/data.py      # derive data.json from the schedule + reference workbook
python build/build.py     # build + recalculate output/…xlsx (LibreOffice)
python build/test_evidence.py; python build/test_ready.py                      # evidence IDs / people; F02 readiness cases
python build/test_logic2.py; for r in more r6 r7 r17 r18 r19 r24; do python build/test_logic2.py $r; done
python build/test_r20.py; python build/test_r20.py r22   # N/A timing, KUL lines, tail quantities
```
