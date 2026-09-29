# MAGCS Skytrax 2026 – Inflight Catering Readiness Workbook

**Deliverable:** `output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx`

It covers all 22 legs of the *Skytrax Agenda 2026*, across 3 itineraries. Requirements are applied from `STD UPLIFT INFORMATION.xlsx` by sector, aircraft, assessed cabin class and uplift station.

## Quick start
1. **Settings.** Leave B4 blank to use the live clock. Check that B5 is your PC's offset from UTC; it is 8 for Malaysia.
2. **Flights.** For each leg, enter the tail/registration (N), caterer (X) and flight PIC (Y).
   - The tail matters for A350 legs: 9M-MAH and the A359s differ in blanket quantity and sales-cart position. An A350 leg cannot be READY until its tail is entered.
   - Six legs carry items loaded at KUL: F05, F07, F11, F13, F17 and F21. For these, enter the carrying flight's KUL departure time in UTC in column AX. Until you do, the workbook assumes the agenda's candidate flight.
3. **Documents.** For each flight, record the galley loading diagram and menu checklist: document number, revision, revision date and file location.
   - None were supplied, so all 44 start as **OUTSTANDING**.
   - A flight cannot be READY while any of them is outstanding.
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
- **Clarifications and N/A use the Outcome dropdown (Checks column AC).** A clarification line is resolved only by choosing "Confirmed – applies / carried as listed" (status Pass) or "Confirmed – not carried on this sector" / "… not applicable to this aircraft / tail" (status N/A), with the written confirmation cited as evidence type + ID. Printed menu cards on refreshment-only flights may be N/A with "Refreshment service – no printed menu card". N/A is allowed nowhere else and needs a named PIC and a different named verifier.
- **Results** need at least two words saying what was checked and found; stock phrases ("All good", "As per menu") are rejected. A justified N/A is removed from both the numerator and the denominator of the completion rate.
- **READY** requires every in-scope check to be complete, including the physical uplift. It also needs zero discrepancies, zero invalid entries, both documents on file and every clarification resolved.

## Items the reference does not settle
These 44 check lines (preparation and on-board lines counted separately) are marked **Clarification required** and hold readiness until confirmed:
- whether EY receives toiletry kits;
- where the A350 toiletry kits are uplifted on KUL→LHR legs;
- whether table cloth is carried on CGK, HAN and HKG, where the matrix and note B78 disagree;
- whether the HKG EY blanket falls under the regional exclusion;
- whether MH0003 on 9 Oct counts as a red-eye for pajamas;
- where sales carts are loaded on legs returning from outstations;
- the aircraft rotation and carrying flight for the PEN and LGK return legs and for the KUL-sourced items.

## Rebuild
```
python build/data.py      # derive data.json from the schedule + reference workbook
python build/build.py     # build + recalculate output/…xlsx (LibreOffice)
python build/test_evidence.py; python build/test_ready.py; python build/test_logic2.py; python build/test_logic2.py more   # regression tests
```
