# MAGCS Skytrax 2026 – Inflight Catering Readiness Workbook

**Deliverable:** `output/MAGCS_Skytrax_2026_Catering_Readiness.xlsx`

It covers all 22 legs of the *Skytrax Agenda 2026*, across 3 itineraries. Requirements are applied from `STD UPLIFT INFORMATION.xlsx` by sector, aircraft, assessed cabin class and uplift station.

## Quick start
1. **Settings.** Leave B4 blank to use the live clock. Check that B5 is your PC's offset from UTC; it is 8 for Malaysia.
2. **Flights.** For each leg, enter the tail/registration (N), caterer (X) and flight PIC (Y).
   - The tail matters for A350 legs: 9M-MAH and the A359s differ in blanket quantity and sales-cart position.
   - Six legs carry items loaded at KUL: F05, F07, F11, F13, F17 and F21. For these, enter the carrying flight's KUL departure time in UTC in column AX. Until you do, the workbook assumes the agenda's candidate flight.
3. **Documents.** For each flight, record the galley loading diagram and menu checklist: document number, revision, revision date and file location.
   - None were supplied, so all 44 start as **OUTSTANDING**.
   - A flight cannot be READY while any of them is outstanding.
4. **Checks.** This is the only place to record results; use the yellow columns T–AE.
   - Record the PIC, a result, evidence, the completion time and a *different* verifier. Enter the completion time in local time at the check station shown in column P.
   - T-24H checks also need a batch ID. Quantity lines need expected and actual quantities.
   - Set the Status last. Column AH shows whether the record is accepted.
5. **Dashboard.** Shows completion for each checkpoint, overdue actions, outstanding discrepancies and each flight's readiness.
6. **P01–P22.** One printable A4 pack per flight: a cover sheet with due times, the attachments register and sign-off, then the checklist from page 2. Print each flight's sheet on its own so the page numbers run per flight.

## Rules built in
- **Due times.** T-7D, T-24H and T-12H are measured back from STD in UTC and shown in local time at the check station. This handles the LHR switch from BST to GMT on 25 Oct and Adelaide's summer time (ACDT) from 4 Oct.
- **What never counts:** a blank check, a placeholder ('-', 'TBC', 'n/a', …), a completion time in the future, before its valid window, after loading or after departure, or a Pass with an unresolved quantity variance or an open corrective action.
- **T-12H is split in two.** The preparation check at the caterer (T-12H PREP) never confirms loading. The physical uplift check (UPLIFT) is valid only inside the uplift window before STD, and only after any carrying flight could have arrived.
- **N/A.** Allowed only on clarification items, and on printed menu cards for refreshment-only flights. It needs a real justification and a verifier. A justified N/A is removed from both the numerator and the denominator of the completion rate.
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
python build/test_ready.py; python build/test_logic2.py; python build/test_logic2.py more   # regression tests
```
