# Financial Assistance Transfer Engine

Google Apps Script that automates copying applicant data from a form
spreadsheet into a records spreadsheet. One click instead of manual retyping.

---

## Install (for the worker)

1. Open the source spreadsheet
2. Extensions → Apps Script
3. Click `Code.gs`, press **Ctrl+A**, then **Delete**
4. Paste the full contents of `Code.gs` from this repo
5. Save (💾) and reload the spreadsheet tab
6. The **Transfer** menu appears in the top bar

> To update later: repeat steps 3–6 with the newest version.

---

## First-time setup

1. **Transfer → Add New Route** — a block appears in the Config tab
2. Fill in: ROUTE NAME, FROM SPREADSHEET, FROM TAB, TO SPREADSHEET, TO TAB
3. **Transfer → Refresh Dropdowns** — populates the FROM/TO TAB dropdowns
4. Fill in mapping rows starting at row 8:

    | FROM CELL | TO COL OR CELL |
    | --------- | -------------- |
    | B4        | C4             |
    | B6        | D4             |
    - `FROM CELL` — the cell on the form (e.g. `B4`)
    - `TO COL OR CELL` — where it goes (e.g. `C4` = column C, header at row 4 → data starts at C5)

5. **Transfer → Validate Config** — fix any red cells
6. **Transfer → Transfer Data**

---

## Transfer menu

| Item                           | Does                                                                                                  |
| ------------------------------ | ----------------------------------------------------------------------------------------------------- |
| **Transfer Data**              | Validates and runs all routes. Skips empty and invalid ones, reports results.                         |
| **Validate Config**            | Checks everything without transferring. Marks bad cells red.                                          |
| **Refresh Dropdowns**          | Repopulates the FROM/TO TAB dropdowns from the actual sheets.                                         |
| **Add New Route**              | Appends a new route block to the right. Existing ones preserved.                                      |
| **Delete Route**               | Click a cell inside a route first, then run this. The block is removed and remaining routes renumber. |
| **Reset & Rebuild All Routes** | ⚠️ Erases all Config data and rebuilds N blank blocks.                                                |

---

## Rules

- All `TO` cells in one route must share the same header row (can't mix `C4` and `D50`)
- One `TO` column per route — duplicates are flagged (would overwrite)
- Empty routes are skipped (no blank rows written)
- Source form cells clear automatically after transfer
- Adding a route never erases existing ones — only Reset does
- Pasting new code never touches your Config data

---

## Files
