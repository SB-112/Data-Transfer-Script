# Data Transfer Engine

Automates copying form data from a source spreadsheet into a
destination summary spreadsheet — one click instead of manual retyping.

Two versions are included:

- **Google Apps Script** (`apps-script/`) — for Google Sheets
- **WPS JS Macro** (`wps/`) — for WPS Office spreadsheets

Pick whichever matches your office's platform.

---

## ⚠️ Notice

This project is **vibe coded** — the code was written with heavy AI
assistance (DeepSeek). The syntax, APIs, and implementation were
largely generated with AI help and refined through testing.

However, the **architecture, design, and problem-solving** are original:

- The config-driven routing concept
- The multi-row group handling logic
- The append-to-first-empty-row strategy
- The overall workflow and structure

AI was a tool; the design decisions were mine.

---

## Which version do I use?

| Platform                  | Use                   |
| ------------------------- | --------------------- |
| Google Sheets             | `apps-script/Code.gs` |
| WPS Office (spreadsheets) | `wps/Code.js`         |

Both solve the same problem. Both work independently — you only
install one.

---

# Google Apps Script version

## Install (for the worker)

1. Open the source spreadsheet
2. Extensions → Apps Script
3. Click `Code.gs`, press **Ctrl+A**, then **Delete**
4. Paste the full contents of `apps-script/Code.gs`
5. Save (💾) and reload the spreadsheet tab
6. The **Transfer** menu appears in the top bar

> To update later: repeat steps 3–6 with the newest version.

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

## Transfer menu

| Item                           | Does                                                                                                  |
| ------------------------------ | ----------------------------------------------------------------------------------------------------- |
| **Transfer Data**              | Validates and runs all routes. Skips empty and invalid ones, reports results.                         |
| **Validate Config**            | Checks everything without transferring. Marks bad cells red.                                          |
| **Find Applicant**             | Prompts for a name and tab, finds the latest entry and its date.                                      |
| **Refresh Dropdowns**          | Repopulates the FROM/TO TAB dropdowns from the actual sheets.                                         |
| **Add New Route**              | Appends a new route block to the right. Existing ones preserved.                                      |
| **Delete Route**               | Click a cell inside a route first, then run this. The block is removed and remaining routes renumber. |
| **Reset & Rebuild All Routes** | ⚠️ Erases all Config data and rebuilds N blank blocks.                                                |

## Rules

- All `TO` cells in one route must share the same header row
- One `TO` column per route — duplicates are flagged
- Empty routes are skipped (no blank rows written)
- Source form cells clear automatically after transfer
- Adding a route never erases existing ones — only Reset does
- Pasting new code never touches your Config data

---

# WPS JS Macro version

## Install (for the worker)

1. Open the source workbook in WPS
2. Go to **Tools → Macros → Edit Macros** (or **Developer → Macros**)
3. Paste the contents of `wps/Code.js`
4. Save the workbook as `.xlsm` (macro-enabled)
5. Assign the macro `transferFormToSummary` to a button (or run it via Developer → Macros)

> To update later: replace the macro code with the newest version.

## Requirements

- WPS Office (with JS Macro support)
- Both workbooks must be open in WPS before running
    - Source workbook (the form)
    - Destination workbook (the summary)
- Routings are defined in the `routings` array at the top of the script

## How it works

The script:

1. Reads values from source cells defined in `routings`
2. Finds the first row in the destination where all mapped columns are empty
3. Writes the group/entry starting there
4. Clears the source form

## Configuration

Routes are hardcoded in the `routings` array. To add a new column:

```javascript
{
  fromWorkbook: "Workbook1.xlsm", fromTab: "Sheet1", fromCell: "A2",
  toWorkbook: "Workbook2.xlsx", toTab: "Sheet1", toCell: "A3",
  isMultiRow: true
}
```
