# Data Transfer Engine

Automates copying form data from a source spreadsheet into a
destination summary spreadsheet — one click instead of manual retyping.

Two versions are included:

- **Google Apps Script** (`googleScript.js`) — for Google Sheets
- **WPS JS Macro** (`wpsScript.js`) — for WPS Office spreadsheets

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
- The dropdown-driven destination routing
- The overall workflow and structure

AI was a tool; the design decisions were mine.

---

## Which version do I use?

| Platform                  | Use               |
| ------------------------- | ----------------- |
| Google Sheets             | `googleScript.js` |
| WPS Office (spreadsheets) | `wpsScript.js`    |

Both solve the same problem. Both work independently — you only
install one.

---

# Google Apps Script version

## Install (for the worker)

1. Open the source spreadsheet
2. Extensions → Apps Script
3. Click `Code.gs`, press **Ctrl+A**, then **Delete**
4. Paste the full contents of `googleScript.js`
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
3. Paste the contents of `wpsScript.js`
4. Save the workbook as `.xlsm` (macro-enabled)
5. Assign the macro `transferFormToSummary` to a button on each form tab
6. Click the button to transfer

> To update later: replace the macro code with the newest version.

## Requirements

- WPS Office (with JS Macro support)
- Both workbooks must be open in WPS before running:
    - Source workbook (the form)
    - Destination workbook (the summary)
- Destination tabs must have consistent layout (same header order on every tab)
- Header row must be on the same row on every destination tab (default: row 3)

## How it works

The script:

1. Reads the **destination tab name** from a dropdown on the source form
2. Verifies the destination tab's headers match the expected layout
3. Reads values from source cells defined in `ROUTINGS`
4. Finds the first row in the destination where all mapped columns are empty
5. Writes the group/entry starting there
6. Clears the source form

## Configuration

Everything client-specific lives in the **CONFIGURATION** block at the
top of `transferFormToSummary()`. To adapt for a new client, edit only
that block — nothing below it should need to change.

```javascript
const SOURCE_WORKBOOK = "FA Forms.xlsm";
const TARGET_WORKBOOK = "Summary.xlsx";
const DROPDOWN_CELL = "H10";
const DEST_HEADER_ROW = 3;

const ROUTINGS = [
    { fromCell: "A2", toCell: "A3", isMultiRow: true, expectedHeader: "Date" },
    { fromCell: "B2", toCell: "B3", isMultiRow: true, expectedHeader: "Name" },
    {
        fromCell: "C2",
        toCell: "C3",
        isMultiRow: true,
        expectedHeader: "Amount",
    },
];
```
