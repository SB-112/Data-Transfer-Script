// ============================================================
// DATA TRANSFER ENGINE — WPS JS Macro
// Version: 3.2.0
// Updated: 2026-10-16
// ------------------------------------------------------------
// Automates copying form data from a source workbook into a
// destination summary workbook. The destination tab is chosen
// via a dropdown on each source tab. Destination columns are
// located by header name, so column order doesn't matter.
// ============================================================

function transferFormToSummary() {
    let app = Application;

    // ============================================================
    // ⚙️ CONFIGURATION — edit this block per deployment
    // ============================================================
    const SOURCE_WORKBOOK = "Workbook1.xlsm";
    const TARGET_WORKBOOK = "Workbook2.xlsx";
    const DROPDOWN_CELL = "E6";
    const DEST_HEADER_ROW = 3;

    // One entry per column to transfer.
    //   fromCell      = top-left cell of the source column on the form
    //   targetHeader  = exact header text in the destination to write into
    //   isMultiRow    = true if the column can have multiple rows per entry
    //   constantValue = (optional) a fixed value to write instead of
    //                   reading from a source cell. Use for columns
    //                   like "Type" that depend on which form this is.
    //
    const ROUTINGS = [
        { fromCell: "A2", targetHeader: "Date", isMultiRow: true },
        { fromCell: "B2", targetHeader: "Name", isMultiRow: true },
        { fromCell: "C2", targetHeader: "Amount", isMultiRow: true },
        {
            targetHeader: "Type",
            constantValue: "Educational",
            isMultiRow: true,
        },
    ];
    // ============================================================
    // 🚫 ENGINE — do not edit below this line
    // ============================================================

    let srcSheet, srcWb, dstWb, dstSheet;
    let targetTab;

    // ==========================================
    // 1. RESOLVE WORKBOOKS, SOURCE TAB, TARGET TAB
    // ==========================================
    try {
        srcWb = app.Workbooks(SOURCE_WORKBOOK);
        dstWb = app.Workbooks(TARGET_WORKBOOK);

        srcSheet = app.ActiveSheet;

        refreshAreaDropdown(srcWb, dstWb, srcSheet, DROPDOWN_CELL);

        targetTab = srcSheet.Range(DROPDOWN_CELL).Value2;
        if (!targetTab || String(targetTab).trim() === "") {
            alert("Please select an area from the dropdown first.");
            return;
        }
        targetTab = String(targetTab).trim();

        dstSheet = dstWb.Sheets(targetTab);
    } catch (e) {
        alert("Error resolving workbooks/sheets:\n" + e.message);
        return;
    }

    // ==========================================
    // 2. BUILD HEADER MAP FROM DESTINATION
    // ==========================================
    let headerMap = {};
    try {
        let lastCol = dstSheet.UsedRange.Columns.Count;
        for (let col = 1; col <= lastCol; col++) {
            let headerVal = dstSheet.Cells(DEST_HEADER_ROW, col).Value2;
            if (headerVal === null || headerVal === undefined) continue;
            let headerStr = String(headerVal).trim();
            if (headerStr === "") continue;
            headerMap[headerStr] = col;
        }
    } catch (e) {
        alert("Error reading destination headers:\n" + e.message);
        return;
    }

    // ==========================================
    // 3. RESOLVE ROUTINGS USING HEADER NAMES
    // ==========================================
    let resolved = [];
    let missingHeaders = [];
    try {
        for (let i = 0; i < ROUTINGS.length; i++) {
            let r = ROUTINGS[i];
            let dstCol = headerMap[r.targetHeader];

            if (!dstCol) {
                missingHeaders.push(r.targetHeader);
                continue;
            }

            // Constant value routing (no source cell)
            if (r.constantValue !== undefined) {
                resolved.push({
                    r: r,
                    srcSheet: srcSheet,
                    dstSheet: dstSheet,
                    isConstant: true,
                    value: r.constantValue,
                    dstRow: DEST_HEADER_ROW + 1,
                    dstCol: dstCol,
                });
                continue;
            }

            // Normal routing (reads from source cell)
            let srcCell = srcSheet.Range(r.fromCell);

            resolved.push({
                r: r,
                srcSheet: srcSheet,
                dstSheet: dstSheet,
                isConstant: false,
                srcRow: srcCell.Row,
                srcCol: srcCell.Column,
                dstRow: DEST_HEADER_ROW + 1,
                dstCol: dstCol,
            });
        }

        if (missingHeaders.length > 0) {
            alert(
                'Destination tab "' +
                    targetTab +
                    '" is missing these headers:\n\n' +
                    missingHeaders.join("\n") +
                    "\n\nPlease check the tab layout.",
            );
            return;
        }
    } catch (e) {
        alert("Error resolving routings:\n" + e.message);
        return;
    }

    // ==========================================
    // 4. CHECK IF FORM IS EMPTY
    // ==========================================
    let anyData = false;
    for (let i = 0; i < resolved.length; i++) {
        let item = resolved[i];
        if (item.isConstant) continue;
        let v = item.srcSheet.Cells(item.srcRow, item.srcCol).Value2;
        if (v !== null && v !== undefined && String(v).trim() !== "") {
            anyData = true;
            break;
        }
    }
    if (!anyData) {
        alert("Nothing to transfer — the form is empty.");
        return;
    }

    // ==========================================
    // 5. CALCULATE MULTI-ROW COUNT
    // ==========================================
    let maxRowsToTransfer = 1;
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];
            if (item.isConstant) continue;
            if (!item.r.isMultiRow) continue;

            let count = 0;
            let maxScan = 10000;

            while (count < maxScan) {
                let cellVal = item.srcSheet.Cells(
                    item.srcRow + count,
                    item.srcCol,
                ).Value2;
                if (cellVal === null || cellVal === undefined) break;
                if (String(cellVal).trim() === "") break;
                count++;
            }

            if (count > maxRowsToTransfer) {
                maxRowsToTransfer = count;
            }
        }
    } catch (e) {
        alert("Error in multi-row check:\n" + e.message);
        return;
    }

    // ==========================================
    // 6. FIND FIRST ALL-BLANK ROW IN TARGET
    // ==========================================
    let targetStartRow;
    try {
        let baseRow = DEST_HEADER_ROW + 1;
        let scanCap = baseRow + 10000;
        let found = -1;

        for (let testRow = baseRow; testRow < scanCap; testRow++) {
            let allEmpty = true;
            for (let i = 0; i < resolved.length; i++) {
                let item = resolved[i];
                let val = item.dstSheet.Cells(testRow, item.dstCol).Value2;
                if (
                    val !== null &&
                    val !== undefined &&
                    String(val).trim() !== ""
                ) {
                    allEmpty = false;
                    break;
                }
            }
            if (allEmpty) {
                found = testRow;
                break;
            }
        }

        targetStartRow = found === -1 ? baseRow : found;
    } catch (e) {
        alert("Error finding destination row:\n" + e.message);
        return;
    }

    // ==========================================
    // 7. TRANSFER DATA
    // ==========================================
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];

            // Constant value → write to every row
            if (item.isConstant) {
                for (let offset = 0; offset < maxRowsToTransfer; offset++) {
                    item.dstSheet.Cells(
                        targetStartRow + offset,
                        item.dstCol,
                    ).Value2 = item.value;
                }
                continue;
            }

            if (item.r.isMultiRow) {
                for (let offset = 0; offset < maxRowsToTransfer; offset++) {
                    let sRow = item.srcRow + offset;
                    let dRow = targetStartRow + offset;

                    let val = item.srcSheet.Cells(sRow, item.srcCol).Value2;
                    if (
                        val !== null &&
                        val !== undefined &&
                        String(val).trim() !== ""
                    ) {
                        item.dstSheet.Cells(dRow, item.dstCol).Value2 = val;
                    }
                }
            } else {
                let val = item.srcSheet.Cells(item.srcRow, item.srcCol).Value2;
                item.dstSheet.Cells(targetStartRow, item.dstCol).Value2 = val;
            }
        }
    } catch (e) {
        alert("Transfer error:\n" + e.message);
        return;
    }

    // ==========================================
    // 8. CLEAR SOURCE FORM
    // ==========================================
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];
            if (item.isConstant) continue;

            if (item.r.isMultiRow) {
                for (let offset = 0; offset < maxRowsToTransfer; offset++) {
                    let sRow = item.srcRow + offset;
                    item.srcSheet.Cells(sRow, item.srcCol).Value2 = "";
                }
            } else {
                item.srcSheet.Cells(item.srcRow, item.srcCol).Value2 = "";
            }
        }
    } catch (e) {
        alert("Clear error:\n" + e.message);
        return;
    }

    alert(
        "Success! Transferred " +
            maxRowsToTransfer +
            ' row(s) to "' +
            targetTab +
            '".',
    );
}

// ============================================================
// DROPDOWN HELPER
// ============================================================

function refreshAreaDropdown(srcWb, dstWb, srcSheet, dropdownCellAddress) {
    let tabNames = [];
    for (let i = 1; i <= dstWb.Sheets.Count; i++) {
        tabNames.push(dstWb.Sheets(i).Name);
    }
    if (tabNames.length === 0) return;

    let dropdownCell = srcSheet.Range(dropdownCellAddress);
    let currentValue = dropdownCell.Value2;

    let listString = tabNames.join(",");

    try {
        dropdownCell.Validation.Delete();
    } catch (e) {
        // no existing validation — fine
    }

    try {
        dropdownCell.Validation.Add(3, 1, 1, listString);
        dropdownCell.Validation.IgnoreBlank = true;
        dropdownCell.Validation.InCellDropdown = true;
    } catch (e) {
        alert("Dropdown error:\n" + e.message);
    }

    let currentStr =
        currentValue === null || currentValue === undefined
            ? ""
            : String(currentValue).trim();

    if (currentStr && tabNames.indexOf(currentStr) !== -1) {
        dropdownCell.Value2 = currentStr;
    } else {
        dropdownCell.Value2 = tabNames[0];
    }
}
