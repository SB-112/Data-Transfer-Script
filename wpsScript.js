// ============================================================
// DATA TRANSFER ENGINE — WPS JS Macro
// Version: 3.0.0
// Updated: 2026-10-15
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

    const ROUTINGS = [
        { fromCell: "A2", toCell: "A3", isMultiRow: true, expectedHeader: "Date"   },
        { fromCell: "B2", toCell: "B3", isMultiRow: true, expectedHeader: "Name"   },
        { fromCell: "C2", toCell: "C3", isMultiRow: true, expectedHeader: "Amount" }
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
    // 2. RESOLVE ROUTINGS INTO ABSOLUTE POSITIONS
    // ==========================================
    let resolved = [];
    try {
        for (let i = 0; i < ROUTINGS.length; i++) {
            let r = ROUTINGS[i];
            let srcCell = srcSheet.Range(r.fromCell);
            let dstCell = dstSheet.Range(r.toCell);

            resolved.push({
                r: r,
                srcSheet: srcSheet,
                dstSheet: dstSheet,
                srcRow: srcCell.Row,
                srcCol: srcCell.Column,
                dstRow: dstCell.Row,
                dstCol: dstCell.Column
            });
        }
    } catch (e) {
        alert("Error resolving routings:\n" + e.message);
        return;
    }

    // ==========================================
    // 3. VERIFY DESTINATION HEADERS
    // ==========================================
    try {
        let headerIssues = [];
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];
            if (!item.r.expectedHeader) continue;

            let actualHeader = item.dstSheet.Cells(DEST_HEADER_ROW, item.dstCol).Value2;
            let actualStr = (actualHeader === null || actualHeader === undefined)
                ? "" : String(actualHeader).trim();

            if (actualStr !== item.r.expectedHeader) {
                headerIssues.push(
                    "Column " + item.dstCol + ": expected \"" + item.r.expectedHeader +
                    "\" but found \"" + actualStr + "\""
                );
            }
        }

        if (headerIssues.length > 0) {
            alert(
                'Destination tab "' + targetTab + '" has unexpected headers:\n\n' +
                headerIssues.join("\n") +
                "\n\nPlease check the layout or contact the developer."
            );
            return;
        }
    } catch (e) {
        alert("Error verifying headers:\n" + e.message);
        return;
    }

    // ==========================================
    // 4. CHECK IF FORM IS EMPTY
    // ==========================================
    let anyData = false;
    for (let i = 0; i < resolved.length; i++) {
        let v = resolved[i].srcSheet.Cells(resolved[i].srcRow, resolved[i].srcCol).Value2;
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
            if (!item.r.isMultiRow) continue;

            let count = 0;
            let maxScan = 10000;

            while (count < maxScan) {
                let cellVal = item.srcSheet.Cells(item.srcRow + count, item.srcCol).Value2;
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
        let baseRow = resolved[0].dstRow;
        let scanCap = baseRow + 10000;
        let found = -1;

        for (let testRow = baseRow; testRow < scanCap; testRow++) {
            let allEmpty = true;
            for (let i = 0; i < resolved.length; i++) {
                let item = resolved[i];
                let val = item.dstSheet.Cells(testRow, item.dstCol).Value2;
                if (val !== null && val !== undefined && String(val).trim() !== "") {
                    allEmpty = false;
                    break;
                }
            }
            if (allEmpty) {
                found = testRow;
                break;
            }
        }

        targetStartRow = (found === -1) ? baseRow : found;
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

            if (item.r.isMultiRow) {
                for (let offset = 0; offset < maxRowsToTransfer; offset++) {
                    let sRow = item.srcRow + offset;
                    let dRow = targetStartRow + offset;

                    let val = item.srcSheet.Cells(sRow, item.srcCol).Value2;
                    if (val !== null && val !== undefined && String(val).trim() !== "") {
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

    alert("Success! Transferred " + maxRowsToTransfer + " row(s) to \"" + targetTab + "\".");
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

    let currentStr = (currentValue === null || currentValue === undefined)
        ? "" : String(currentValue).trim();

    if (currentStr && tabNames.indexOf(currentStr) !== -1) {
        dropdownCell.Value2 = currentStr;
    } else {
        dropdownCell.Value2 = tabNames[0];
    }
}