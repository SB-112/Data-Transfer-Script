// ============================================================
// WPS Form Transfer Script
// Version: 2.1.0
// Updated: 2026-10-15
// ------------------------------------------------------------
// Copies form data from a source workbook into a destination
// summary table. Handles multi-row families where only the
// first row carries date/amount but every row carries a name.
// ============================================================

function transferFormToSummary() {
    let app = Application;

    // ==========================================
    // CONFIGURABLE ROUTINGS
    // ==========================================
    let routings = [
        // Add Routes Here
    ];

    // ==========================================
    // 0. RESOLVE ALL WORKBOOKS / SHEETS UPFRONT
    // ==========================================
    let resolved = [];
    try {
        for (let i = 0; i < routings.length; i++) {
            let r = routings[i];
            let srcWb = app.Workbooks(r.fromWorkbook);
            let srcSheet = srcWb.Sheets(r.fromTab);
            let dstWb = app.Workbooks(r.toWorkbook);
            let dstSheet = dstWb.Sheets(r.toTab);

            let srcCell = srcSheet.Range(r.fromCell);
            let dstCell = dstSheet.Range(r.toCell);

            resolved.push({
                r: r,
                srcSheet: srcSheet,
                dstSheet: dstSheet,
                srcRow: srcCell.Row,
                srcCol: srcCell.Column,
                dstRow: dstCell.Row,
                dstCol: dstCell.Column,
            });
        }
    } catch (e) {
        alert("Error opening workbooks or sheets:\n" + e.message);
        return;
    }

    // ==========================================
    // 1. CHECK IF FORM IS EMPTY
    // ==========================================
    let anyData = false;
    for (let i = 0; i < resolved.length; i++) {
        let v = resolved[i].srcSheet.Cells(
            resolved[i].srcRow,
            resolved[i].srcCol,
        ).Value2;
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
    // 2. CALCULATE MULTI-ROW COUNT
    // Uses the widest source column across all multi-row
    // routings. The Name column is typically the widest,
    // which matches how families fill the form.
    // ==========================================
    let maxRowsToTransfer = 1;
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];
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
        alert("Error in Multi-Row check:\n" + e.message);
        return;
    }

    // ==========================================
    // 3. FIND DESTINATION START ROW
    // Scan downward from the base row and find the first
    // row where ALL mapped destination columns are empty.
    // That's where the new family gets appended.
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
    // 4. TRANSFER DATA
    // ==========================================
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];

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
        alert("Transfer Error:\n" + e.message);
        return;
    }

    // ==========================================
    // 5. CLEAR SOURCE FORM
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
        alert("Clear Error:\n" + e.message);
        return;
    }

    alert(
        "Success! Transferred " +
            maxRowsToTransfer +
            " row(s) and cleared the form.",
    );
}
