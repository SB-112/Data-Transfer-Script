function transferFormToSummary() {
    let app = Application;
    
    // ==========================================
    // CONFIGURABLE ROUTINGS
    // Each routing: { fromWorkbook, fromTab, fromCell, toWorkbook, toTab, toCell, isMultiRow }
    // ==========================================
    let routings = [
        { 
            fromWorkbook: "Workbook1.xlsm", fromTab: "Sheet1", fromCell: "A2", 
            toWorkbook: "Workbook2.xlsx", toTab: "Sheet1", toCell: "A3", 
            isMultiRow: true 
        },
        { 
            fromWorkbook: "Workbook1.xlsm", fromTab: "Sheet1", fromCell: "B2", 
            toWorkbook: "Workbook2.xlsx", toTab: "Sheet1", toCell: "B3", 
            isMultiRow: true 
        },
        { 
            fromWorkbook: "Workbook1.xlsm", fromTab: "Sheet1", fromCell: "C2", 
            toWorkbook: "Workbook2.xlsx", toTab: "Sheet1", toCell: "C3", 
            isMultiRow: true 
        }
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
                dstCol: dstCell.Column
            });
        }
    } catch (e) {
        alert("Error opening workbooks or sheets:\n" + e.message);
        return;
    }
    
    // ==========================================
    // 1. CHECK IF FORM IS EMPTY (early exit)
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
    // 2. CALCULATE MULTI-ROW COUNT
    // Uses the LONGEST non-empty run from any multi-row routing.
    // ==========================================
    let maxRowsToTransfer = 1;
    try {
        for (let i = 0; i < resolved.length; i++) {
            let item = resolved[i];
            if (!item.r.isMultiRow) continue;
            
            let count = 0;
            let maxScan = 10000; // safety cap
            
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
        alert("Error in Multi-Row check:\n" + e.message);
        return;
    }
    
    // ==========================================
    // 3. FIND DESTINATION START ROW
    // Uses the FIRST routing's destination column to find the next empty row.
    // Falls back to the base row if column is empty.
    // ==========================================
    let targetStartRow;
    try {
        let first = resolved[0];
        let lastRowInCol = first.dstSheet.Cells(1048576, first.dstCol).End(-4162).Row; // xlUp = -4162
        
        // If the column is empty (lastRow returns 1 and baseRow is 3), use baseRow
        if (lastRowInCol < first.dstRow) {
            targetStartRow = first.dstRow;
        } else {
            targetStartRow = lastRowInCol + 1;
        }
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
    
    alert("Success! Transferred " + maxRowsToTransfer + " row(s) and cleared the form.");
}