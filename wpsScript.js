// ============================================================
// FINANCIAL ASSISTANCE TRANSFER ENGINE — WPS JS MACRO EDITION
// Version: 1.0.1
// Platform: WPS Spreadsheets (ET) / JS Macro (JSA)
// ------------------------------------------------------------
// Paste into a single JS Module in the WPS Macro Editor.
// Save the workbook as .xlsm or .et so macros persist.
// ============================================================

const VERSION = "1.0.1";

// WPS/Office constants (used instead of Excel enums for clarity)
const xlNone = -4142;
const xlCenter = -4108;
const xlValidateList = 3;
const xlValidAlertStop = 1;
const xlBetween = 1;
const vbYesNo = 4;
const vbYes = 6;
const msoControlPopup = 10;

// ============================================================
// SECTION 1: UTILITIES
// ============================================================

/** "A" -> 1, "C" -> 3, "AA" -> 27 */
function colLetterToIndex(letters) {
    var idx = 0;
    for (var i = 0; i < letters.length; i++) {
        idx = idx * 26 + (letters.charCodeAt(i) - 64);
    }
    return idx;
}

/**
 * Parses "C"    -> { colIndex: 3, headerRow: 1 }
 *         "C4"   -> { colIndex: 3, headerRow: 4 }
 *         invalid -> null
 */
function parseTargetReference(refStr) {
    if (refStr === null || refStr === undefined) return null;
    var str = refStr.toString().trim().toUpperCase();
    if (!str) return null;
    var match = str.match(/^([A-Z]+)(\d*)$/);
    if (!match) return null;
    return {
        colIndex: colLetterToIndex(match[1]),
        headerRow: match[2] ? parseInt(match[2], 10) : 1,
    };
}

/** Parses source cell like "B4" -> { row: 4, col: 2 }. null if invalid. */
function parseSourceCell(refStr) {
    if (refStr === null || refStr === undefined) return null;
    var str = refStr.toString().trim().toUpperCase();
    if (!str) return null;
    var match = str.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    return {
        row: parseInt(match[2], 10),
        col: colLetterToIndex(match[1]),
    };
}

/**
 * Opens a spreadsheet by name, path, or full path.
 * Returns null on failure. Reuses already-open workbooks.
 */
function getSpreadsheetFromInput(input) {
    if (input === null || input === undefined) return null;
    var str = input.toString().trim();
    if (!str) return null;

    try {
        for (var i = 1; i <= Application.Workbooks.Count; i++) {
            var wb = Application.Workbooks.Item(i);
            if (wb.Name === str || wb.FullName === str) {
                return wb;
            }
        }
        return Application.Workbooks.Open(str);
    } catch (e) {
        return null;
    }
}

/** Safe worksheet lookup. Returns null if not found. */
function getSheetOrNull(workbook, name) {
    if (!workbook || !name) return null;
    try {
        return workbook.Worksheets.Item(name);
    } catch (e) {
        return null;
    }
}

/** Safe ActiveWorkbook / Config sheet accessor. */
function getConfigSheet() {
    var wb = Application.ActiveWorkbook;
    return getSheetOrNull(wb, "Config");
}

// ============================================================
// SECTION 2: CONFIG PARSER
// ============================================================

/**
 * Reads the entire Config sheet in one call and returns
 * an array of route descriptor objects.
 */
function parseConfig() {
    var configSheet = getConfigSheet();
    if (!configSheet) throw new Error("Config sheet is missing.");

    var usedRange;
    try {
        usedRange = configSheet.UsedRange;
    } catch (e) {
        return [];
    }

    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;
    if (lastRow < 1 || lastCol < 1) return [];

    // Batched read -> 2D array [rows][cols], 0-based
    var grid = configSheet.Range(
        configSheet.Cells(1, 1),
        configSheet.Cells(lastRow, lastCol),
    ).Value2;

    // Locate route banner columns from row 1
    var blockCols = [];
    for (var c = 0; c < grid[0].length; c++) {
        if (String(grid[0][c]).indexOf("--- ROUTE ") !== -1) {
            blockCols.push(c);
        }
    }

    var routes = [];
    for (var i = 0; i < blockCols.length; i++) {
        var col = blockCols[i];
        var valueCol = col + 1;

        var route = {
            index: i,
            startCol: col + 1, // 1-based
            routeName: String(grid[1][valueCol] || "").trim(),
            fromInput: grid[2][valueCol],
            fromTabName: String(grid[3][valueCol] || "").trim(),
            toInput: grid[4][valueCol],
            toTabName: String(grid[5][valueCol] || "").trim(),
            mappings: [],
        };

        // Mapping rows start at row 8 (grid index 7)
        for (var r = 7; r < grid.length; r++) {
            var src = grid[r][col];
            var tgt = grid[r][valueCol];
            if (src === "" || tgt === "") continue;

            var srcParsed = parseSourceCell(src);
            var tgtParsed = parseTargetReference(tgt);

            route.mappings.push({
                rowNum: r + 1,
                sourceCellRef: String(src).trim(),
                sourceCell: srcParsed,
                targetRef: String(tgt).trim(),
                targetCol: tgtParsed ? tgtParsed.colIndex : null,
                headerRow: tgtParsed ? tgtParsed.headerRow : null,
            });
        }

        routes.push(route);
    }

    return routes;
}

/** Returns 1-based start columns of every route banner. */
function getRouteBlockColumns(configSheet) {
    var usedRange = configSheet.UsedRange;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;
    if (lastCol === 0) return [];

    var row1 = configSheet.Range(
        configSheet.Cells(1, 1),
        configSheet.Cells(1, lastCol),
    ).Value2[0];

    var cols = [];
    for (var c = 0; c < row1.length; c++) {
        if (String(row1[c]).indexOf("--- ROUTE ") !== -1) cols.push(c + 1);
    }
    return cols;
}

// ============================================================
// SECTION 3: VALIDATOR
// ============================================================

/**
 * Validates every route. Returns:
 *   { valid: [routeIndex...],
 *     invalid: [{ index, route, errors: [string] }] }
 */
function validateConfig(routes) {
    var configSheet = getConfigSheet();
    var result = { valid: [], invalid: [] };

    for (var i = 0; i < routes.length; i++) {
        var r = routes[i];
        var errors = [];
        var startCol = r.startCol;

        var sourceSS = getSpreadsheetFromInput(r.fromInput);
        var targetSS = getSpreadsheetFromInput(r.toInput);

        // FROM SPREADSHEET
        if (!sourceSS)
            errors.push("FROM SPREADSHEET is invalid or inaccessible.");
        markCell(configSheet, 3, startCol + 1, !sourceSS);

        // TO SPREADSHEET
        if (!targetSS)
            errors.push("TO SPREADSHEET is invalid or inaccessible.");
        markCell(configSheet, 5, startCol + 1, !targetSS);

        // FROM TAB
        var srcSheet = sourceSS
            ? getSheetOrNull(sourceSS, r.fromTabName)
            : null;
        if (sourceSS && !srcSheet) {
            errors.push('FROM TAB "' + r.fromTabName + '" not found.');
        } else if (!sourceSS && r.fromTabName) {
            errors.push('FROM TAB "' + r.fromTabName + '" cannot be checked.');
        }
        markCell(configSheet, 4, startCol + 1, sourceSS && !srcSheet);

        // TO TAB
        var tgtSheet = targetSS ? getSheetOrNull(targetSS, r.toTabName) : null;
        if (targetSS && !tgtSheet) {
            errors.push('TO TAB "' + r.toTabName + '" not found.');
        } else if (!targetSS && r.toTabName) {
            errors.push('TO TAB "' + r.toTabName + '" cannot be checked.');
        }
        markCell(configSheet, 6, startCol + 1, targetSS && !tgtSheet);

        // Mappings
        if (r.mappings.length === 0) {
            errors.push("No mapping rows configured (row 8 and below).");
        }

        var headerRows = {};
        var seenTargetCols = {};

        r.mappings.forEach(function (m) {
            var rowNum = m.rowNum;
            var badSrc = !m.sourceCell;
            var badTgt = m.targetCol === null;

            markCell(configSheet, rowNum, startCol, badSrc);
            markCell(configSheet, rowNum, startCol + 1, badTgt);

            if (badSrc) {
                errors.push(
                    "Row " +
                        rowNum +
                        ': invalid FROM CELL "' +
                        m.sourceCellRef +
                        '".',
                );
            }
            if (badTgt) {
                errors.push(
                    "Row " +
                        rowNum +
                        ': invalid TO reference "' +
                        m.targetRef +
                        '".',
                );
            }

            if (!badTgt) {
                headerRows[m.headerRow] = true;
                if (seenTargetCols[m.targetCol]) {
                    errors.push(
                        "Row " +
                            rowNum +
                            ': TO column "' +
                            m.targetRef +
                            '" is used more than once in this route — data would overwrite.',
                    );
                }
                seenTargetCols[m.targetCol] = true;
            }
        });

        var distinctHeaders = Object.keys(headerRows);
        if (distinctHeaders.length > 1) {
            errors.push(
                "Mappings point to different header rows (" +
                    distinctHeaders.join(", ") +
                    "). All TO cells in one route must share the same header row.",
            );
        }

        if (errors.length === 0) result.valid.push(i);
        else result.invalid.push({ index: i, route: r, errors: errors });
    }

    return result;
}

/** Paints a config cell red (#f4cccc) if bad, default if good. */
function markCell(sheet, row, col, isBad) {
    try {
        var cell = sheet.Cells(row, col);
        if (isBad) {
            // BGR numeric for #f4cccc
            cell.Interior.Color = 13421772;
        } else {
            cell.Interior.ColorIndex = xlNone;
        }
    } catch (e) {
        // ignore — cell may not exist yet
    }
}

// ============================================================
// SECTION 4: TRANSFER ENGINE
// ============================================================

/** Main entry point — Transfer Data menu item. */
function transferData() {
    var routes;

    try {
        routes = parseConfig();
    } catch (e) {
        return MsgBox("Config error: " + e.message);
    }

    if (routes.length === 0) {
        return MsgBox("No route blocks found in Config.");
    }

    var validation = validateConfig(routes);

    if (validation.valid.length === 0) {
        return MsgBox(
            "All " +
                routes.length +
                " route(s) failed validation.\n" +
                "Check the red cells in the Config tab.",
        );
    }

    var successCount = 0;
    var skippedEmpty = 0;
    var runtimeErrors = [];

    Application.ScreenUpdating = false;

    validation.valid.forEach(function (idx) {
        var r = routes[idx];
        try {
            var outcome = transferRoute(r);
            if (outcome === "wrote") successCount++;
            else if (outcome === "empty") skippedEmpty++;
        } catch (err) {
            runtimeErrors.push(
                (r.routeName || "Route " + (idx + 1)) + ": " + err.message,
            );
        }
    });

    Application.ScreenUpdating = true;

    var lines = [];
    lines.push("Version " + VERSION);
    lines.push("");
    lines.push("Transferred: " + successCount + " route(s).");
    if (skippedEmpty > 0)
        lines.push("Skipped (no data): " + skippedEmpty + ".");
    if (validation.invalid.length > 0) {
        lines.push(
            "Skipped (invalid config): " + validation.invalid.length + ".",
        );
    }
    if (runtimeErrors.length > 0) {
        lines.push("");
        lines.push("Runtime errors:");
        lines.push(runtimeErrors.join("\n"));
    }
    MsgBox(lines.join("\n"));
}

/**
 * Transfers one route.
 * Returns "wrote", "empty", or throws on unexpected error.
 */
function transferRoute(r) {
    var sourceSS = getSpreadsheetFromInput(r.fromInput);
    var targetSS = getSpreadsheetFromInput(r.toInput);
    if (!sourceSS || !targetSS)
        throw new Error("Source or target workbook missing.");

    var srcSheet = getSheetOrNull(sourceSS, r.fromTabName);
    var tgtSheet = getSheetOrNull(targetSS, r.toTabName);
    if (!srcSheet || !tgtSheet)
        throw new Error("Source or target tab missing.");

    // --- Read all source cells ---
    var srcValues = r.mappings.map(function (m) {
        return srcSheet.Range(m.sourceCellRef).Value2;
    });

    // --- Skip if every source cell is empty ---
    var hasAnyData = srcValues.some(function (v) {
        return v !== "" && v !== null && v !== undefined;
    });
    if (!hasAnyData) return "empty";

    // --- Scan start row = header row + 1 ---
    var headerRow = r.mappings[0].headerRow;
    var startRow = headerRow + 1;

    // --- Target columns (sorted, unique) ---
    var targetCols = r.mappings
        .map(function (m) {
            return m.targetCol;
        })
        .filter(function (v, i, a) {
            return a.indexOf(v) === i;
        })
        .sort(function (a, b) {
            return a - b;
        });

    // --- Find first empty row ---
    var targetRow = findFirstEmptyRow(tgtSheet, startRow, targetCols);

    // --- Build row array and write in one call ---
    var minCol = targetCols[0];
    var maxCol = targetCols[targetCols.length - 1];
    var width = maxCol - minCol + 1;
    var rowArr = new Array(width).fill("");

    r.mappings.forEach(function (m, i) {
        rowArr[m.targetCol - minCol] = srcValues[i];
    });

    tgtSheet.Range(
        tgtSheet.Cells(targetRow, minCol),
        tgtSheet.Cells(targetRow, maxCol),
    ).Value2 = [rowArr];

    // --- Clear all source cells ---
    r.mappings.forEach(function (m) {
        try {
            srcSheet.Range(m.sourceCellRef).ClearContents();
        } catch (e) {}
    });

    return "wrote";
}

/**
 * Returns first row >= startRow where every column in `cols`
 * is empty. Reads the candidate block in one call.
 * Hard-capped at 10,000 rows.
 */
function findFirstEmptyRow(sheet, startRow, cols) {
    var usedRange = sheet.UsedRange;
    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var scanEnd = Math.max(lastRow + 1, startRow);
    if (scanEnd - startRow > 10000) scanEnd = startRow + 10000;

    var height = Math.max(1, scanEnd - startRow + 1);
    var minCol = cols[0];
    var maxCol = cols[cols.length - 1];

    var block = sheet.Range(
        sheet.Cells(startRow, minCol),
        sheet.Cells(startRow + height - 1, maxCol),
    ).Value2;

    // Value2 may return a scalar if the range is 1x1
    if (!block || !block.length) return startRow;

    for (var r = 0; r < block.length; r++) {
        var rowVals = block[r];
        var empty = true;
        for (var c = 0; c < cols.length; c++) {
            var val = rowVals[cols[c] - minCol];
            if (val !== "" && val !== null && val !== undefined) {
                empty = false;
                break;
            }
        }
        if (empty) return startRow + r;
    }
    return startRow + block.length;
}

// ============================================================
// SECTION 5: UI / MENU / CONFIG DRAWING
// ============================================================

/** Fires when a workbook opens. Registers the Transfer menu. */
function Application_WorkbookOpen(wk) {
    var configSheet = getSheetOrNull(wk, "Config");
    if (!configSheet) {
        configSheet = wk.Worksheets.Add();
        configSheet.Name = "Config";
        initConfig(configSheet);
    }

    try {
        CommandBars("Transfer").Delete();
    } catch (e) {}

    var menuBar = CommandBars.Add("Transfer");
    menuBar.Visible = true;

    var popup = menuBar.Controls.Add(msoControlPopup);
    popup.Caption = "Transfer";

    var item1 = popup.Controls.Add();
    item1.Caption = "Transfer Data";
    item1.OnAction = "transferData";

    var item2 = popup.Controls.Add();
    item2.Caption = "Validate Config";
    item2.OnAction = "validateConfigMenu";

    popup.Controls.Add(); // separator

    var item3 = popup.Controls.Add();
    item3.Caption = "Refresh Dropdowns";
    item3.OnAction = "updateSheetDropdowns";

    var item4 = popup.Controls.Add();
    item4.Caption = "Add New Route";
    item4.OnAction = "addNewRoute";

    var item5 = popup.Controls.Add();
    item5.Caption = "Delete Route (click inside one first)";
    item5.OnAction = "deleteRoute";

    var item6 = popup.Controls.Add();
    item6.Caption = "Reset & Rebuild All Routes";
    item6.OnAction = "resetAllRoutes";
}

function validateConfigMenu() {
    var routes;
    try {
        routes = parseConfig();
    } catch (e) {
        return MsgBox("Config error: " + e.message);
    }
    if (routes.length === 0) return MsgBox("No routes found.");

    var result = validateConfig(routes);
    var lines = ["Version " + VERSION, ""];
    lines.push("Valid: " + result.valid.length + " / " + routes.length);

    result.invalid.forEach(function (inv) {
        lines.push("");
        lines.push("X " + (inv.route.routeName || "Route " + (inv.index + 1)));
        inv.errors.forEach(function (e) {
            lines.push("   - " + e);
        });
    });

    if (result.invalid.length === 0) lines.push("\nAll good.");
    MsgBox(lines.join("\n"));
}

/** Draws one route block. Shared by init / add / reset. */
function drawRouteBlock(configSheet, routeNumber, startCol, currentId) {
    var titleRange = configSheet.Range(
        configSheet.Cells(1, startCol),
        configSheet.Cells(1, startCol + 1),
    );
    titleRange.Merge();
    titleRange.Value2 = "--- ROUTE " + routeNumber + " CONFIGURATION ---";
    titleRange.Font.Bold = true;
    titleRange.Interior.Color = 14600616; // #4a86e8 approx in BGR
    titleRange.Font.Color = 16777215; // white
    titleRange.HorizontalAlignment = xlCenter;

    var rows = [
        ["ROUTE NAME:", "Route " + routeNumber],
        ["FROM SPREADSHEET (URL or ID):", currentId],
        ["FROM TAB:", ""],
        ["TO SPREADSHEET (URL or ID):", ""],
        ["TO TAB:", ""],
    ];
    rows.forEach(function (pair, i) {
        var labelCell = configSheet.Cells(2 + i, startCol);
        labelCell.Value2 = pair[0];
        labelCell.Font.Bold = true;
        if (pair[1] !== "") {
            configSheet.Cells(2 + i, startCol + 1).Value2 = pair[1];
        }
    });

    // Mapping header row
    var h1 = configSheet.Cells(7, startCol);
    h1.Value2 = "FROM CELL:";
    h1.Font.Bold = true;
    h1.Interior.Color = 15396585; // #e8eaed approx

    var h2 = configSheet.Cells(7, startCol + 1);
    h2.Value2 = "TO COL OR CELL (e.g. C or C4):";
    h2.Font.Bold = true;
    h2.Interior.Color = 15396585;

    configSheet.Columns(startCol).ColumnWidth = 240;
    configSheet.Columns(startCol + 1).ColumnWidth = 280;
    configSheet.Columns(startCol + 2).ColumnWidth = 40;
}

function initConfig(configSheet) {
    configSheet.Cells.Clear();
    var currentId = Application.ActiveWorkbook.Name;
    drawRouteBlock(configSheet, 1, 1, currentId);
}

function addNewRoute() {
    var wb = Application.ActiveWorkbook;
    var configSheet = getSheetOrNull(wb, "Config");
    if (!configSheet) {
        configSheet = wb.Worksheets.Add();
        configSheet.Name = "Config";
    }

    var blockCols = getRouteBlockColumns(configSheet);
    var nextRouteNumber = blockCols.length + 1;
    var startCol =
        blockCols.length > 0 ? blockCols[blockCols.length - 1] + 3 : 1;

    drawRouteBlock(configSheet, nextRouteNumber, startCol, wb.Name);
    MsgBox(
        "Added Route " + nextRouteNumber + " block. Existing data preserved.",
    );
}

function deleteRoute() {
    var wb = Application.ActiveWorkbook;
    var configSheet = getSheetOrNull(wb, "Config");
    if (!configSheet) return MsgBox("Config sheet missing.");

    if (wb.ActiveSheet.Name !== "Config") {
        return MsgBox(
            "Open the Config tab and click a cell inside the route " +
                "you want to delete, then try again.",
        );
    }

    var activeCol = Application.ActiveCell.Column;

    var blockCols = getRouteBlockColumns(configSheet);
    var targetStartCol = null;
    for (var i = 0; i < blockCols.length; i++) {
        var start = blockCols[i];
        if (activeCol >= start && activeCol <= start + 2) {
            targetStartCol = start;
            break;
        }
    }

    if (targetStartCol === null) {
        return MsgBox(
            "That cell isn't inside a route block. Click a cell inside " +
                "the route you want to delete.",
        );
    }

    var routeName =
        String(configSheet.Cells(2, targetStartCol + 1).Value2 || "").trim() ||
        "Route at column " + targetStartCol;

    var confirm = MsgBox(
        'Delete "' + routeName + '"?\n\nThis cannot be undone.',
        vbYesNo,
    );
    if (confirm !== vbYes) return;

    // Physically delete the 3 columns of this block.
    configSheet
        .Range(
            configSheet.Columns(targetStartCol),
            configSheet.Columns(targetStartCol + 2),
        )
        .Delete();

    renumberRouteBlocks(configSheet);

    var remaining = getRouteBlockColumns(configSheet).length;
    MsgBox(
        'Route "' +
            routeName +
            '" deleted.\n' +
            remaining +
            " route(s) remaining.",
    );
}

/** Rewrites every route banner in row 1 sequentially. */
function renumberRouteBlocks(configSheet) {
    var blockCols = getRouteBlockColumns(configSheet);
    for (var i = 0; i < blockCols.length; i++) {
        var startCol = blockCols[i];
        configSheet.Cells(1, startCol).Value2 =
            "--- ROUTE " + (i + 1) + " CONFIGURATION ---";
    }
}

function resetAllRoutes() {
    var confirm = MsgBox("This will ERASE all Config data. Continue?", vbYesNo);
    if (confirm !== vbYes) return;

    var countStr = InputBox("How many blank routes?");
    if (countStr === null || countStr === undefined || countStr === "") return;

    var count = parseInt(String(countStr).trim(), 10);
    if (isNaN(count) || count < 1) return MsgBox("Invalid number.");

    var wb = Application.ActiveWorkbook;
    var configSheet = getSheetOrNull(wb, "Config");
    if (!configSheet) {
        configSheet = wb.Worksheets.Add();
        configSheet.Name = "Config";
    }
    configSheet.Cells.Clear();

    var currentId = wb.Name;
    for (var i = 1; i <= count; i++) {
        drawRouteBlock(configSheet, i, 1 + (i - 1) * 3, currentId);
    }
    MsgBox("Created " + count + " route(s).");
}

function updateSheetDropdowns() {
    var wb = Application.ActiveWorkbook;
    var configSheet = getSheetOrNull(wb, "Config");
    if (!configSheet) return MsgBox("Config sheet missing.");

    var blockCols = getRouteBlockColumns(configSheet);
    if (blockCols.length === 0) return MsgBox("No route blocks found.");

    for (var k = 0; k < blockCols.length; k++) {
        var startCol = blockCols[k];
        var fromInput = configSheet.Cells(3, startCol + 1).Value2;
        var toInput = configSheet.Cells(5, startCol + 1).Value2;

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        if (sourceSS) applyTabDropdown(configSheet, sourceSS, 4, startCol + 1);
        if (targetSS) applyTabDropdown(configSheet, targetSS, 6, startCol + 1);
    }

    MsgBox("Updated dropdowns for " + blockCols.length + " route(s).");
}

function applyTabDropdown(configSheet, workbook, row, col) {
    var names = [];
    for (var i = 1; i <= workbook.Worksheets.Count; i++) {
        names.push(workbook.Worksheets.Item(i).Name);
    }
    if (names.length === 0) return;

    var cell = configSheet.Cells(row, col);

    try {
        cell.Validation.Delete();
    } catch (e) {}

    cell.Validation.Add(
        xlValidateList,
        xlValidAlertStop,
        xlBetween,
        names.join(","),
    );
    cell.Validation.InCellDropdown = true;
    cell.Validation.IgnoreBlank = true;

    var current = cell.Value2;
    if (current === "" || current === null || current === undefined) {
        cell.Value2 = names[0];
    }
}
