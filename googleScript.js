// ============================================================
// FINANCIAL ASSISTANCE TRANSFER ENGINE
// Version: 1.0.1
// Updated: 2026-XX-XX
// ------------------------------------------------------------
// Full setup instructions: see README.md in the repo.
// ============================================================

const VERSION = "1.0.1";

// ============================================================
// SECTION 1: UTILITIES
// ============================================================

/** "A" → 1, "C" → 3, "AA" → 27 */
function colLetterToIndex(letters) {
    var idx = 0;
    for (var i = 0; i < letters.length; i++) {
        idx = idx * 26 + (letters.charCodeAt(i) - 64);
    }
    return idx;
}

/**
 * Parses "C"    → { colIndex: 3, headerRow: 1 }
 *         "C4"   → { colIndex: 3, headerRow: 4 }
 *         invalid → null
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

/** Parses source cell like "B4" → { row: 4, col: 2 }. null if invalid. */
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

/** Opens a spreadsheet by URL or ID. Returns null on failure. */
function getSpreadsheetFromInput(input) {
    if (input === null || input === undefined) return null;
    var str = input.toString().trim();
    if (!str) return null;
    try {
        return str.indexOf("http") === 0
            ? SpreadsheetApp.openByUrl(str)
            : SpreadsheetApp.openById(str);
    } catch (e) {
        return null;
    }
}

// ============================================================
// SECTION 2: CONFIG PARSER
// ============================================================

/**
 * Reads the entire Config sheet in one API call and returns
 * an array of route descriptor objects.
 */
function parseConfig() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) throw new Error("Config sheet is missing.");

    var lastRow = configSheet.getLastRow();
    var lastCol = configSheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return [];

    var grid = configSheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Locate route banner columns from row 1 (0-based in grid).
    var blockCols = [];
    for (var c = 0; c < grid[0].length; c++) {
        if (grid[0][c].toString().indexOf("--- ROUTE ") !== -1) {
            blockCols.push(c);
        }
    }

    var routes = [];
    for (var i = 0; i < blockCols.length; i++) {
        var col = blockCols[i];
        var valueCol = col + 1;

        var route = {
            index: i,
            startCol: col + 1, // 1-based for range ops
            routeName: String(grid[1][valueCol] || "").trim(),
            fromInput: grid[2][valueCol],
            fromTabName: String(grid[3][valueCol] || "").trim(),
            toInput: grid[4][valueCol],
            toTabName: String(grid[5][valueCol] || "").trim(),
            mappings: [],
        };

        // Mapping rows start at row 8 (grid index 7).
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
    var lastCol = configSheet.getLastColumn();
    if (lastCol === 0) return [];
    var row1 = configSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var cols = [];
    for (var c = 0; c < row1.length; c++) {
        if (row1[c].toString().indexOf("--- ROUTE ") !== -1) cols.push(c + 1);
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
 *
 * Also paints bad config cells red.
 */
function validateConfig(routes) {
    var configSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");

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
        var srcSheet = null;
        if (sourceSS) {
            srcSheet = sourceSS.getSheetByName(r.fromTabName);
            if (!srcSheet) {
                errors.push('FROM TAB "' + r.fromTabName + '" not found.');
            }
        } else if (r.fromTabName) {
            errors.push('FROM TAB "' + r.fromTabName + '" cannot be checked.');
        }
        markCell(configSheet, 4, startCol + 1, sourceSS && !srcSheet);

        // TO TAB
        var tgtSheet = null;
        if (targetSS) {
            tgtSheet = targetSS.getSheetByName(r.toTabName);
            if (!tgtSheet) {
                errors.push('TO TAB "' + r.toTabName + '" not found.');
            }
        } else if (r.toTabName) {
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

                // Duplicate target column check (would overwrite).
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

        // All mappings must share the same header row.
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

/** Paints a config cell red if bad, default if good. */
function markCell(sheet, row, col, isBad) {
    try {
        sheet.getRange(row, col).setBackground(isBad ? "#f4cccc" : null);
    } catch (e) {
        // ignore — cell may not exist yet
    }
}

// ============================================================
// SECTION 4: TRANSFER ENGINE
// ============================================================

/** Main entry point — Transfer Data menu item. */
function transferData() {
    var ui = SpreadsheetApp.getUi();
    var routes;

    try {
        routes = parseConfig();
    } catch (e) {
        return ui.alert("Config error: " + e.message);
    }

    if (routes.length === 0) {
        return ui.alert("No route blocks found in Config.");
    }

    var validation = validateConfig(routes);

    if (validation.valid.length === 0) {
        return ui.alert(
            "All " +
                routes.length +
                " route(s) failed validation.\n" +
                "Check the red cells in the Config tab.",
        );
    }

    var successCount = 0;
    var skippedEmpty = 0;
    var runtimeErrors = [];

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

    var lines = [];
    lines.push("Version " + VERSION);
    lines.push("");
    lines.push("✅ Transferred: " + successCount + " route(s).");
    if (skippedEmpty > 0)
        lines.push("⏭️ Skipped (no data): " + skippedEmpty + ".");
    if (validation.invalid.length > 0) {
        lines.push(
            "⚠️ Skipped (invalid config): " + validation.invalid.length + ".",
        );
    }
    if (runtimeErrors.length > 0) {
        lines.push("");
        lines.push("Runtime errors:");
        lines.push(runtimeErrors.join("\n"));
    }
    ui.alert(lines.join("\n"));
}

/**
 * Transfers one route.
 * Returns "wrote", "empty", or throws on unexpected error.
 */
function transferRoute(r) {
    var sourceSS = getSpreadsheetFromInput(r.fromInput);
    var targetSS = getSpreadsheetFromInput(r.toInput);
    var srcSheet = sourceSS.getSheetByName(r.fromTabName);
    var tgtSheet = targetSS.getSheetByName(r.toTabName);

    if (!srcSheet || !tgtSheet)
        throw new Error("Source or target tab missing.");

    // --- Batched READ of all source cells via RangeList ---
    var srcA1s = r.mappings.map(function (m) {
        return m.sourceCellRef;
    });
    var srcRanges = srcSheet.getRangeList(srcA1s).getRanges();
    var srcValues = srcRanges.map(function (rg) {
        return rg.getValue();
    });

    // --- Skip if every source cell is empty ---
    var hasAnyData = srcValues.some(function (v) {
        return v !== "" && v !== null && v !== undefined;
    });
    if (!hasAnyData) return "empty";

    // --- Scan start row = header row + 1 (all equal after validation) ---
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

    // --- Find first row where all target columns are empty ---
    var targetRow = findFirstEmptyRow(tgtSheet, startRow, targetCols);

    // --- Build the row array and write in one call ---
    var minCol = targetCols[0];
    var maxCol = targetCols[targetCols.length - 1];
    var width = maxCol - minCol + 1;
    var rowArr = new Array(width).fill("");

    r.mappings.forEach(function (m, i) {
        rowArr[m.targetCol - minCol] = srcValues[i];
    });

    tgtSheet.getRange(targetRow, minCol, 1, width).setValues([rowArr]);

    // --- Batched clear of all source cells ---
    srcSheet.getRangeList(srcA1s).clearContent();

    return "wrote";
}

/**
 * Returns first row >= startRow where every column in `cols`
 * is empty. Reads the candidate block in one call.
 * Hard-capped at 10,000 rows.
 */
function findFirstEmptyRow(sheet, startRow, cols) {
    var lastRow = sheet.getLastRow();
    var scanEnd = Math.max(lastRow + 1, startRow);
    if (scanEnd - startRow > 10000) scanEnd = startRow + 10000;

    var height = Math.max(1, scanEnd - startRow + 1);
    var minCol = cols[0];
    var maxCol = cols[cols.length - 1];

    var block = sheet
        .getRange(startRow, minCol, height, maxCol - minCol + 1)
        .getValues();

    for (var r = 0; r < block.length; r++) {
        var empty = true;
        for (var c = 0; c < cols.length; c++) {
            if (block[r][cols[c] - minCol] !== "") {
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

function onOpen(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
        configSheet = ss.insertSheet("Config");
        initConfig(configSheet);
    }

    SpreadsheetApp.getUi()
        .createMenu("Transfer")
        .addItem("Transfer Data", "transferData")
        .addItem("Validate Config", "validateConfigMenu")
        .addSeparator()
        .addItem("Refresh Dropdowns", "updateSheetDropdowns")
        .addItem("Add New Route", "addNewRoute")
        .addItem("Delete Route (click inside one first)", "deleteRoute")
        .addItem("Reset & Rebuild All Routes", "resetAllRoutes")
        .addToUi();
}

function validateConfigMenu() {
    var ui = SpreadsheetApp.getUi();
    var routes;
    try {
        routes = parseConfig();
    } catch (e) {
        return ui.alert("Config error: " + e.message);
    }
    if (routes.length === 0) return ui.alert("No routes found.");

    var result = validateConfig(routes);
    var lines = ["Version " + VERSION, ""];
    lines.push("Valid: " + result.valid.length + " / " + routes.length);

    result.invalid.forEach(function (inv) {
        lines.push("");
        lines.push("❌ " + (inv.route.routeName || "Route " + (inv.index + 1)));
        inv.errors.forEach(function (e) {
            lines.push("   • " + e);
        });
    });

    if (result.invalid.length === 0) lines.push("\nAll good.");
    ui.alert(lines.join("\n"));
}

/** Draws one route block. Shared by init / add / reset. */
function drawRouteBlock(configSheet, routeNumber, startCol, currentId) {
    var titleRange = configSheet.getRange(1, startCol, 1, 2);
    titleRange.merge();
    titleRange
        .setValue("--- ROUTE " + routeNumber + " CONFIGURATION ---")
        .setFontWeight("bold")
        .setBackground("#4a86e8")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");

    var rows = [
        ["ROUTE NAME:", "Route " + routeNumber],
        ["FROM SPREADSHEET (URL or ID):", currentId],
        ["FROM TAB:", ""],
        ["TO SPREADSHEET (URL or ID):", ""],
        ["TO TAB:", ""],
    ];
    rows.forEach(function (pair, i) {
        configSheet
            .getRange(2 + i, startCol)
            .setValue(pair[0])
            .setFontWeight("bold");
        if (pair[1] !== "") {
            configSheet.getRange(2 + i, startCol + 1).setValue(pair[1]);
        }
    });

    // Hint so the worker knows this field is free-form
    configSheet
        .getRange(2, startCol + 1)
        .setNote(
            "Give this route a short descriptive name.\n" +
                "Examples:\n" +
                "  • Educational → Barangay A\n" +
                "  • Food Aid → Zone 3\n" +
                "  • House Aid → Brgy Malinis",
        );

    configSheet
        .getRange(7, startCol)
        .setValue("FROM CELL:")
        .setFontWeight("bold")
        .setBackground("#e8eaed");
    configSheet
        .getRange(7, startCol + 1)
        .setValue("TO COL OR CELL (e.g. C or C4):")
        .setFontWeight("bold")
        .setBackground("#e8eaed");

    configSheet.setColumnWidth(startCol, 240);
    configSheet.setColumnWidth(startCol + 1, 280);
    configSheet.setColumnWidth(startCol + 2, 40);
}

function initConfig(configSheet) {
    configSheet.clear();
    var currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    drawRouteBlock(configSheet, 1, 1, currentId);
}

function addNewRoute() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config") || ss.insertSheet("Config");
    var blockCols = getRouteBlockColumns(configSheet);
    var nextRouteNumber = blockCols.length + 1;
    var startCol =
        blockCols.length > 0 ? blockCols[blockCols.length - 1] + 3 : 1;

    drawRouteBlock(configSheet, nextRouteNumber, startCol, ss.getId());
    SpreadsheetApp.getUi().alert(
        "Added Route " + nextRouteNumber + " block. Existing data preserved.",
    );
}

/**
 * Deletes the route block containing the currently active cell,
 * then renumbers all remaining banners sequentially.
 *
 * Precondition: the worker clicks a cell inside the Config tab
 * (any cell within a route block), then runs this from the menu.
 */
function deleteRoute() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return ui.alert("Config sheet missing.");

    // Must be looking at the Config tab.
    var activeSheet = ss.getActiveSheet();
    if (activeSheet.getName() !== "Config") {
        return ui.alert(
            "Open the Config tab and click a cell inside the route " +
                "you want to delete, then try again.",
        );
    }

    var activeCell = configSheet.getActiveCell();
    var activeCol = activeCell.getColumn();

    // Which route block contains this cell?
    var blockCols = getRouteBlockColumns(configSheet); // 1-based start cols
    var targetStartCol = null;
    for (var i = 0; i < blockCols.length; i++) {
        var start = blockCols[i];
        if (activeCol >= start && activeCol <= start + 2) {
            targetStartCol = start;
            break;
        }
    }

    if (targetStartCol === null) {
        return ui.alert(
            "That cell isn't inside a route block. Click a cell inside " +
                "the route you want to delete.",
        );
    }

    // Get the route name for the confirmation prompt.
    var routeName =
        String(
            configSheet.getRange(2, targetStartCol + 1).getValue() || "",
        ).trim() || "Route at column " + targetStartCol;

    var confirm = ui.alert(
        "Delete Route",
        'Delete "' + routeName + '"?\n\nThis cannot be undone.',
        ui.ButtonSet.YES_NO,
    );
    if (confirm !== ui.Button.YES) return;

    // Physically delete the 3 columns of this block.
    configSheet.deleteColumns(targetStartCol, 3);

    // Renumber remaining banners sequentially.
    renumberRouteBlocks(configSheet);

    var remaining = getRouteBlockColumns(configSheet).length;
    ui.alert(
        'Route "' +
            routeName +
            '" deleted.\n' +
            remaining +
            " route(s) remaining.",
    );
}

/**
 * Rewrites every route banner in row 1 to say "--- ROUTE N CONFIGURATION ---"
 * with N sequential from left to right.
 */
function renumberRouteBlocks(configSheet) {
    var blockCols = getRouteBlockColumns(configSheet);
    for (var i = 0; i < blockCols.length; i++) {
        var startCol = blockCols[i];
        var cell = configSheet.getRange(1, startCol);
        // The banner is merged across 2 cells — resetting the top-left
        // value updates the whole merged range.
        cell.setValue("--- ROUTE " + (i + 1) + " CONFIGURATION ---");
    }
}

function resetAllRoutes() {
    var ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
        "Warning",
        "This will ERASE all Config data. Continue?",
        ui.ButtonSet.YES_NO,
    );
    if (confirm !== ui.Button.YES) return;

    var resp = ui.prompt(
        "Reset Configs",
        "How many blank routes?",
        ui.ButtonSet.OK_CANCEL,
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    var count = parseInt(resp.getResponseText().trim(), 10);
    if (isNaN(count) || count < 1) return ui.alert("Invalid number.");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config") || ss.insertSheet("Config");
    configSheet.clear();

    var currentId = ss.getId();
    for (var i = 1; i <= count; i++) {
        drawRouteBlock(configSheet, i, 1 + (i - 1) * 3, currentId);
    }
    ui.alert("Created " + count + " route(s).");
}

function updateSheetDropdowns() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return ui.alert("Config sheet missing.");

    var blockCols = getRouteBlockColumns(configSheet);
    if (blockCols.length === 0) return ui.alert("No route blocks found.");

    for (var k = 0; k < blockCols.length; k++) {
        var startCol = blockCols[k];
        var fromInput = configSheet.getRange(3, startCol + 1).getValue();
        var toInput = configSheet.getRange(5, startCol + 1).getValue();

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        if (sourceSS) applyTabDropdown(configSheet, sourceSS, 4, startCol + 1);
        if (targetSS) applyTabDropdown(configSheet, targetSS, 6, startCol + 1);
    }

    ui.alert("Updated dropdowns for " + blockCols.length + " route(s).");
}

function applyTabDropdown(configSheet, spreadsheet, row, col) {
    var names = spreadsheet.getSheets().map(function (s) {
        return s.getName();
    });
    if (names.length === 0) return;

    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(names, true)
        .setAllowInvalid(false)
        .build();

    var cell = configSheet.getRange(row, col);
    cell.setDataValidation(rule);
    if (cell.getValue() === "") cell.setValue(names[0]);
}
