// Runs when Sheet opens or refreshes
function onOpen(e) {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = mainSheet.getSheetByName("Config");

    if (configSheet == null) {
        configSheet = mainSheet.insertSheet("Config");
        initConfig(configSheet);
    }

    SpreadsheetApp.getUi()
        .createMenu("Transfer")
        .addItem("Refresh Tab Dropdowns", "updateSheetDropdowns")
        .addItem("Add New Partner Table", "addNewPartnerTable")
        .addItem("Reset & Rebuild All Tables", "resetAllPartnerTables")
        .addItem("Transfer All Form Data", "transferData")
        .addToUi();
}

// Default initialization
function initConfig(configSheet) {
    configSheet.clear();
    generateTablesLoop(configSheet, 1);
}

// Appends ONE new Partner block to the right without erasing existing inputs
function addNewPartnerTable() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) configSheet = ss.insertSheet("Config");

    var blockCols = getPartnerBlockColumns(configSheet);
    var nextPartnerNumber = blockCols.length + 1;
    var startCol =
        blockCols.length > 0 ? blockCols[blockCols.length - 1] + 3 : 1;
    var currentId = ss.getId();

    // Draw ONLY the new block
    var titleRange = configSheet.getRange(1, startCol, 1, 2);
    titleRange.merge();
    titleRange.setValue(
        "--- PARTNER " + nextPartnerNumber + " CONFIGURATION ---",
    );
    titleRange
        .setFontWeight("bold")
        .setBackground("#4a86e8")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");

    configSheet
        .getRange(2, startCol)
        .setValue("PARTNER NAME / LABEL:")
        .setFontWeight("bold");
    configSheet
        .getRange(2, startCol + 1)
        .setValue("Partner " + nextPartnerNumber);

    configSheet
        .getRange(3, startCol)
        .setValue("FROM SPREADSHEET (URL or ID):")
        .setFontWeight("bold");
    configSheet.getRange(3, startCol + 1).setValue(currentId);

    configSheet
        .getRange(4, startCol)
        .setValue("FROM TAB:")
        .setFontWeight("bold");
    configSheet
        .getRange(5, startCol)
        .setValue("TO SPREADSHEET (URL or ID):")
        .setFontWeight("bold");
    configSheet.getRange(6, startCol).setValue("TO TAB:").setFontWeight("bold");

    configSheet
        .getRange(7, startCol)
        .setValue("FROM CELL:")
        .setFontWeight("bold")
        .setBackground("#e8eaed");
    configSheet
        .getRange(7, startCol + 1)
        .setValue("TO COL OR CELL (e.g. C or C5):")
        .setFontWeight("bold")
        .setBackground("#e8eaed");

    configSheet.setColumnWidth(startCol, 240);
    configSheet.setColumnWidth(startCol + 1, 280);
    configSheet.setColumnWidth(startCol + 2, 40);

    SpreadsheetApp.getUi().alert(
        "Added Partner " +
            nextPartnerNumber +
            " block! Existing table data was preserved.",
    );
}

// Full reset/rebuild option (destroys current layout)
function resetAllPartnerTables() {
    var ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
        "Warning",
        "This will ERASE all current configurations and tables. Continue?",
        ui.ButtonSet.YES_NO,
    );
    if (confirm !== ui.Button.YES) return;

    var response = ui.prompt(
        "Reset Configs",
        "How many blank Partner tables do you want to start with?",
        ui.ButtonSet.OK_CANCEL,
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;

    var count = parseInt(response.getResponseText().trim(), 10);
    if (isNaN(count) || count < 1) return ui.alert("Invalid number.");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) configSheet = ss.insertSheet("Config");

    configSheet.clear();
    generateTablesLoop(configSheet, count);
    ui.alert(
        "Successfully created " +
            count +
            " Partner configuration table(s) side-by-side!",
    );
}

// FOR LOOP: Helper to draw partner tables side-by-side
function generateTablesLoop(configSheet, totalTables) {
    var currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var startCol = 1;

    for (var i = 1; i <= totalTables; i++) {
        var titleRange = configSheet.getRange(1, startCol, 1, 2);
        titleRange.merge();
        titleRange.setValue("--- PARTNER " + i + " CONFIGURATION ---");
        titleRange
            .setFontWeight("bold")
            .setBackground("#4a86e8")
            .setFontColor("#ffffff")
            .setHorizontalAlignment("center");

        configSheet
            .getRange(2, startCol)
            .setValue("PARTNER NAME / LABEL:")
            .setFontWeight("bold");
        configSheet.getRange(2, startCol + 1).setValue("Partner " + i);

        configSheet
            .getRange(3, startCol)
            .setValue("FROM SPREADSHEET (URL or ID):")
            .setFontWeight("bold");
        configSheet.getRange(3, startCol + 1).setValue(currentId);

        configSheet
            .getRange(4, startCol)
            .setValue("FROM TAB:")
            .setFontWeight("bold");
        configSheet
            .getRange(5, startCol)
            .setValue("TO SPREADSHEET (URL or ID):")
            .setFontWeight("bold");
        configSheet
            .getRange(6, startCol)
            .setValue("TO TAB:")
            .setFontWeight("bold");

        configSheet
            .getRange(7, startCol)
            .setValue("FROM CELL:")
            .setFontWeight("bold")
            .setBackground("#e8eaed");
        configSheet
            .getRange(7, startCol + 1)
            .setValue("TO COL OR CELL (e.g. C or C5):")
            .setFontWeight("bold")
            .setBackground("#e8eaed");

        configSheet.setColumnWidth(startCol, 240);
        configSheet.setColumnWidth(startCol + 1, 280);
        configSheet.setColumnWidth(startCol + 2, 40);

        startCol += 3;
    }
}

// Helper to open spreadsheet by URL or ID
function getSpreadsheetFromInput(input) {
    if (!input) return null;
    var str = input.toString().trim();
    if (str === "") return null;

    try {
        if (str.startsWith("http://") || str.startsWith("https://")) {
            return SpreadsheetApp.openByUrl(str);
        } else {
            return SpreadsheetApp.openById(str);
        }
    } catch (e) {
        return null;
    }
}

// Finds start columns for all partner banners
function getPartnerBlockColumns(configSheet) {
    var cols = [];
    var lastCol = configSheet.getLastColumn();
    if (lastCol === 0) return cols;

    var rowValues = configSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < rowValues.length; c++) {
        if (rowValues[c].toString().indexOf("--- PARTNER ") !== -1) {
            cols.push(c + 1);
        }
    }
    return cols;
}

// Updates tab dropdowns dynamically for all partner blocks
function updateSheetDropdowns() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return;

    var blockCols = getPartnerBlockColumns(configSheet);
    if (blockCols.length === 0)
        return SpreadsheetApp.getUi().alert(
            "No Partner blocks found in Config!",
        );

    for (var k = 0; k < blockCols.length; k++) {
        var startCol = blockCols[k];
        var fromInput = configSheet.getRange(3, startCol + 1).getValue();
        var toInput = configSheet.getRange(5, startCol + 1).getValue();

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        if (sourceSS) {
            var sourceNames = sourceSS.getSheets().map((s) => s.getName());
            var sourceRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(sourceNames, true)
                .setAllowInvalid(false)
                .build();
            configSheet.getRange(4, startCol + 1).setDataValidation(sourceRule);
            if (configSheet.getRange(4, startCol + 1).getValue() === "") {
                configSheet.getRange(4, startCol + 1).setValue(sourceNames[0]);
            }
        }

        if (targetSS) {
            var targetNames = targetSS.getSheets().map((s) => s.getName());
            var targetRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(targetNames, true)
                .setAllowInvalid(false)
                .build();
            configSheet.getRange(6, startCol + 1).setDataValidation(targetRule);
            if (configSheet.getRange(6, startCol + 1).getValue() === "") {
                configSheet.getRange(6, startCol + 1).setValue(targetNames[0]);
            }
        }
    }

    SpreadsheetApp.getUi().alert(
        "Updated dropdowns for all " + blockCols.length + " table(s)!",
    );
}

// Converts Column Letter (e.g. "A", "C") or Cell Ref (e.g. "C5") to Column Number & Header Row
function parseTargetReference(refStr) {
    if (!refStr) return null;
    var str = refStr.toString().trim().toUpperCase();

    var match = str.match(/^([A-Z]+)(\d*)$/);
    if (!match) return null;

    var colLetters = match[1];
    var startRow = match[2] ? parseInt(match[2], 10) : 1;

    var colIndex = 0;
    for (var i = 0; i < colLetters.length; i++) {
        colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 64);
    }

    return { colIndex: colIndex, headerRow: startRow };
}

// Main Transfer Function using Direct Column/Cell Placement
function transferData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var config = ss.getSheetByName("Config");
    if (!config) return ui.alert("Config Sheet is Missing!");

    var blockCols = getPartnerBlockColumns(config);
    if (blockCols.length === 0) return ui.alert("No Partner blocks found!");

    var successCount = 0;
    var totalRows = config.getLastRow();

    for (var i = 0; i < blockCols.length; i++) {
        var startCol = blockCols[i];

        var partnerName = config
            .getRange(2, startCol + 1)
            .getValue()
            .toString()
            .trim();
        var fromInput = config.getRange(3, startCol + 1).getValue();
        var fromTabName = config
            .getRange(4, startCol + 1)
            .getValue()
            .toString()
            .trim();
        var toInput = config.getRange(5, startCol + 1).getValue();
        var toTabName = config
            .getRange(6, startCol + 1)
            .getValue()
            .toString()
            .trim();

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        if (!sourceSS || !targetSS) continue;

        var fromSheet = sourceSS.getSheetByName(fromTabName);
        var toSheet = targetSS.getSheetByName(toTabName);

        if (!fromSheet || !toSheet) continue;

        // Read mapping entries starting from Row 8
        if (totalRows < 8) continue;
        var rawMappings = config
            .getRange(8, startCol, totalRows - 7, 2)
            .getValues();
        var mappings = rawMappings.filter((r) => r[0] !== "" && r[1] !== "");

        if (mappings.length === 0) continue;

        var baseHeaderRow = 1;
        var parsedMappings = [];

        mappings.forEach(([sourceCell, targetRef]) => {
            var parsed = parseTargetReference(targetRef);
            if (parsed) {
                if (parsed.headerRow > baseHeaderRow)
                    baseHeaderRow = parsed.headerRow;
                parsedMappings.push({
                    sourceCell: sourceCell.toString().trim(),
                    targetCol: parsed.colIndex,
                });
            }
        });

        if (parsedMappings.length === 0) continue;

        // Find first empty row in target column after the header row
        var targetRow = baseHeaderRow + 1;
        var firstTargetCol = parsedMappings[0].targetCol;

        while (toSheet.getRange(targetRow, firstTargetCol).getValue() !== "") {
            targetRow++;
        }

        // Write values to target
        parsedMappings.forEach((item) => {
            var val = fromSheet.getRange(item.sourceCell).getValue();
            toSheet.getRange(targetRow, item.targetCol).setValue(val);
        });

        // Clear source form fields
        parsedMappings.forEach((item) => {
            fromSheet.getRange(item.sourceCell).clearContent();
        });

        successCount++;
    }

    ui.alert(
        "Successfully processed and transferred data for " +
            successCount +
            " Partner(s)!",
    );
}
