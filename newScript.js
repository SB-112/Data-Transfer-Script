// Runs when Sheet opens or refreshes
function onOpen(e) {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = mainSheet.getSheetByName("Config");

    // Creates config if not present
    if (configSheet == null) {
        configSheet = mainSheet.insertSheet("Config");
        initConfig(configSheet);
    }

    // Add custom menu bar
    SpreadsheetApp.getUi()
        .createMenu("Transfer")
        .addItem("Refresh Tab Dropdowns", "updateSheetDropdowns")
        .addItem("Generate Partner Tables", "generateMultiplePartnerTables")
        .addItem("Transfer All Form Data", "transferData")
        .addToUi();
}

// Default initialization with 1 block
function initConfig(configSheet) {
    configSheet.clear();
    generateTablesLoop(configSheet, 1);
}

// Prompts user for how many partner tables to generate
function generateMultiplePartnerTables() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
        "Generate Partner Configs",
        "How many Partner tables do you want to create side-by-side?",
        ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    var count = parseInt(response.getResponseText().trim(), 10);
    if (isNaN(count) || count < 1) {
        return ui.alert("Please enter a valid number greater than 0.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) configSheet = ss.insertSheet("Config");

    configSheet.clear(); // Clear existing layout
    generateTablesLoop(configSheet, count);
    ui.alert(
        "Successfully created " +
            count +
            " Partner configuration table(s) side-by-side!",
    );
}

// FOR LOOP: Creates X number of partner config tables SIDE-BY-SIDE (3 columns apart)
function generateTablesLoop(configSheet, totalTables) {
    var currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var startCol = 1; // Start at Column A (1)

    for (var i = 1; i <= totalTables; i++) {
        // Header Banner
        var titleRange = configSheet.getRange(1, startCol, 1, 2);
        titleRange.merge();
        titleRange.setValue("--- PARTNER " + i + " CONFIGURATION ---");
        titleRange
            .setFontWeight("bold")
            .setBackground("#4a86e8")
            .setFontColor("#ffffff")
            .setHorizontalAlignment("center");

        // Settings
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

        // Mapping Headers
        configSheet
            .getRange(7, startCol)
            .setValue("FROM CELL:")
            .setFontWeight("bold")
            .setBackground("#e8eaed");
        configSheet
            .getRange(7, startCol + 1)
            .setValue("TO HEADER:")
            .setFontWeight("bold")
            .setBackground("#e8eaed");

        // Column widths
        configSheet.setColumnWidth(startCol, 240);
        configSheet.setColumnWidth(startCol + 1, 280);
        configSheet.setColumnWidth(startCol + 2, 40); // Spacer column between tables

        // Offset startCol by 3 columns for the next partner block (e.g. Cols 1-2, then 4-5, then 7-8)
        startCol += 3;
    }
}

// Helper function to resolve Spreadsheet object
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

// Finds start columns for all partner banners in Row 1
function getPartnerBlockColumns(configSheet) {
    var cols = [];
    var lastCol = configSheet.getLastColumn();
    if (lastCol === 0) return cols;

    var rowValues = configSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < rowValues.length; c++) {
        if (rowValues[c].toString().indexOf("--- PARTNER ") !== -1) {
            cols.push(c + 1); // 1-based column index
        }
    }
    return cols;
}

// FOR LOOP: Updates dropdowns for all side-by-side partner tables
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

        // FROM TAB Dropdown
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

        // TO TAB Dropdown
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

// FOR LOOP: Reads dynamic mappings down to the last row for each column pair
function transferData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var config = ss.getSheetByName("Config");
    if (!config) return ui.alert("Config Sheet is Missing!");

    var blockCols = getPartnerBlockColumns(config);
    if (blockCols.length === 0) return ui.alert("No Partner blocks found!");

    var successCount = 0;
    var totalRows = config.getLastRow();

    // FOR LOOP: Runs data transfer for each side-by-side block
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

        // Read all mapping rows starting from Row 8 down to the end of the sheet
        if (totalRows < 8) continue;
        var rawMappings = config
            .getRange(8, startCol, totalRows - 7, 2)
            .getValues();
        var mappings = rawMappings.filter((r) => r[0] !== "" && r[1] !== "");

        if (mappings.length === 0) continue;

        var headers = toSheet
            .getRange(1, 1, 1, toSheet.getLastColumn())
            .getValues()[0];
        var newRow = new Array(headers.length).fill("");

        mappings.forEach(([cellRef, headerName]) => {
            var colIndex = headers.indexOf(headerName);
            if (colIndex !== -1) {
                newRow[colIndex] = fromSheet.getRange(cellRef).getValue();
                fromSheet.getRange(cellRef).clearContent();
            }
        });

        toSheet.appendRow(newRow);
        successCount++;
    }

    ui.alert(
        "Successfully processed and transferred data for " +
            successCount +
            " Partner(s)!",
    );
}
