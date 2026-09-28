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
        "How many Partner tables do you want to create?",
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
    generateTablesLoop(configSheet, count); // Loop to create tables
    ui.alert(
        "Successfully created " + count + " Partner configuration table(s)!",
    );
}

// FOR LOOP: Creates X number of identical partner config tables
function generateTablesLoop(configSheet, totalTables) {
    var currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var startRow = 1;

    for (var i = 1; i <= totalTables; i++) {
        // Header Banner
        var titleRange = configSheet.getRange(startRow, 1, 1, 2);
        titleRange.merge();
        titleRange.setValue("--- PARTNER " + i + " CONFIGURATION ---");
        titleRange
            .setFontWeight("bold")
            .setBackground("#4a86e8")
            .setFontColor("#ffffff")
            .setHorizontalAlignment("center");

        // Table Fields
        configSheet
            .getRange(startRow + 1, 1)
            .setValue("PARTNER NAME / LABEL:")
            .setFontWeight("bold");
        configSheet.getRange(startRow + 1, 2).setValue("Partner " + i);

        configSheet
            .getRange(startRow + 2, 1)
            .setValue("FROM SPREADSHEET (URL or ID):")
            .setFontWeight("bold");
        configSheet.getRange(startRow + 2, 2).setValue(currentId);

        configSheet
            .getRange(startRow + 3, 1)
            .setValue("FROM TAB:")
            .setFontWeight("bold");
        configSheet
            .getRange(startRow + 4, 1)
            .setValue("TO SPREADSHEET (URL or ID):")
            .setFontWeight("bold");
        configSheet
            .getRange(startRow + 5, 1)
            .setValue("TO TAB:")
            .setFontWeight("bold");

        // Mapping Headers
        configSheet
            .getRange(startRow + 7, 1)
            .setValue("FROM CELL:")
            .setFontWeight("bold")
            .setBackground("#e8eaed");
        configSheet
            .getRange(startRow + 7, 2)
            .setValue("TO HEADER:")
            .setFontWeight("bold")
            .setBackground("#e8eaed");

        // 3 Blank Mapping Rows
        configSheet
            .getRange(startRow + 8, 1, 3, 2)
            .setBorder(true, true, true, true, true, true);

        // Offset startRow for the next table iteration in the loop
        startRow += 13;
    }

    configSheet.setColumnWidth(1, 250);
    configSheet.setColumnWidth(2, 350);
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

// Finds start rows for all partner banners
function getPartnerBlockRows(configSheet) {
    var textFinder = configSheet.createTextFinder("--- PARTNER ");
    var results = textFinder.findAll();
    var rows = [];

    for (var i = 0; i < results.length; i++) {
        rows.push(results[i].getRow());
    }
    return rows;
}

// FOR LOOP: Updates dropdowns for all generated partner tables
function updateSheetDropdowns() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return;

    var blockRows = getPartnerBlockRows(configSheet);
    if (blockRows.length === 0)
        return SpreadsheetApp.getUi().alert(
            "No Partner blocks found in Config!",
        );

    for (var k = 0; k < blockRows.length; k++) {
        var startRow = blockRows[k];
        var fromInput = configSheet.getRange(startRow + 2, 2).getValue();
        var toInput = configSheet.getRange(startRow + 4, 2).getValue();

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        // FROM TAB Dropdown
        if (sourceSS) {
            var sourceNames = sourceSS.getSheets().map((s) => s.getName());
            var sourceRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(sourceNames, true)
                .setAllowInvalid(false)
                .build();
            configSheet.getRange(startRow + 3, 2).setDataValidation(sourceRule);
            if (configSheet.getRange(startRow + 3, 2).getValue() === "") {
                configSheet.getRange(startRow + 3, 2).setValue(sourceNames[0]);
            }
        }

        // TO TAB Dropdown
        if (targetSS) {
            var targetNames = targetSS.getSheets().map((s) => s.getName());
            var targetRule = SpreadsheetApp.newDataValidation()
                .requireValueInList(targetNames, true)
                .setAllowInvalid(false)
                .build();
            configSheet.getRange(startRow + 5, 2).setDataValidation(targetRule);
            if (configSheet.getRange(startRow + 5, 2).getValue() === "") {
                configSheet.getRange(startRow + 5, 2).setValue(targetNames[0]);
            }
        }
    }

    SpreadsheetApp.getUi().alert(
        "Updated dropdowns for all " + blockRows.length + " table(s)!",
    );
}

// FOR LOOP: Loops through every table block and processes the transfers
function transferData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var config = ss.getSheetByName("Config");
    if (!config) return ui.alert("Config Sheet is Missing!");

    var blockRows = getPartnerBlockRows(config);
    if (blockRows.length === 0) return ui.alert("No Partner blocks found!");

    var successCount = 0;

    // FOR LOOP: Runs data transfer for each generated table
    for (var i = 0; i < blockRows.length; i++) {
        var startRow = blockRows[i];

        var partnerName = config
            .getRange(startRow + 1, 2)
            .getValue()
            .toString()
            .trim();
        var fromInput = config.getRange(startRow + 2, 2).getValue();
        var fromTabName = config
            .getRange(startRow + 3, 2)
            .getValue()
            .toString()
            .trim();
        var toInput = config.getRange(startRow + 4, 2).getValue();
        var toTabName = config
            .getRange(startRow + 5, 2)
            .getValue()
            .toString()
            .trim();

        var sourceSS = getSpreadsheetFromInput(fromInput);
        var targetSS = getSpreadsheetFromInput(toInput);

        if (!sourceSS || !targetSS) continue; // Skip unconfigured/invalid blocks

        var fromSheet = sourceSS.getSheetByName(fromTabName);
        var toSheet = targetSS.getSheetByName(toTabName);

        if (!fromSheet || !toSheet) continue;

        // Cell mapping range for this block
        var rawMappings = config.getRange(startRow + 8, 1, 3, 2).getValues();
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
