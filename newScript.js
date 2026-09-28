// Runs when Sheet opens or refreshes
function onOpen(e) {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = mainSheet.getSheetByName("Config");

    // Creates config if not present
    if (configSheet == null) {
        configSheet = mainSheet.insertSheet("Config");
        initConfig(configSheet);
    }

    // Add menu bar
    SpreadsheetApp.getUi()
        .createMenu("Transfer")
        .addItem("Refresh Sheet Dropdowns", "updateSheetDropdowns")
        .addItem("Transfer Form Data", "transferData")
        .addToUi();
}

// Initializes the Config tab if newly created
function initConfig(configSheet) {
    configSheet.getRange("A1").setValue("FROM SPREADSHEET (URL or ID):");
    configSheet.getRange("A2").setValue("FROM TAB:");
    configSheet.getRange("A3").setValue("TO SPREADSHEET (URL or ID):");
    configSheet.getRange("A4").setValue("TO TAB:");

    configSheet.getRange("A6").setValue("FROM CELL:");
    configSheet.getRange("B6").setValue("TO HEADER:");

    configSheet.getRange("A1:A4").setFontWeight("bold");
    configSheet
        .getRange("A6:B6")
        .setFontWeight("bold")
        .setBackground("#e8eaed");

    // Optional default value: set source and target as current spreadsheet
    var currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
    configSheet.getRange("B1").setValue(currentId);
    configSheet.getRange("B3").setValue(currentId);

    configSheet.setColumnWidth(1, 220);
    configSheet.setColumnWidth(2, 300);

    updateSheetDropdowns();
}

// Helper function to resolve a Spreadsheet from either an ID or a full URL
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

// Dynamically updates dropdowns based on provided URLs or IDs
function updateSheetDropdowns() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) return;

    var fromInput = configSheet.getRange("B1").getValue();
    var toInput = configSheet.getRange("B3").getValue();

    var sourceSS = getSpreadsheetFromInput(fromInput);
    var targetSS = getSpreadsheetFromInput(toInput);

    // Set From Sheet Dropdown
    if (sourceSS) {
        var sourceNames = sourceSS.getSheets().map((s) => s.getName());
        var sourceRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(sourceNames, true)
            .setAllowInvalid(false)
            .build();
        configSheet.getRange("B2").setDataValidation(sourceRule);
    } else {
        configSheet.getRange("B2").clearDataValidation();
    }

    // Set To Sheet Dropdown
    if (targetSS) {
        var targetNames = targetSS.getSheets().map((s) => s.getName());
        var targetRule = SpreadsheetApp.newDataValidation()
            .requireValueInList(targetNames, true)
            .setAllowInvalid(false)
            .build();
        configSheet.getRange("B4").setDataValidation(targetRule);
    } else {
        configSheet.getRange("B4").clearDataValidation();
    }
}

// Executes data transfer across external spreadsheets
function transferData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = ss.getSheetByName("Config");
    if (!config)
        return SpreadsheetApp.getUi().alert("Config Sheet is Missing!");

    var fromInput = config.getRange("B1").getValue();
    var fromTabName = config.getRange("B2").getValue().toString().trim();
    var toInput = config.getRange("B3").getValue();
    var toTabName = config.getRange("B4").getValue().toString().trim();

    // Resolve external spreadsheets
    var sourceSS = getSpreadsheetFromInput(fromInput);
    var targetSS = getSpreadsheetFromInput(toInput);

    if (!sourceSS)
        return SpreadsheetApp.getUi().alert(
            "Error: Invalid 'FROM' Spreadsheet URL or ID.",
        );
    if (!targetSS)
        return SpreadsheetApp.getUi().alert(
            "Error: Invalid 'TO' Spreadsheet URL or ID.",
        );

    var fromSheet = sourceSS.getSheetByName(fromTabName);
    var toSheet = targetSS.getSheetByName(toTabName);

    if (!fromSheet)
        return SpreadsheetApp.getUi().alert(
            "Error: Source tab '" + fromTabName + "' not found.",
        );
    if (!toSheet)
        return SpreadsheetApp.getUi().alert(
            "Error: Destination tab '" + toTabName + "' not found.",
        );

    // Fetch cell mappings starting from row 7
    var lastRow = config.getLastRow();
    if (lastRow < 7)
        return SpreadsheetApp.getUi().alert("Error: No mappings set.");

    var rawMappings = config.getRange(7, 1, lastRow - 6, 2).getValues();
    var mappings = rawMappings.filter((r) => r[0] !== "" && r[1] !== "");
    if (mappings.length === 0)
        return SpreadsheetApp.getUi().alert("Error: No valid mappings set.");

    // Read destination headers & create blank row template
    var headers = toSheet
        .getRange(1, 1, 1, toSheet.getLastColumn())
        .getValues()[0];
    var newRow = new Array(headers.length).fill("");

    // Map values directly into header slots & clear form fields in source sheet
    mappings.forEach(([cellRef, headerName]) => {
        var colIndex = headers.indexOf(headerName);
        if (colIndex !== -1) {
            newRow[colIndex] = fromSheet.getRange(cellRef).getValue();
            fromSheet.getRange(cellRef).clearContent();
        }
    });

    // Append row & notify user
    toSheet.appendRow(newRow);
    SpreadsheetApp.getUi().alert(
        "Success: Form data has been transferred across files!",
    );
}
