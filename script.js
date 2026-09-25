// Runs when Sheet opens or refreshes
function onOpen(e) {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = mainSheet.getSheetByName("Config");

    // creates config if not present
    if (configSheet == null) {
        configSheet = mainSheet.insertSheet("Config");
        initConfig(configSheet);
    } else {
        updateSheetDropdowns(configSheet);
    }

    // Add menu bar
    SpreadsheetApp.getUi()
        .createMenu("Transfer")
        .addItem("Transfer Form Data", "transferData")
        .addToUi();
}

// Initializes the Config tab if newly created
function initConfig(configSheet) {
    configSheet.getRange("A1").setValue("FROM SHEET:");
    configSheet.getRange("A2").setValue("TO SHEET:");
    configSheet.getRange("A4").setValue("FROM CELL:");
    configSheet.getRange("B4").setValue("TO HEADER:");

    configSheet.getRange("A1:A2").setFontWeight("bold");
    configSheet
        .getRange("A4:B4")
        .setFontWeight("bold")
        .setBackground("#e8eaed");

    var topTable = configSheet.getRange("A1:B2");
    topTable.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID,
    );

    var rows = configSheet.getMaxRows() - 3;
    var mappingTable = configSheet.getRange(4, 1, rows, 2);
    mappingTable.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID,
    );

    updateSheetDropdowns(configSheet);

    configSheet.setColumnWidth(1, 150);
    configSheet.setColumnWidth(2, 200);
}

// Sheetdropdown
function updateSheetDropdowns(configSheet) {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = mainSheet.getSheets();
    var sheetNames = [];

    // Gather names of sheets except for Config
    for (var i = 0; i < allSheets.length; i++) {
        var name = allSheets[i].getName();
        if (name !== "Config") {
            sheetNames.push(name);
        }
    }

    if (sheetNames.length > 0) {
        var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(sheetNames, true)
            .setAllowInvalid(false)
            .build();

        configSheet.getRange("B1").setDataValidation(rule);
        configSheet.getRange("B2").setDataValidation(rule);

        if (configSheet.getRange("B1").getValue() === "") {
            configSheet.getRange("B1").setValue(sheetNames[0]);
        }
        if (configSheet.getRange("B2").getValue() === "") {
            configSheet.getRange("B2").setValue(sheetNames[1] || sheetNames[0]);
        }
    }
}

function transferData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = ss.getSheetByName("Config");
    if (!config)
        return SpreadsheetApp.getUi().alert("Config Sheet is Missing!");

    // Get source and destination sheets
    var fromSheet = ss.getSheetByName(
        config.getRange("B1").getValue().toString().trim(),
    );
    var toSheet = ss.getSheetByName(
        config.getRange("B2").getValue().toString().trim(),
    );
    if (!fromSheet || !toSheet)
        return SpreadsheetApp.getUi().alert(
            "Error: Selected sheet tab not found.",
        );

    // Fetch the mapping data starting from roww 5
    var rawMappings = config
        .getRange(5, 1, config.getLastRow() - 4, 2)
        .getValues();
    var mappings = rawMappings.filter((r) => r[0] !== "" && r[1] !== "");
    if (mappings.length === 0)
        return SpreadsheetApp.getUi().alert("Error: No mappings set.");

    // Read destination headers & create blank row template
    var headers = toSheet
        .getRange(1, 1, 1, toSheet.getLastColumn())
        .getValues()[0];
    var newRow = new Array(headers.length).fill("");

    // Map values directly into header slots & clear form fields
    mappings.forEach(([cellRef, headerName]) => {
        var colIndex = headers.indexOf(headerName);
        if (colIndex !== -1) {
            newRow[colIndex] = fromSheet.getRange(cellRef).getValue();
            fromSheet.getRange(cellRef).clearContent();
        }
    });

    // Aappend row & notify user
    toSheet.appendRow(newRow);
    SpreadsheetApp.getUi().alert(
        "Success: Form data has been transferred and reset!",
    );
}
