/**
 * Runs automatically every time the Google Sheet is opened or refreshed.
 * Creates the Config tab if missing, refreshes sheet dropdowns, and adds a custom menu bar.
 */
function onOpen(e) {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = mainSheet.getSheetByName('Config');

  // 1. Auto-create Config sheet if missing
  if (configSheet == null) {
    configSheet = mainSheet.insertSheet('Config');
    initConfig(configSheet);
  } else {
    // Refresh dropdown values in B1 and B2 in case new tabs were added
    updateSheetDropdowns(configSheet);
  }

  // 2. Add top-level menu bar item
  SpreadsheetApp.getUi()
    .createMenu('Transfer')
    .addItem('Transfer Form Data', 'transferData')
    .addToUi();
}

/**
 * Generates initial layout, borders, labels, and dropdown rules for the Config tab.
 */
function initConfig(configSheet) {
  // Set text labels
  configSheet.getRange("A1").setValue("FROM SHEET:");
  configSheet.getRange("A2").setValue("TO SHEET:");
  configSheet.getRange("A4").setValue("FROM CELL:");
  configSheet.getRange("B4").setValue("TO COLUMN HEADER:");

  // Set header styles
  configSheet.getRange("A1:A2").setFontWeight("bold");
  configSheet.getRange("A4:B4").setFontWeight("bold").setBackground("#e8eaed");

  // Apply borders
  var tableStyling = configSheet.getRange(1, 1, 2, 2);
  tableStyling.setBorder(
    true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID
  );

  var newTableStyling = configSheet.getRange(4, 1, 15, 2);
  newTableStyling.setBorder(
    true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID
  );

  // Set initial sheet dropdowns
  updateSheetDropdowns(configSheet);

  // Adjust column widths for clean presentation
  configSheet.setColumnWidth(1, 150);
  configSheet.setColumnWidth(2, 200);
}

/**
 * Updates Data Validation dropdowns in B1 and B2 with current sheet tab names.
 */
function updateSheetDropdowns(configSheet) {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = mainSheet.getSheets();
  var sheetNames = [];

  // Gather names of all sheets except 'Config'
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (name !== 'Config') {
      sheetNames.push(name);
    }
  }

  if (sheetNames.length > 0) {
    // Build Data Validation rule
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(sheetNames, true)
      .setAllowInvalid(false)
      .build();

    configSheet.getRange("B1").setDataValidation(rule);
    configSheet.getRange("B2").setDataValidation(rule);

    // Fill defaults if currently empty
    if (configSheet.getRange("B1").getValue() === "") {
      configSheet.getRange("B1").setValue(sheetNames[0]);
    }
    if (configSheet.getRange("B2").getValue() === "") {
      configSheet.getRange("B2").setValue(sheetNames[1] || sheetNames[0]);
    }
  }
}

/**
 * Primary action function: Reads cell mappings and logs data from the source sheet to the summary sheet.
 * Assign this function to an on-sheet button!
 */
function transferData() {
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = mainSheet.getSheetByName('Config');

  if (configSheet == null) {
    SpreadsheetApp.getUi().alert("Error: 'Config' sheet was not found.");
    return;
  }

  // 1. Get sheet names from configuration dropdowns/cells
  var fromSheetName = configSheet.getRange("B1").getValue().toString().trim();
  var toSheetName = configSheet.getRange("B2").getValue().toString().trim();

  var fromSheet = mainSheet.getSheetByName(fromSheetName);
  var toSheet = mainSheet.getSheetByName(toSheetName);

  if (!fromSheet) {
    SpreadsheetApp.getUi().alert("Error: Source sheet '" + fromSheetName + "' does not exist.");
    return;
  }
  if (!toSheet) {
    SpreadsheetApp.getUi().alert("Error: Destination sheet '" + toSheetName + "' does not exist.");
    return;
  }

  // 2. Fetch mapping table starting at Row 5 in a single batch
  var lastRow = configSheet.getLastRow();
  if (lastRow < 5) {
    SpreadsheetApp.getUi().alert("Error: No cell mappings found starting at row 5.");
    return;
  }

  var rawMappings = configSheet.getRange(5, 1, lastRow - 4, 2).getValues();
  var mappings = [];

  for (var i = 0; i < rawMappings.length; i++) {
    var fromCell = rawMappings[i][0].toString().trim();
    var toHeader = rawMappings[i][1].toString().trim();
    if (fromCell !== "" && toHeader !== "") {
      mappings.push([fromCell, toHeader]);
    }
  }

  if (mappings.length === 0) {
    SpreadsheetApp.getUi().alert("Error: No valid mapping rows defined.");
    return;
  }

  // 3. Read target headers from Row 1 of the destination sheet
  var lastCol = toSheet.getLastColumn();
  if (lastCol === 0) {
    SpreadsheetApp.getUi().alert("Error: Destination sheet has no column headers in Row 1.");
    return;
  }
  var headers = toSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Initialize record array sized to target header count
  var newRow = new Array(headers.length).fill("");

  // 4. Extract values from source cells into mapped header positions
  for (var j = 0; j < mappings.length; j++) {
    var cellRef = mappings[j][0];
    var targetHeader = mappings[j][1];

    try {
      var val = fromSheet.getRange(cellRef).getValue();
      var colIndex = headers.indexOf(targetHeader);

      if (colIndex !== -1) {
        newRow[colIndex] = val;
      }
    } catch (err) {
      Logger.log("Invalid cell reference: " + cellRef);
    }
  }

  // 5. Append row to destination database sheet
  toSheet.appendRow(newRow);

  // 6. Provide confirmation alert
  SpreadsheetApp.getUi().alert("Success: Form data transferred to '" + toSheetName + "'!");
}