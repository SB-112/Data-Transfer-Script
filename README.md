# Google Sheets Data Transfer Script

A configurable Google Apps Script for transferring data from a form sheet to a summary sheet in Google Sheets.

The script allows users to define which cells from a form should be transferred to specific columns in a summary sheet without modifying the script itself.

Note: This script is designed specifically for Google Sheets and Google Apps Script.

## Features

- Automatically creates a Config sheet when the spreadsheet is first opened.
- Automatically detects available sheet tabs.
- Uses dropdowns to select the source and destination sheets.
- Allows configurable FROM CELL -> TO COLUMN HEADER mappings.
- Transfers form data into a new row in the summary sheet.
- Reuses the same script for different form layouts by changing the configuration.
- Adds a Transfer menu to the Google Sheets toolbar.
- Automatically resets form input fields after data transfer.

## How to Use

### A. Setup

1. Open or create your Google Spreadsheet.
2. Create at least two tabs:
    - Form Tab
    - Summary Table Tab
3. Open the script.js file in this repository and copy the entire script.
4. In your Google Spreadsheet, navigate to Extensions -> Apps Script.
5. Paste the code into the Apps Script editor and save the project.
6. Return to your Google Spreadsheet and refresh the page.
7. A Config sheet and a Transfer menu will appear automatically.

### B. Configuration

Open the Config sheet.

1. Select the sheets using the dropdowns:
    - FROM SHEET: The sheet containing the form/application data.
    - TO SHEET: The sheet where the transferred data will be added.

2. Create the data mappings:

| FROM CELL | TO COLUMN HEADER |
| :-------- | :--------------- |
| CELL ID   | HEADER NAME      |
| CELL ID   | HEADER NAME      |
| CELL ID   | HEADER NAME      |

- The FROM CELL refers to the location of the data in the form sheet.
- The TO COLUMN HEADER must match a column header in Row 1 of the summary sheet.

### C. Transfer the Data

1. Enter the required information into the form sheet.
2. Open the Transfer menu in the Google Sheets toolbar.
3. Select Transfer Form Data.
4. The script will read the configurd cells, add the information as a new row in the summary sheet, and reset the form fields.

## License

Open-Source
