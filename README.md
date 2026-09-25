#Data-Transfer-Script

A configurable Google Apps Script for transferring data from a form sheet to a summary sheet in Google Sheets.

The script allows users to define which cells from a form should be transferred to specific columns in a summary sheet, without modifying the script itself.

Note: This script is designed specifically for Google Sheets and Google Apps Script.

Features
Automatically creates a Config sheet when the script is first opened.
Automatically detects available sheet tabs.
Uses dropdowns to select the source and destination sheets.
Allows configurable FROM CELL → TO COLUMN HEADER mappings.
Transfers form data into a new row in the summary sheet.
Uses the same script for different form layouts by changing the configuration.
Adds a Transfer menu to the Google Sheets toolbar.
How to Use
A. Setup
Open or create your Google Spreadsheet.
Create at least two sheets:
One sheet containing the form/application data.
One sheet containing the summary/database.
Go to the GitHub repository and open the script.js file.
Copy the entire script.
Return to your Google Spreadsheet.
Open Extensions → Apps Script.
Paste the script into the Apps Script editor.
Save the project.
Return to your Google Spreadsheet and refresh the page.
A Config sheet and a Transfer menu should appear automatically.
B. Configure the Script

Open the Config sheet.

1. Select the sheets

Use the dropdowns to select:

FROM SHEET: The sheet containing the form/application data.
TO SHEET: The sheet where the transferred data will be added. 2. Create the data mappings

Starting from row 5, enter the mappings:
SAMPLE:
FROM CELL -----------TO COLUMN HEADER
F12 -----------------Full Name
F13------------------Address
F14------------------Contact Number

The FROM CELL refers to the location of the data in the form sheet.

The TO COLUMN HEADER must exactly match a column header in Row 1 of the summary sheet.

You can add as many mappings as needed.

C. Transfer the Data
Make sure the Config sheet is properly configured.
Enter the required information into the form sheet.
Open the Transfer menu in the Google Sheets toolbar.
Select Transfer Form Data.
The script will read the configured cells and add the information as a new row in the summary sheet.

LICENSE:
This project is open-source.
