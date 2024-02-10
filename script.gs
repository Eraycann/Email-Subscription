// Script Properties
var sheetName = 'Email List';
var scriptProperties = PropertiesService.getScriptProperties();
var sheet;

function setup() {
  // Spreadsheet setup
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // If the sheet doesn't exist, create it
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue('Email');
  }

  // Save sheet ID to script properties
  scriptProperties.setProperty('sheetId', spreadsheet.getId());
}

function doPost(e) {
  try {
    // Setup if not already done
    if (!sheet) {
      setup();
    }

    // Extract email from the form data
    var email = e.parameter['Email'];

    // Check if the email is already in the sheet
    if (email && !emailExists(email)) {
      // Add a new row with the email
      sheet.appendRow([email]);
      return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'message': 'Email added successfully.' })).setMimeType(ContentService.MimeType.JSON);
    } else {
      // Email already exists
      return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': 'Email already exists.' })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'message': error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function emailExists(email) {
  // Check if the email already exists in the sheet
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === email) {
      return true;
    }
  }
  return false;
}
