const SHEET_NAME = 'Share_Revoke'; // Name of sheet
const CELL_TARGET_FILE_ID = { 'row': 4, 'col': 3 }; // Cell position of target file ID
const CELL_ACCESS_TYPE = {'row': 7, 'col': 3}; // 
const RANGE_OFFSET_ACCOUNTS_LIST = { 'row': 10, 'col': 3 }; // Row & column offset for the range of target accounts list

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Share & Revoke')
    .addItem('Share', 'shareFile')
    .addSeparator()
    .addItem('Revoke Access', 'revokeAccess')
    .addToUi();
}

function shareFile() {
  var ui = SpreadsheetApp.getUi();
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    let fileId = sheet.getRange(CELL_TARGET_FILE_ID.row, CELL_TARGET_FILE_ID.col).getValue();
    let emailAddresses = sheet
      .getRange(RANGE_OFFSET_ACCOUNTS_LIST.row, RANGE_OFFSET_ACCOUNTS_LIST.col, sheet.getLastRow() - RANGE_OFFSET_ACCOUNTS_LIST.row + 1, 1)
      .getValues()
      .flat();
    let targetFile = DriveApp.getFileById(fileId);
  } catch (error) {
    let message = errorMessage_(error);
    ui.alert(message);
  }
}
function revokeAccess() { }

/**
 * Standarized error message
 * @param {Object} e Error object returned by try-catch
 * @return {string} message Standarized error message
 */
function errorMessage_(e) {
  var message = `Error: line - ${e.lineNumber}\n${e.stack}`;
  return message;
}