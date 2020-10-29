const SHEET_NAME = 'Share_Revoke'; // Name of sheet
const CELL_TARGET_FILE_ID = { 'row': 4, 'col': 3 }; // Cell position of target file ID
const CELL_ACCESS_TYPE = { 'row': 7, 'col': 3 }; // Cell position of access type to set to & revoke from the file
const RANGE_OFFSET_ACCOUNTS_LIST = { 'row': 10, 'col': 3 }; // Row & column offset for the range of target accounts list
const SAMPLE_SPREADSHEET_ID = '13fpOAKDFdkNqwYugPP6KkOWU56CUIh_GnxdwsTQKMro' // Spreadsheet ID of the sample spreadsheet

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
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let sheet = activeSpreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // Create a new sheet named SHEET_NAME if there were no existing sheet of the same name.
      sheet = SpreadsheetApp.openById(SAMPLE_SPREADSHEET_ID)
        .getSheetByName(SHEET_NAME)
        .copyTo(activeSpreadsheet)
        .setName(SHEET_NAME);
      throw new Error(`No existing sheet named "${SHEET_NAME}"; a new sheet was created.\nEnter the values and try again.`);
    }
    // Read the target file ID from spreadsheet
    let fileId = sheet.getRange(CELL_TARGET_FILE_ID.row, CELL_TARGET_FILE_ID.col).getValue();
    // Read the access type to grant to the accounts
    let accessType = sheet.getRange(CELL_ACCESS_TYPE.row, CELL_ACCESS_TYPE.col).getValue();
    // Read the email addresses to share the file to
    let emailAddresses = sheet
      .getRange(RANGE_OFFSET_ACCOUNTS_LIST.row, RANGE_OFFSET_ACCOUNTS_LIST.col, sheet.getLastRow() - RANGE_OFFSET_ACCOUNTS_LIST.row + 1, 1)
      .getValues()
      .flat();
    // Get the file
    let targetFile = DriveApp.getFileById(fileId);
    let fileName = targetFile.getName();
    // Final confirmation before sharing file
    let alertMessage = `Sharing "${fileName}" to the listed accounts as ${accessType}.\nAre you sure you want to continue?`;
    let alertResponse = ui.alert('Continue?', alertMessage, ui.ButtonSet.YES_NO);
    if (alertResponse != ui.Button.YES) {
      throw new Error('Canceled.');
    }
    // Share file to the email addresses
    if (accessType == 'Viewer 閲覧者') {
      targetFile.addViewers(emailAddresses);
      console.log(`Added viewers to ${fileName}:\n${emailAddresses.join('\n')}`);
    } else if (accessType == 'Commenter 閲覧者（コメント可）') {
      targetFile.addCommenters(emailAddresses);
      console.log(`Added commenters to ${fileName}:\n${emailAddresses.join('\n')}`);
    } else if (accessType == 'Editor 編集者') {
      targetFile.addEditors(emailAddresses);
      console.log(`Added editors to ${fileName}:\n${emailAddresses.join('\n')}`);
    } else {
      throw new Error(`Undefined Access Type: ${accessType}`);
    }
    ui.alert('Sharing complete.');
  } catch (error) {
    let message = errorMessage_(error);
    ui.alert(message);
  }
}

function revokeAccess() {
  var ui = SpreadsheetApp.getUi();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let sheet = activeSpreadsheet.getSheetByName(SHEET_NAME);
    // Read the target file ID from spreadsheet
    let fileId = sheet.getRange(CELL_TARGET_FILE_ID.row, CELL_TARGET_FILE_ID.col).getValue();
    // Read the access type to grant to the accounts
    let accessType = sheet.getRange(CELL_ACCESS_TYPE.row, CELL_ACCESS_TYPE.col).getValue();
    // Read the email addresses to share the file to
    let emailAddresses = sheet
      .getRange(RANGE_OFFSET_ACCOUNTS_LIST.row, RANGE_OFFSET_ACCOUNTS_LIST.col, sheet.getLastRow() - RANGE_OFFSET_ACCOUNTS_LIST.row + 1, 1)
      .getValues()
      .flat();
    // Get the file
    let targetFile = DriveApp.getFileById(fileId);
    let fileName = targetFile.getName();
    // Final confirmation
    let alertMessage = `Revoking "${accessType}" access of "${fileName}" for the listed accounts.\nAre you sure you want to continue?`;
    let alertResponse = ui.alert('Continue?', alertMessage, ui.ButtonSet.YES_NO);
    if (alertResponse != ui.Button.YES) {
      throw new Error('Canceled.');
    }
    // Share file to the email addresses
    if (accessType == 'Viewer 閲覧者') {
      emailAddresses.forEach(emailAddress => targetFile.removeViewer(emailAddress));
      console.log(`Removed viewers from ${fileName}:\n${emailAddresses.join('\n')}`);
    } else if (accessType == 'Commenter 閲覧者（コメント可）') {
      emailAddresses.forEach(emailAddress => targetFile.removeCommenter(emailAddress));
      console.log(`Removed commenters from ${fileName}:\n${emailAddresses.join('\n')}`);
    } else if (accessType == 'Editor 編集者') {
      emailAddresses.forEach(emailAddress => targetFile.removeEditor(emailAddress));
      console.log(`Removed editors from ${fileName}:\n${emailAddresses.join('\n')}`);
    } else {
      throw new Error(`Undefined Access Type: ${accessType}`);
    }
    ui.alert('Revoking access complete.');
  } catch (error) {
    let message = errorMessage_(error);
    ui.alert(message);
  }
}

/**
 * Standarized error message
 * @param {Object} e Error object returned by try-catch
 * @return {string} message Standarized error message
 */
function errorMessage_(e) {
  var message = `Error: line - ${e.lineNumber}\n${e.stack}`;
  return message;
}