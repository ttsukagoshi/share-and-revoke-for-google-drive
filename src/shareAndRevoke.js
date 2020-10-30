// MIT License
// 
// Copyright (c) 2020 Taro TSUKAGOSHI
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// 
// For latest information, see https://github.com/ttsukagoshi/share-and-revoke-for-google-drive

const CONFIG = {
  'sheetName': 'Share_Revoke', // Name of sheet
  'cellTargetFileId': { 'row': 4, 'col': 3 }, // Cell position of target file ID
  'cellAccessType': { 'row': 7, 'col': 3 }, // Cell position of access type to set to & revoke from the file
  'cellAdditionalComment': { 'row': 10, 'col': 3 }, // Cell position of additional comments to add in the email notice to the shared accounts.
  'rangeOffsetAccountsList': { 'row': 13, 'col': 3 }, // Row & column offset for the range of target accounts list
  'sampleSpreadsheetId': '13fpOAKDFdkNqwYugPP6KkOWU56CUIh_GnxdwsTQKMro' // Spreadsheet ID of the sample spreadsheet
};

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
    let sheet = activeSpreadsheet.getSheetByName(CONFIG.sheetName);
    if (!sheet) {
      // Create a new sheet named CONFIG.sheetName if there were no existing sheet of the same name.
      SpreadsheetApp.openById(CONFIG.sampleSpreadsheetId)
        .getSheetByName(CONFIG.sheetName)
        .copyTo(activeSpreadsheet)
        .setName(CONFIG.sheetName);
      throw new Error(`No existing sheet named "${CONFIG.sheetName}"; a new sheet was created.\nEnter the values and try again.`);
    }
    // Read the target file ID from spreadsheet
    let fileId = sheet.getRange(CONFIG.cellTargetFileId.row, CONFIG.cellTargetFileId.col).getValue();
    // Read the access type to grant to the accounts
    let accessType = sheet.getRange(CONFIG.cellAccessType.row, CONFIG.cellAccessType.col).getValue();
    // Read the email addresses to share the file to
    let emailAddresses = sheet
      .getRange(CONFIG.rangeOffsetAccountsList.row, CONFIG.rangeOffsetAccountsList.col, sheet.getLastRow() - CONFIG.rangeOffsetAccountsList.row + 1, 1)
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
    console.log(message);
    ui.alert(message);
  }
}

function revokeAccess() {
  var ui = SpreadsheetApp.getUi();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let sheet = activeSpreadsheet.getSheetByName(CONFIG.sheetName);
    // Read the target file ID from spreadsheet
    let fileId = sheet.getRange(CONFIG.cellTargetFileId.row, CONFIG.cellTargetFileId.col).getValue();
    // Read the access type to grant to the accounts
    let accessType = sheet.getRange(CONFIG.cellAccessType.row, CONFIG.cellAccessType.col).getValue();
    // Read the email addresses to share the file to
    let emailAddresses = sheet
      .getRange(CONFIG.rangeOffsetAccountsList.row, CONFIG.rangeOffsetAccountsList.col, sheet.getLastRow() - CONFIG.rangeOffsetAccountsList.row + 1, 1)
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
    console.log(message);
    ui.alert(message);
  }
}

/**
 * Standarized error message
 * @param {Object} e Error object returned by try-catch
 * @return {string} message Standarized error message
 */
function errorMessage_(e) {
  var message = e.stack;
  return message;
}