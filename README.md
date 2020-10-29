# Manage Access to a Google Drive File
Using Google Sheets and Google Apps Script

A simple Google Apps Script solution to share a file in Google Drive to multiple accounts at once, and to revoke that access.

## How to Use
### Share (Grant Access to File)
1. Copy the sample spreadsheet [[Sample] Share & Revoke for Google Drive](https://docs.google.com/spreadsheets/d/13fpOAKDFdkNqwYugPP6KkOWU56CUIh_GnxdwsTQKMro/edit?usp=sharin.g) from `File` > `Create Copy`.
2. Fill in the yellow cells in the spreadsheet: the ID of the Google Drive file of which to manage access, the access type to grant, and email address(es) to grant access to.
3. From the menu, `Share & Revoke` > `Share`. Note that you will be asked to authorize the script the first time you execute it.

### Revoke Access  
`Share & Revoke` > `Revoke Access`  
Note that the revoking access type must correspond to the actual permission granted to the user. Access is managed using the Google Apps Script methods `File`.[`removeViewer()`](https://developers.google.com/apps-script/reference/drive/file#removevieweremailaddress)/[`removeCommenter()`](https://developers.google.com/apps-script/reference/drive/file#removecommenteremailaddress)/[`removeEditor()`](https://developers.google.com/apps-script/reference/drive/file#removeeditoremailaddress). See the link to see how cases where access types do not match are handled.