function onOpen() {
  SpreadsheetApp.getUi().createMenu('Import Gmail Contacts')
    .addItem('Allow permissions', 'allowPermisions')
    .addItem('Import Contacts To Sheet', 'importGmailContactsToActiveSheet')
    .addToUi();
}

function allowPermisions() {
  SpreadsheetApp.getActive().toast('Permissions are allowed. Now run the menu Import Gmail Contacts->Import Contacts To Sheet');
}

function importGmailContactsToActiveSheet() {
  const activeSpreadsheet = SpreadsheetApp.getActive();
  activeSpreadsheet.toast('Importing contacts (takes a while)');
  const sheet = getOrCreateSheet('Contacts', activeSpreadsheet);
  const contacts = ContactsApp.getContacts();
  const values = contacts.map(mapContactsInfo);
  sheet.getRange(2,1,values.length,values[0].length).setValues(values);
  const header = ['Email','Name','Phone','Addresses','All emails'];
  createHeader(sheet, header);
  sheet.autoResizeColumns(1,header.length);
  sheet.sort(1);
  activeSpreadsheet.setActiveSheet(sheet); // focus sheet
  activeSpreadsheet.toast('You will find your contacts here in the sheet Contacts');
}

const createHeader = (sheet, header) => {
  const headerRange = sheet.getRange(1,1,1,header.length);
  headerRange.setValues([header]);
  sheet.setFrozenRows(1);
  headerRange.setFontWeight('bold');
}

const mapContactsInfo = contact => [
   contact.getPrimaryEmail(),
   contact.getFullName() || contact.getGivenName(),
   contact.getPhones().map(phone=>phone.getPhoneNumber()).join(", "),
   contact.getAddresses().map(a => a.getAddress()).join(", "),
   contact.getEmails().map(email=>email.getAddress()).join(", ")
 ];

const getOrCreateSheet = (sheetName, activeSpreadsheet) => {
  const existingSheet = activeSpreadsheet.getSheetByName(sheetName);
  return existingSheet || activeSpreadsheet.insertSheet(sheetName,1);
}