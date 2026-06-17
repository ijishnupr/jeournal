// Paste this entire file into your Google Sheet's Apps Script editor
// Extensions → Apps Script → replace everything → Save → Deploy

const SHEET_NAME = 'Trades';
const HEADERS = [
  'Timestamp','Date','Entry Time','Exit Time','Time Taken',
  'Symbol','Market','Direction','Entry Price','Exit Price','Quantity',
  'Gross P&L','P&L %','Net P&L','Stop Loss','Take Profit','Charges',
  'Strategy','Notes','Mistakes','Rating','Image URL','Drive File ID'
];

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return out({ trades: [] });
    const rows = sheet.getDataRange().getValues();
    rows.shift(); // remove header row
    return out({ trades: rows });
  } catch(err) { return out({ error: err.message }); }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'image') {
      const bytes = Utilities.base64Decode(data.base64);
      const blob  = Utilities.newBlob(bytes, data.mimeType, data.filename);
      const file  = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return out({ success: true, fileId: file.getId(),
                   url: 'https://drive.google.com/file/d/' + file.getId() + '/view' });
    }

    if (data.type === 'trade') {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      let sheet   = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
      if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
      sheet.appendRow(data.row);
      return out({ success: true });
    }

    return out({ error: 'Unknown request type' });
  } catch(err) { return out({ error: err.message }); }
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
