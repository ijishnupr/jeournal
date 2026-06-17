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
    // Normalise Date objects → plain strings so the browser inputs accept them
    // Sheets stores dates as Date and times as Date with year=1899 (epoch)
    const clean = rows.map(row => row.map(cell => {
      if (cell instanceof Date) {
        if (cell.getFullYear() <= 1900) {
          // Time value — format as HH:MM
          const h = String(cell.getHours()).padStart(2,'0');
          const mn = String(cell.getMinutes()).padStart(2,'0');
          return h + ':' + mn;
        }
        // Date value — format as YYYY-MM-DD
        const y=cell.getFullYear(), m=String(cell.getMonth()+1).padStart(2,'0'), d=String(cell.getDate()).padStart(2,'0');
        return y+'-'+m+'-'+d;
      }
      return cell;
    }));
    return out({ trades: clean });
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

    if (data.type === 'update') {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) return out({ error: 'Sheet not found' });
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(data.timestamp)) {
          sheet.getRange(i + 1, 1, 1, data.row.length).setValues([data.row]);
          return out({ success: true });
        }
      }
      return out({ error: 'Trade not found — it may have been deleted' });
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
