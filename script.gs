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
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!s || s.getLastRow() <= 1) return out({ trades: [] });
    const rows = s.getDataRange().getValues();
    rows.shift();
    const clean = rows.map((row, idx) => {
      const processed = row.map((cell, colIdx) => {
        if (cell instanceof Date) {
          if (colIdx === 0) return cell.toISOString(); // Timestamp — keep full ISO string
          if (cell.getFullYear() <= 1900) {
            // Time value stored by Sheets — format as HH:MM
            return String(cell.getHours()).padStart(2,'0') + ':' + String(cell.getMinutes()).padStart(2,'0');
          }
          // Date value — format as YYYY-MM-DD
          return cell.getFullYear() + '-' + String(cell.getMonth()+1).padStart(2,'0') + '-' + String(cell.getDate()).padStart(2,'0');
        }
        return cell;
      });
      processed.push(idx + 2); // sheet row number (1-indexed + header row) used for fast edits
      return processed;
    });
    return out({ trades: clean });
  } catch(err) { return out({ error: err.message }); }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'image') {
      const file = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType, data.filename));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return out({ success: true, fileId: file.getId(), url: 'https://drive.google.com/file/d/' + file.getId() + '/view' });
    }

    if (data.type === 'update') {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return out({ error: 'Sheet not found' });
      // Prefer row-number lookup (fast, unambiguous)
      if (data.rowNum && data.rowNum > 1 && data.rowNum <= sheet.getLastRow()) {
        sheet.getRange(data.rowNum, 1, 1, data.row.length).setValues([data.row]);
        return out({ success: true });
      }
      // Fallback: match by timestamp string (handles Date auto-conversion)
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        const cell = values[i][0];
        const cellStr = cell instanceof Date ? cell.toISOString() : String(cell);
        const sent = String(data.timestamp);
        if (cellStr === sent || cellStr.split('T')[0] === sent || sent.split('T')[0] === cellStr) {
          sheet.getRange(i + 1, 1, 1, data.row.length).setValues([data.row]);
          return out({ success: true });
        }
      }
      return out({ error: 'Trade not found' });
    }

    if (data.type === 'trade') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(SHEET_NAME);
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
