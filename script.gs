// Paste this entire file into your Google Sheet's Apps Script editor
// Extensions → Apps Script → replace everything → Save → Deploy as Web App

const SHEET_NAME = 'Trades';
const HEADERS = [
  'Timestamp','Date','Entry Time','Exit Time','Time Taken',
  'Symbol','Market','Direction','Entry Price','Exit Price','Quantity',
  'P&L','P&L %','Stop Loss','Take Profit',
  'Strategy','Notes','Mistakes','Rating','Image URL','Drive File ID'
];

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const s = ss.getSheetByName(SHEET_NAME);
    if (!s || s.getLastRow() <= 1) return out({ trades: [] });
    const lastRow = s.getLastRow();
    // Force the Date column to an unambiguous ISO display format. getDisplayValues()
    // then returns exactly what the sheet shows — no Date-object/timezone round-trip
    // at all, so it can never drift from what you see in the spreadsheet.
    s.getRange(2, 2, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    const range = s.getDataRange();
    const rows = range.getValues();
    const display = range.getDisplayValues();
    rows.shift(); // remove header row
    display.shift();
    const clean = rows.map((row, idx) => {
      const processed = row.map((cell, colIdx) => {
        if (colIdx === 1) return display[idx][colIdx]; // Date — read exactly as shown in the sheet
        if (cell instanceof Date) {
          if (colIdx === 0) return cell.toISOString(); // Timestamp
          return Utilities.formatDate(cell, tz, 'HH:mm'); // time-only cells
        }
        return cell;
      });
      processed.push(idx + 2); // append sheet row number for fast edits
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
      // Fast path: update by row number
      if (data.rowNum && data.rowNum > 1 && data.rowNum <= sheet.getLastRow()) {
        sheet.getRange(data.rowNum, 2).setNumberFormat('@'); // keep Date column as plain text — avoids Sheets/Apps Script timezone conversion bugs
        sheet.getRange(data.rowNum, 1, 1, data.row.length).setValues([data.row]);
        return out({ success: true });
      }
      // Fallback: match by timestamp
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        const cell = values[i][0];
        const cellStr = cell instanceof Date ? cell.toISOString() : String(cell);
        const sent = String(data.timestamp);
        if (cellStr === sent || cellStr.split('T')[0] === sent || sent.split('T')[0] === cellStr) {
          sheet.getRange(i + 1, 2).setNumberFormat('@');
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
      const rowNum = sheet.getLastRow() + 1;
      sheet.getRange(rowNum, 2).setNumberFormat('@'); // keep Date column as plain text — avoids Sheets/Apps Script timezone conversion bugs
      sheet.getRange(rowNum, 1, 1, data.row.length).setValues([data.row]);
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
