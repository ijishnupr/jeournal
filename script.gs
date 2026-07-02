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
  if (e && e.parameter && e.parameter.type === 'portfolio') return getPortfolio();
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

// Reads the FD / bonds / stock sheets (never Trades/backtest) for the Portfolio tab.
// Each sheet is turned into an array of {header: value} objects using its own
// header row, so it adapts to column changes without needing index mapping.
function getPortfolio() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Force any date-looking columns in these sheets to an unambiguous ISO
    // display format first, same reasoning as the Trades Date column fix —
    // getDisplayValues() then can't drift from what the sheet actually shows.
    const fdSheet = ss.getSheetByName('FD');
    if (fdSheet && fdSheet.getLastRow() > 1) {
      fdSheet.getRange(2, 3, fdSheet.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd'); // MATURE DATE
    }
    const bondsSheet = ss.getSheetByName('bonds');
    if (bondsSheet && bondsSheet.getLastRow() > 1) {
      bondsSheet.getRange(2, 3, bondsSheet.getLastRow() - 1, 2).setNumberFormat('yyyy-mm-dd'); // Invested Date, Mature Date
    }

    const readSheet = (name, filterCol, mustBeNumber) => {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return [];
      const range = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn());
      const values = range.getValues();
      const display = range.getDisplayValues();
      const headers = values[0];
      const rows = [];
      for (let r = 1; r < values.length; r++) {
        const key = values[r][filterCol];
        const keep = mustBeNumber ? (typeof key === 'number' && key > 0) : (key !== '' && key !== null);
        if (!keep) continue;
        const obj = {};
        headers.forEach((h, c) => {
          if (!h) return;
          const cell = values[r][c];
          obj[h] = (cell instanceof Date) ? display[r][c] : cell;
        });
        rows.push(obj);
      }
      return rows;
    };

    return out({
      fd: readSheet('FD', 1, true),       // filter on AMOUNT so the "ONLY REINVESTING ALLOWED..." note row is skipped
      bonds: readSheet('bonds', 0, false), // filter on Name
      stock: readSheet('stock', 0, false), // filter on Stock
    });
  } catch (err) { return out({ error: err.message }); }
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
