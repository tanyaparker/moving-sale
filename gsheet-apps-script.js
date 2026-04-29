function doPost(e) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data   = JSON.parse(e.postData.contents);
  const newRow = sheet.getLastRow() + 1;

  // Force column 6 (pickup) to plain text BEFORE writing — prevents date auto-conversion
  sheet.getRange(newRow, 6).setNumberFormat('@');

  sheet.getRange(newRow, 1, 1, 6).setValues([[
    new Date(),
    data.item   || '',
    data.name   || '',
    data.intent || '',
    data.bid    || '',
    data.pickup || ''
  ]]);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const item    = (e.parameter.item || '').toLowerCase().trim();
  const rows    = sheet.getDataRange().getValues();
  const results = rows.slice(1)
    .filter(row => row[1].toString().toLowerCase().trim() === item)
    .map(row => ({
      name:   row[2].toString(),
      intent: row[3].toString(),
      bid:    row[4].toString(),
      // New rows: stored as plain text, returned as-is
      // Old rows already converted by Sheets: format back to readable "May 8"
      pickup: row[5] instanceof Date
        ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'MMM d')
        : row[5].toString()
    }));
  return ContentService.createTextOutput(JSON.stringify(results))
    .setMimeType(ContentService.MimeType.JSON);
}
