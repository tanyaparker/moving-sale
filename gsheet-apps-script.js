function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateItemsSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('Items');
  if (!sheet) {
    sheet = ss.insertSheet('Items');
    sheet.appendRow(['id','name','price','description','imageURL','sold','emoji','color']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action || 'responses';
  if (action === 'items') return getItems();
  return getResponses(e.parameter.item || '');
}

function getItems() {
  const sheet = getOrCreateItemsSheet();
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return json([]);
  return json(rows.slice(1).map(row => ({
    id:          row[0].toString(),
    name:        row[1].toString(),
    price:       Number(row[2]),
    description: row[3].toString(),
    imageURL:   row[4].toString(),
    sold:        row[5] === true || row[5] === 'TRUE',
    emoji:       row[6].toString(),
    color:       row[7].toString()
  })));
}

function getResponses(itemName) {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const item    = itemName.toLowerCase().trim();
  const rows    = sheet.getDataRange().getValues();
  return json(rows.slice(1)
    .filter(r => r[1].toString().toLowerCase().trim() === item)
    .map(r => ({
      name:   r[2].toString(),
      intent: r[3].toString(),
      bid:    r[4].toString(),
      pickup: r[5] instanceof Date
        ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'MMM d')
        : r[5].toString()
    })));
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  switch (data.action) {
    case 'addItem':    return addItem(data);
    case 'updateSold': return updateSold(data);
    case 'deleteItem': return deleteItem(data);
    default:           return addResponse(data);
  }
}

function addItem(data) {
  const sheet = getOrCreateItemsSheet();
  const id    = Date.now().toString();
  sheet.appendRow([id, data.name||'', Number(data.price)||0,
    data.description||'', data.imageURL||'', false,
    data.emoji||'🛋️', data.color||'#fad4e0']);
  return json({ status: 'ok', id });
}

function updateSold(data) {
  const sheet = getOrCreateItemsSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === data.id.toString()) {
      sheet.getRange(i + 1, 6).setValue(data.sold);
      return json({ status: 'ok' });
    }
  }
  return json({ status: 'error', message: 'Item not found' });
}

function deleteItem(data) {
  const sheet = getOrCreateItemsSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === data.id.toString()) {
      sheet.deleteRow(i + 1);
      return json({ status: 'ok' });
    }
  }
  return json({ status: 'error', message: 'Item not found' });
}

function addResponse(data) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 6).setNumberFormat('@');
  sheet.getRange(newRow, 1, 1, 6).setValues([[
    new Date(), data.item||'', data.name||'',
    data.intent||'', data.bid||'', data.pickup||''
  ]]);
  return json({ status: 'ok' });
}
