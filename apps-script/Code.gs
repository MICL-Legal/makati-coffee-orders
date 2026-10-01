// Coffee order receiver: the GitHub Pages form posts each order here; one row per order.
const HEADERS = ['Time', 'Name', 'Drink', 'Café', 'Qty', 'Temp', 'Size', 'Milk', 'Sugar', 'Notes'];
const FIELDS = ['name', 'drink', 'cafe', 'qty', 'temp', 'size', 'milk', 'sugar', 'notes'];

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.setName('Orders');
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Strip leading = + - @ so nobody can inject spreadsheet formulas.
function clean_(v) {
  return String(v == null ? '' : v).slice(0, 200).replace(/^[=+\-@]+/, '');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.drink) throw new Error('missing drink');
    sheet_().appendRow([new Date()].concat(FIELDS.map(f => clean_(d[f]))));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('ok');
}
