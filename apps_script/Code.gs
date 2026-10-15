/**
 * Last War Command – site sync.
 * Paste into the Last War Tracker sheet: Extensions > Apps Script.
 * Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 *
 * The site reads and writes two tabs in this spreadsheet:
 *   Site Checklist – one row per plan step the site has marked
 *   Site Notes     – notes to consider in the next analysis
 */

const CHECK_TAB = 'Site Checklist';
const NOTES_TAB = 'Site Notes';
const CHECK_HEAD = ['Step ID', 'Done', 'Updated'];
const NOTES_HEAD = ['Note ID', 'Note', 'Added', 'Reviewed'];

function tab_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
}

function findRow_(sh, id) {
  const ids = rows_(sh).map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function state_() {
  const checks = {};
  rows_(tab_(CHECK_TAB, CHECK_HEAD)).forEach(r => { if (r[0]) checks[r[0]] = r[1] === true; });
  const notes = rows_(tab_(NOTES_TAB, NOTES_HEAD))
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      text: String(r[1]),
      at: r[2] instanceof Date ? r[2].getTime() : Number(r[2]) || Date.now(),
      reviewed: r[3] instanceof Date ? r[3].getTime() : (r[3] ? Number(r[3]) : null)
    }))
    .sort((a, b) => b.at - a.at);
  return { checks, notes };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(state_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const m = JSON.parse(e.postData.contents);
    const now = new Date();
    if (m.action === 'check') {
      const sh = tab_(CHECK_TAB, CHECK_HEAD);
      const row = findRow_(sh, m.id);
      const vals = [[String(m.id), !!m.done, now]];
      if (row > 0) sh.getRange(row, 1, 1, 3).setValues(vals);
      else sh.appendRow(vals[0]);
    } else if (m.action === 'note-add') {
      const text = String(m.text || '').slice(0, 2000);
      if (text.trim()) tab_(NOTES_TAB, NOTES_HEAD).appendRow([String(m.id), text, now, '']);
    } else if (m.action === 'note-review') {
      const sh = tab_(NOTES_TAB, NOTES_HEAD);
      const row = findRow_(sh, m.id);
      if (row > 0) sh.getRange(row, 4).setValue(m.reviewed ? now : '');
    } else if (m.action === 'note-delete') {
      const sh = tab_(NOTES_TAB, NOTES_HEAD);
      const row = findRow_(sh, m.id);
      if (row > 0) sh.deleteRow(row);
    }
    return json_(state_());
  } finally {
    lock.releaseLock();
  }
}
