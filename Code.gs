// Serveur commun NyiriiCollect 2.0 : à coller dans Google Apps Script, lié à un Google Sheet.
const KEY = 'CHANGEZ-MOI'; // même valeur que la « clé d'accès » de l'application
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const san = v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
function doGet() { return out({ok: true, app: 'NyiriiCollect 2.0'}); }
function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ok: false, error: 'format'}); }
  if (b.key !== KEY) return out({ok: false, error: 'clé invalide'});
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { b.records.forEach(upsert); } finally { lock.releaseLock(); }
  return out({ok: true, ids: b.records.map(r => r.id)});
}
function upsert(r) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(r.type) || ss.insertSheet(r.type);
  let hdr = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : ['id'];
  Object.keys(r.flat).forEach(k => { if (hdr.indexOf(k) < 0) hdr.push(k); });
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(x => x[0]) : [];
  const i = ids.indexOf(r.id), row = i < 0 ? sh.getLastRow() + 1 : i + 2;
  sh.getRange(row, 1, 1, hdr.length).setValues([hdr.map(h => h === 'id' ? r.id : san(r.flat[h] === undefined ? '' : r.flat[h]))]);
}
