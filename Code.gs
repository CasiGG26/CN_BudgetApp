const TOKEN = 'VERANDER-MIJ';   // kies een eigen geheime code
const TZ = 'Europe/Brussels';

function sheet_(naam, kop) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(naam);
  if (!sh) { sh = ss.insertSheet(naam); sh.appendRow(kop); }
  return sh;
}
const uit_ = () => sheet_('Uitgaven', ['id','datum','bedrag','categorie','notitie','wie','type']);
const bud_ = () => sheet_('Budget', ['categorie','bedrag']);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  if (e.parameter.token !== TOKEN) return out_({error: 'Verkeerde code'});
  const uitgaven = uit_().getDataRange().getValues().slice(1).map(r => ({
    id: String(r[0]),
    datum: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]),
    bedrag: r[2], categorie: r[3], notitie: r[4], wie: r[5], type: r[6] || 'uit'
  }));
  const budgetten = {};
  bud_().getDataRange().getValues().slice(1).forEach(r => budgetten[r[0]] = r[1]);
  return out_({uitgaven, budgetten});
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.token !== TOKEN) return out_({error: 'Verkeerde code'});
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (d.actie === 'add') {
      uit_().appendRow([d.id, d.datum, d.bedrag, d.categorie, d.notitie, d.wie, d.type || 'uit']);
    } else if (d.actie === 'del') {
      const sh = uit_(), ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat().map(String);
      const i = ids.indexOf(String(d.id)); if (i > 0) sh.deleteRow(i + 1);
    } else if (d.actie === 'budget') {
      const sh = bud_(), cats = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
      const i = cats.indexOf(d.categorie);
      if (i > 0) sh.getRange(i + 1, 2).setValue(d.bedrag); else sh.appendRow([d.categorie, d.bedrag]);
    }
    return out_({ok: true});
  } finally { lock.releaseLock(); }
}
