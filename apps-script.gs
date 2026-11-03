// متجر غيوم: ربط الموقع بـ Google Sheets
// 1) الصق هذا الكود في Extensions > Apps Script داخل الشيت  2) شغّل setup مرة واحدة
// 3) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone) وانسخ الرابط

const SS = SpreadsheetApp.getActiveSpreadsheet();
const safe = v => { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; };

function setup() {
  const mk = (name, rows) => {
    let sh = SS.getSheetByName(name) || SS.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    return sh;
  };
  mk('الطلبات', [['الوقت', 'الرقم', 'الاسم', 'الهاتف', 'المدينة', 'العنوان', 'المنتج', 'الكمية', 'ملاحظات', 'الحالة']]);
  mk('الإعدادات', [
    ['الاسم', 'اسم المنتج'],
    ['الوصف', 'وصف المنتج'],
    ['السعر', ''],
    ['السعر قبل الخصم', ''],
    ['واتساب', ''],
    ['صورة الواجهة', ''],
    ['الصور', ''],
    ['الخيارات', ''],
    ['حسابات التواصل', '']
  ]);
  mk('الأسئلة', [['السؤال', 'الجواب']]);
  mk('الآراء', [['الاسم', 'النجوم', 'التعليق']]);
}

function doPost(e) {
  const o = JSON.parse(e.postData.contents);
  const sh = SS.getSheetByName('الطلبات');
  sh.appendRow([new Date(), o.code, o.name, o.phone, o.city, o.address, o.product, o.qty, o.notes, 'جديد'].map((v, i) => i === 0 ? v : safe(v)));
  return ContentService.createTextOutput('ok');
}

function doGet() {
  const kv = {};
  const c = SS.getSheetByName('الإعدادات');
  if (c) c.getDataRange().getValues().forEach(r => { if (r[0] && r[1] !== '') kv[String(r[0]).trim()] = String(r[1]).trim(); });
  const P = {}, out = {};
  if (kv['الاسم']) P.name = kv['الاسم'];
  if (kv['الوصف']) P.desc = kv['الوصف'];
  if (kv['السعر']) P.price = kv['السعر'];
  if (kv['السعر قبل الخصم']) P.oldPrice = kv['السعر قبل الخصم'];
  if (kv['الصور']) P.images = kv['الصور'].split(/\s+/).filter(Boolean);
  if (kv['الخيارات']) P.options = kv['الخيارات'].split('\n').map(l => {
    const p = l.split(':');
    return { label: p[0].trim(), values: (p[1] || '').split(/[،,]/).map(x => x.trim()).filter(Boolean) };
  }).filter(o => o.values.length);
  out.product = P;
  if (kv['واتساب']) out.whatsapp = kv['واتساب'].replace(/\D/g, '');
  if (kv['صورة الواجهة']) out.heroImage = kv['صورة الواجهة'];
  if (kv['حسابات التواصل']) out.social = kv['حسابات التواصل'].split('\n').map(l => {
    const p = l.split('|');
    return { name: p[0].trim(), url: (p[1] || '').trim() };
  }).filter(s => /^https?:\/\//.test(s.url));
  const rows = n => { const s = SS.getSheetByName(n); return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : []; };
  const faq = rows('الأسئلة').filter(r => r[0] && r[1]).map(r => [String(r[0]), String(r[1])]);
  if (faq.length) out.faq = faq;
  const rev = rows('الآراء').filter(r => r[0] && r[2]).map(r => ({ name: String(r[0]), stars: Math.min(5, Math.max(1, parseInt(r[1]) || 5)), text: String(r[2]) }));
  if (rev.length) out.reviews = rev;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
