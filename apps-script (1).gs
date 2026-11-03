// متجر غيوم: خلفية الموقع + لوحة التحكم عبر Google Sheets
// 1) ضع كلمة سر قوية في السطر التالي  2) شغّل setup مرة واحدة ووافق على الصلاحيات
// 3) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone) وانسخ رابط /exec
// بعد أي تعديل لاحق: Deploy > Manage deployments > Edit > New version

const ADMIN_PASSWORD = '12345678';   // مثال: 'Ghiom#2026-xyz'  (إن بقيت فارغة تتعطل لوحة التحكم)
const SS = SpreadsheetApp.getActiveSpreadsheet();
const KEYS = ['الاسم', 'الوصف', 'السعر', 'السعر قبل الخصم', 'واتساب', 'صورة الواجهة', 'الصور', 'الخيارات', 'حسابات التواصل'];
const safe = v => { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; };
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function setup() {
  const mk = (name, rows) => {
    const sh = SS.getSheetByName(name) || SS.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    return sh;
  };
  mk('الطلبات', [['الوقت', 'الرقم', 'الاسم', 'الهاتف', 'المدينة', 'العنوان', 'المنتج', 'الكمية', 'ملاحظات', 'الحالة']]);
  mk('الإعدادات', KEYS.map(k => [k, k === 'الاسم' ? 'اسم المنتج' : '']));
  mk('الأسئلة', [['السؤال', 'الجواب']]);
  mk('الآراء', [['الاسم', 'النجوم', 'التعليق']]);
  SS.getSheetByName('الإعدادات').getRange('B:B').setNumberFormat('@');
  DriveApp.getRootFolder(); // لطلب صلاحية رفع الصور
}

function kvMap() {
  const kv = {}, c = SS.getSheetByName('الإعدادات');
  if (c) c.getDataRange().getValues().forEach(r => { if (r[0] && r[1] !== '') kv[String(r[0]).trim()] = String(r[1]).trim(); });
  return kv;
}

function setKV(obj) {
  const sh = SS.getSheetByName('الإعدادات');
  const rows = sh.getDataRange().getValues();
  KEYS.forEach(k => {
    if (!(k in obj)) return;
    const val = safe(obj[k]);
    let i = rows.findIndex(r => String(r[0]).trim() === k);
    if (i < 0) { sh.appendRow([k, '']); i = sh.getLastRow() - 1; rows[i] = [k, '']; }
    sh.getRange(i + 1, 2).setNumberFormat('@').setValue(val);
  });
}

function doGet() {
  const kv = kvMap(), P = {}, res = {};
  if (kv['الاسم']) P.name = kv['الاسم'];
  if (kv['الوصف']) P.desc = kv['الوصف'];
  if (kv['السعر']) P.price = kv['السعر'];
  if (kv['السعر قبل الخصم']) P.oldPrice = kv['السعر قبل الخصم'];
  if (kv['الصور']) P.images = kv['الصور'].split(/\s+/).filter(Boolean);
  if (kv['الخيارات']) P.options = kv['الخيارات'].split('\n').map(l => {
    const p = l.split(':');
    return { label: p[0].trim(), values: (p[1] || '').split(/[،,]/).map(x => x.trim()).filter(Boolean) };
  }).filter(o => o.values.length);
  res.product = P;
  if (kv['واتساب']) res.whatsapp = kv['واتساب'].replace(/\D/g, '');
  if (kv['صورة الواجهة']) res.heroImage = kv['صورة الواجهة'];
  if (kv['حسابات التواصل']) res.social = kv['حسابات التواصل'].split('\n').map(l => {
    const p = l.split('|');
    return { name: p[0].trim(), url: (p[1] || '').trim() };
  }).filter(s => /^https?:\/\//.test(s.url));
  const rows = n => { const s = SS.getSheetByName(n); return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : []; };
  const faq = rows('الأسئلة').filter(r => r[0] && r[1]).map(r => [String(r[0]), String(r[1])]);
  if (faq.length) res.faq = faq;
  const rev = rows('الآراء').filter(r => r[0] && r[2]).map(r => ({ name: String(r[0]), stars: Math.min(5, Math.max(1, parseInt(r[1]) || 5)), text: String(r[2]) }));
  if (rev.length) res.reviews = rev;
  return out(res);
}

function doPost(e) {
  let o;
  try { o = JSON.parse(e.postData.contents); } catch (_) { return out({ ok: false }); }

  // طلب جديد من الموقع
  if (!o.action) {
    SS.getSheetByName('الطلبات').appendRow([new Date(), safe(o.code), safe(o.name), safe(o.phone), safe(o.city), safe(o.address), safe(o.product), safe(o.qty), safe(o.notes), 'جديد']);
    return out({ ok: true });
  }

  // أوامر لوحة التحكم (كلمة سر)
  if (!ADMIN_PASSWORD || o.pw !== ADMIN_PASSWORD) { Utilities.sleep(1200); return out({ ok: false, error: 'auth' }); }

  if (o.action === 'orders') {
    const sh = SS.getSheetByName('الطلبات'), n = sh.getLastRow() - 1;
    const v = n > 0 ? sh.getRange(2, 1, n, 10).getValues() : [];
    const list = v.map(r => ({ time: r[0] instanceof Date ? r[0].getTime() : 0, code: r[1], name: r[2], phone: r[3], city: r[4], address: r[5], product: r[6], qty: r[7], notes: r[8], status: r[9] }));
    return out({ ok: true, orders: list.reverse().slice(0, 300) });
  }
  if (o.action === 'status') {
    const sh = SS.getSheetByName('الطلبات'), codes = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();
    const i = codes.findIndex(r => String(r[0]) === String(o.code));
    if (i < 1) return out({ ok: false, error: 'not_found' });
    sh.getRange(i + 1, 10).setValue(safe(o.status));
    return out({ ok: true });
  }
  if (o.action === 'save') { setKV(o.settings || {}); return out({ ok: true }); }
  if (o.action === 'faq' || o.action === 'reviews') {
    const sh = SS.getSheetByName(o.action === 'faq' ? 'الأسئلة' : 'الآراء');
    const w = o.action === 'faq' ? 2 : 3;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
    const rows = (o.rows || []).map(r => { const x = r.slice(0, w).map(safe); while (x.length < w) x.push(''); return x; });
    if (rows.length) sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows);
    return out({ ok: true });
  }
  if (o.action === 'upload') {
    if (!/^image\/(jpeg|png|webp)$/.test(o.mime || '') || !o.data || o.data.length > 7000000) return out({ ok: false, error: 'bad_image' });
    const it = DriveApp.getFoldersByName('Ghiom Images');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Ghiom Images');
    const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(o.data), o.mime, Date.now() + '.jpg'));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return out({ ok: true, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600' });
  }
  return out({ ok: false, error: 'unknown' });
}
