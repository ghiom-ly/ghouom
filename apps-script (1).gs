// متجر غيوم: خلفية الموقع + لوحة التحكم عبر Google Sheets
// 1) غيِّر ADMIN_PASSWORD إلى كلمة سر قوية (حروف + أرقام + رموز، 12 حرفًا على الأقل)
// 2) شغّل setup مرة واحدة ووافق على الصلاحيات
// 3) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone) وانسخ رابط /exec
// بعد أي تعديل لاحق: Deploy > Manage deployments > Edit > New version

// كلمة السر محفوظة في Script Properties (غير مرئية في الكود)
// لضبطها: Project Settings > Script Properties > أضف: ADMIN_PASSWORD = كلمة سرك
const ADMIN_PASSWORD = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';

const SS = SpreadsheetApp.getActiveSpreadsheet();
const KEYS = ['الاسم', 'الوصف', 'السعر', 'السعر قبل الخصم', 'واتساب', 'صورة الواجهة', 'الصور', 'الخيارات', 'حسابات التواصل'];
const VALID_STATUSES = ['جديد', 'قيد التجهيز', 'تم الشحن', 'مكتمل', 'ملغي'];
const MAX_ORDERS = 500;

const safe = v => { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; };
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const trim = v => String(v == null ? '' : v).trim().slice(0, 500);

// --- حماية ضد التخمين المتكرر لكلمة السر ---
function checkBruteForce() {
  const props = PropertiesService.getScriptProperties();
  const lockUntil = parseInt(props.getProperty('bf_lock') || '0');
  if (Date.now() < lockUntil) return false; // محجوب
  return true;
}

function recordFailedAuth() {
  const props = PropertiesService.getScriptProperties();
  const fails = parseInt(props.getProperty('bf_fails') || '0') + 1;
  props.setProperty('bf_fails', String(fails));
  if (fails >= 5) {
    props.setProperty('bf_lock', String(Date.now() + 15 * 60 * 1000)); // 15 دقيقة
    props.setProperty('bf_fails', '0');
  }
}

function recordSuccessAuth() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('bf_fails', '0');
  props.setProperty('bf_lock', '0');
}

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
  mk('المنتجات', [['json']]);
  SS.getSheetByName('الإعدادات').getRange('B:B').setNumberFormat('@');
  DriveApp.getRootFolder();
}

function getProducts() {
  const sh = SS.getSheetByName('المنتجات');
  if (!sh || sh.getLastRow() < 2) return [];
  try {
    const raw = String(sh.getRange(2, 1).getValue() || '');
    return raw ? JSON.parse(raw).slice(0, 100) : [];
  } catch (_) { return []; }
}

function saveProducts(arr) {
  const sh = SS.getSheetByName('المنتجات') || SS.insertSheet('المنتجات');
  if (sh.getLastRow() === 0) sh.getRange(1, 1).setValue('json');
  sh.getRange(2, 1).setNumberFormat('@').setValue(JSON.stringify(arr.slice(0, 100)));
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
    const val = safe(String(obj[k] || '').slice(0, 5000));
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
  if (kv['الصور']) P.images = kv['الصور'].split(/\s+/).filter(Boolean).slice(0, 20);
  if (kv['الخيارات']) P.options = kv['الخيارات'].split('\n').map(l => {
    const p = l.split(':');
    return { label: p[0].trim(), values: (p[1] || '').split(/[،,]/).map(x => x.trim()).filter(Boolean) };
  }).filter(o => o.values.length).slice(0, 10);
  res.product = P;
  if (kv['واتساب']) res.whatsapp = kv['واتساب'].replace(/\D/g, '').slice(0, 20);
  if (kv['صورة الواجهة']) res.heroImage = kv['صورة الواجهة'];
  if (kv['حسابات التواصل']) res.social = kv['حسابات التواصل'].split('\n').map(l => {
    const p = l.split('|');
    return { name: p[0].trim(), url: (p[1] || '').trim() };
  }).filter(s => /^https?:\/\//.test(s.url)).slice(0, 10);
  const getRows = n => { const s = SS.getSheetByName(n); return s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : []; };
  const faq = getRows('الأسئلة').filter(r => r[0] && r[1]).map(r => [String(r[0]), String(r[1])]).slice(0, 30);
  if (faq.length) res.faq = faq;
  const rev = getRows('الآراء').filter(r => r[0] && r[2]).map(r => ({ name: String(r[0]), stars: Math.min(5, Math.max(1, parseInt(r[1]) || 5)), text: String(r[2]) })).slice(0, 50);
  if (rev.length) res.reviews = rev;
  const prods = getProducts();
  if (prods.length) res.products = prods;
  return out(res);
}

function doPost(e) {
  let o;
  try { o = JSON.parse(e.postData.contents); } catch (_) { return out({ ok: false, error: 'bad_request' }); }

  // طلب جديد من الموقع
  if (!o.action) {
    const name = trim(o.name), phone = trim(o.phone), city = trim(o.city), address = trim(o.address);
    if (!name || !phone || !city || !address) return out({ ok: false, error: 'missing_fields' });
    if (!/^[+\d][\d\s\-()]{5,}$/.test(phone)) return out({ ok: false, error: 'invalid_phone' });
    const sh = SS.getSheetByName('الطلبات');
    if (sh.getLastRow() > MAX_ORDERS + 1) return out({ ok: false, error: 'capacity' });
    sh.appendRow([new Date(), safe(o.code), safe(name), safe(phone), safe(city), safe(address), safe(trim(o.product)), safe(trim(o.qty)), safe(trim(o.notes)), 'جديد']);
    return out({ ok: true });
  }

  // أوامر لوحة التحكم (كلمة سر)
  if (!ADMIN_PASSWORD) return out({ ok: false, error: 'auth' });

  if (!checkBruteForce()) {
    return out({ ok: false, error: 'locked', message: 'حساب محجوب مؤقتًا. حاول بعد 15 دقيقة.' });
  }

  if (o.pw !== ADMIN_PASSWORD) {
    Utilities.sleep(1200);
    recordFailedAuth();
    return out({ ok: false, error: 'auth' });
  }

  recordSuccessAuth();

  if (o.action === 'orders') {
    const sh = SS.getSheetByName('الطلبات'), n = sh.getLastRow() - 1;
    const v = n > 0 ? sh.getRange(2, 1, n, 10).getValues() : [];
    const list = v.map(r => ({ time: r[0] instanceof Date ? r[0].getTime() : 0, code: r[1], name: r[2], phone: r[3], city: r[4], address: r[5], product: r[6], qty: r[7], notes: r[8], status: r[9] }));
    return out({ ok: true, orders: list.reverse().slice(0, MAX_ORDERS) });
  }

  if (o.action === 'status') {
    if (!VALID_STATUSES.includes(o.status)) return out({ ok: false, error: 'invalid_status' });
    const sh = SS.getSheetByName('الطلبات'), codes = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();
    const i = codes.findIndex(r => String(r[0]) === String(o.code));
    if (i < 1) return out({ ok: false, error: 'not_found' });
    sh.getRange(i + 1, 10).setValue(safe(o.status));
    return out({ ok: true });
  }

  if (o.action === 'save') { setKV(o.settings || {}); return out({ ok: true }); }

  if (o.action === 'saveProducts') {
    const arr = (o.products || []).slice(0, 100).map(p => ({
      name: trim(p.name || '').slice(0, 200),
      desc: trim(p.desc || '').slice(0, 1000),
      price: trim(p.price || '').slice(0, 100),
      oldPrice: trim(p.oldPrice || '').slice(0, 100),
      images: (p.images || []).slice(0, 20).map(u => String(u).slice(0, 500)),
      options: (p.options || []).slice(0, 10).map(opt => ({
        label: String(opt.label || '').slice(0, 100),
        values: (opt.values || []).slice(0, 20).map(v => String(v).slice(0, 100))
      }))
    }));
    saveProducts(arr);
    return out({ ok: true });
  }

  if (o.action === 'faq' || o.action === 'reviews') {
    const sh = SS.getSheetByName(o.action === 'faq' ? 'الأسئلة' : 'الآراء');
    const w = o.action === 'faq' ? 2 : 3;
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
    const rows = (o.rows || []).slice(0, 100).map(r => { const x = r.slice(0, w).map(v => safe(trim(v))); while (x.length < w) x.push(''); return x; });
    if (rows.length) sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows);
    return out({ ok: true });
  }

  if (o.action === 'upload') {
    if (!/^image\/(jpeg|png|webp)$/.test(o.mime || '') || !o.data || o.data.length > 5000000) return out({ ok: false, error: 'bad_image' });
    const it = DriveApp.getFoldersByName('Ghiom Images');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Ghiom Images');
    const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(o.data), o.mime, 'img_' + new Date().getTime() + '.jpg'));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return out({ ok: true, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600' });
  }

  return out({ ok: false, error: 'unknown' });
}
