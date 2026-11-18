/**
 * أفروسينا — تسجيل العملاء في Google Sheet
 * الخطوات في README.md
 */
const ADMIN_KEY = 'غيّر-هذا-المفتاح';       // اكتب نفس المفتاح في لوحة التحكم > الأمان والربط
const OWNER_PHONE = '+201228489636';          // اختياري: رقمك لاستقبال تنبيه واتساب تلقائي
const CALLMEBOT_KEY = '';                     // اختياري: مفتاح خدمة CallMeBot (اتركه فارغاً لإيقافه)
const SHEET = 'Customers';
const COLS = ['id','name','phone','src','first','last','visits','note','device'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(COLS);
    s.setFrozenRows(1);
    s.getRange('C:C').setNumberFormat('@');
  }
  return s;
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const digits_ = p => String(p || '').replace(/\D/g, '');
const clip_ = (v, n) => String(v || '').slice(0, n);

function findRow_(s, ph) {
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (digits_(v[i][2]) === ph) return i + 1;
  return -1;
}

function notify_(name, ph) {
  if (!CALLMEBOT_KEY) return;
  try {
    UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(OWNER_PHONE) +
      '&text=' + encodeURIComponent('عميل جديد دخل الكتالوج\nالاسم: ' + name + '\nالهاتف: +' + ph) +
      '&apikey=' + encodeURIComponent(CALLMEBOT_KEY), { muteHttpExceptions: true });
  } catch (e) {}
}

function doGet() { return out_({ ok: 1, service: 'avrosina-customers' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents), s = sheet_(), now = new Date().toISOString();

    // عام: تسجيل زائر أو طلب من نموذج
    if (b.a === 'visit') {
      const ph = digits_(b.phone);
      if (ph.length < 8 || ph.length > 15) return out_({ err: 'phone' });
      const r = findRow_(s, ph);
      if (r > 0) {
        s.getRange(r, 6).setValue(now);
        s.getRange(r, 7).setValue((Number(s.getRange(r, 7).getValue()) || 0) + 1);
        if (!s.getRange(r, 2).getValue() && b.name) s.getRange(r, 2).setValue(clip_(b.name, 80));
      } else {
        s.appendRow(['c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
          clip_(b.name, 80), ph, clip_(b.src || 'visitor', 40), now, now, 1, '', clip_(b.dev, 140)]);
        notify_(clip_(b.name, 80), ph);
      }
      return out_({ ok: 1 });
    }

    // الإدارة فقط
    if (b.k !== ADMIN_KEY) return out_({ err: 'مفتاح الإدارة غير صحيح' });

    if (b.a === 'list') {
      const v = s.getDataRange().getValues(); v.shift();
      return out_({ rows: v.map(r => ({ id: r[0], name: r[1], phone: digits_(r[2]), src: r[3], first: r[4], last: r[5], visits: r[6], note: r[7] })) });
    }
    if (b.a === 'add') {
      const x = b.row || {}, ph = digits_(x.phone);
      if (ph.length < 8 || ph.length > 15) return out_({ err: 'phone' });
      const r = findRow_(s, ph);
      if (r > 0) {
        if (x.name) s.getRange(r, 2).setValue(clip_(x.name, 80));
        if (x.note) s.getRange(r, 8).setValue(clip_(x.note, 200));
      } else {
        s.appendRow([clip_(x.id, 30) || 'c' + Date.now().toString(36), clip_(x.name, 80), ph, 'manual', now, now, 0, clip_(x.note, 200), '']);
      }
      return out_({ ok: 1 });
    }
    if (b.a === 'del') {
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(b.id)) { s.deleteRow(i + 1); break; }
      return out_({ ok: 1 });
    }
    return out_({ err: 'unknown' });
  } catch (err) {
    return out_({ err: String(err) });
  } finally {
    lock.releaseLock();
  }
}
