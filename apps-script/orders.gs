/**
 * سوبر ماركت جي جي — استقبال الطلبات وتسجيلها في Google Sheet.
 *
 * طريقة الاستخدام (مرة واحدة): انظر README.md → "ربط الطلبات بجوجل شيت".
 * هذا السكربت منفصل عن سكربت المنتجات الحالي، ويجب أن يُنشأ داخل ملف Google Sheet جديد للطلبات.
 */

const SHEET_NAME = 'الطلبات';

const HEADERS = [
  'رقم الطلب', 'التاريخ', 'الفرع', 'الاسم', 'الهاتف',
  'طريقة الاستلام', 'العنوان', 'ملاحظات', 'الأصناف', 'الإجمالي (د.أ)', 'الحالة'
];

// الفروع المسموحة (يجب أن تطابق أسماء الفروع في index.html تماماً)
const ALLOWED_BRANCHES = ['الفرع الرئيسي - عبدون', 'فرع النخيل'];
const MAX_ITEMS = 100;

// فتح الرابط من المتصفح يؤكد فقط أن السكربت شغال
function doGet() {
  return ContentService.createTextOutput('Orders endpoint is running');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(15000);
    locked = true;

    const order = JSON.parse(e.postData.contents);
    const problem = validate_(order);
    if (problem) return json_({ ok: false, error: problem });

    const sheet = getSheet_();

    // منع تسجيل نفس الطلب مرتين (يتكرر لو ضغط الزبون "إرسال" أكثر من مرة)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
      if (ids.indexOf(order.orderId) !== -1) return json_({ ok: true, duplicate: true });
    }

    const total = order.items.reduce((sum, i) => sum + i.price * i.qty, 0);
    const itemsText = order.items.map(i =>
      i.name + ' × ' + i.qty + (i.barcode ? ' (' + i.barcode + ')' : '') +
      ' = ' + (i.price * i.qty).toFixed(2)
    ).join('\n');

    const row = [
      order.orderId,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      order.branch,
      order.name,
      order.phone,
      order.mode,
      order.address || '',
      order.notes || '',
      itemsText,
      total.toFixed(2),
      'جديد'
    ].map(clean_);

    // تنسيق "نص" يمنع تنفيذ أي معادلة قد يكتبها شخص خبيث في الاسم أو الملاحظات
    const target = sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length);
    target.setNumberFormat('@');
    target.setValues([row]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function validate_(o) {
  if (!o || typeof o !== 'object') return 'bad payload';
  if (typeof o.orderId !== 'string' || !/^JJ-[0-9A-Z-]{6,30}$/.test(o.orderId)) return 'bad orderId';
  if (ALLOWED_BRANCHES.indexOf(o.branch) === -1) return 'bad branch';
  if (typeof o.name !== 'string' || !o.name.trim() || o.name.length > 100) return 'bad name';
  if (typeof o.phone !== 'string' || !/^[+\d][\d\s\-()]{6,25}$/.test(o.phone)) return 'bad phone';
  if (o.mode !== 'توصيل' && o.mode !== 'استلام من الفرع') return 'bad mode';
  if (!Array.isArray(o.items) || o.items.length < 1 || o.items.length > MAX_ITEMS) return 'bad items';
  for (const i of o.items) {
    if (!i || typeof i.name !== 'string' || !i.name || i.name.length > 200) return 'bad item name';
    if (!Number.isInteger(i.qty) || i.qty < 1 || i.qty > 99) return 'bad qty';
    if (typeof i.price !== 'number' || !isFinite(i.price) || i.price < 0 || i.price > 100000) return 'bad price';
  }
  return null;
}

function clean_(v) {
  return String(v == null ? '' : v).slice(0, 2000);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
