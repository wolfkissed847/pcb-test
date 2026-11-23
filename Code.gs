// วางโค้ดนี้ใน Apps Script แล้ว Deploy > Manage deployments > แก้ไข > New version
// เวอร์ชัน 2: บันทึกแบบต่อท้าย (ไม่ทับข้อมูลเก่า) เก็บประวัติการเทสทุกครั้ง ลงแท็บ "TestLog"
// ข้อมูลเดิมในแท็บแรก (รหัส/ผลเทส/หมายเหตุ/เวลา) จะถูก "คัดลอก" มาไว้ใน TestLog ครั้งแรกครั้งเดียว
// แท็บเดิมไม่ถูกแก้ไขหรือลบ — เก็บไว้เป็นสำรอง

const SHEET_ID = '1yMq1FdR3Rs2cN5dAuUfjwKNM5so7D0yFHtqbC-fGszU';
const LOG_NAME = 'TestLog';
const HEADERS = ['วันที่', 'วัน', 'เวลา', 'รหัส', 'ผลเทส', 'หมายเหตุ', 'ครั้งที่', 'ประเภท', 'id'];
const OLD_HEADERS = ['รหัส', 'ผลเทส', 'หมายเหตุ', 'เวลา'];
const DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
const FAIL = '❌';

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function dayName_(date) {
  const p = String(date).split('-').map(Number);
  if (p.length < 3 || !p[0]) return '';
  return DAYS[new Date(p[0], p[1] - 1, p[2]).getDay()];
}

// คืนแท็บ TestLog (สร้าง + คัดลอกข้อมูลเก่าครั้งแรก)
function getLog_() {
  const ss = ss_();
  let sh = ss.getSheetByName(LOG_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(LOG_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  importOld_(ss, sh);
  sh.hideColumns(HEADERS.length);   // ซ่อนคอลัมน์ id (ใช้กันบันทึกซ้ำ)
  return sh;
}

// คัดลอกข้อมูลจากแท็บเดิม (อ่านอย่างเดียว ไม่แก้แท็บเดิม)
function importOld_(ss, log) {
  const old = ss.getSheets().find(s =>
    s.getName() !== LOG_NAME && s.getLastRow() > 1 &&
    s.getRange(1, 1, 1, 4).getDisplayValues()[0].join('|') === OLD_HEADERS.join('|'));
  if (!old) return;
  const vals = old.getRange(2, 1, old.getLastRow() - 1, 4).getDisplayValues().filter(r => r[0] !== '');
  vals.sort((a, b) => String(a[3]).localeCompare(String(b[3])));
  const out = vals.map((r, i) => {
    const t = String(r[3]);
    const date = t.slice(0, 10), time = t.slice(11);
    return [date, dayName_(date), time, r[0], r[1], r[2], '1', 'ใหม่', 'old-' + (i + 1)];
  });
  if (out.length) {
    log.getRange(2, 1, out.length, HEADERS.length).setNumberFormat('@').setValues(out);
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

// ดึงประวัติทั้งหมด
function doGet() {
  return withLock_(() => {
    const sh = getLog_();
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getDisplayValues() : [];
    return json_({
      ok: true, v: 2,
      rows: rows.filter(r => r[3] !== '').map(r => ({
        date: r[0], day: r[1], time: r[2], code: r[3], result: r[4], note: r[5],
        attempt: Number(r[6]) || 1, type: r[7], id: r[8]
      }))
    });
  });
}

// เพิ่มผลเทส 1 ครั้ง (ต่อท้ายเสมอ)
function doPost(e) {
  return withLock_(() => {
    const d = JSON.parse(e.postData.contents);
    const code = String(d.code || '').trim();
    if (!code) return json_({ ok: false, error: 'no code' });
    const id = String(d.id || '');

    const sh = getLog_();
    const n = sh.getLastRow() - 1;
    let attempt = 1, hadFail = false;
    if (n > 0) {
      const data = sh.getRange(2, 4, n, 6).getDisplayValues();   // รหัส, ผลเทส, หมายเหตุ, ครั้งที่, ประเภท, id
      for (const r of data) {
        if (id && r[5] === id) return json_({ ok: true, dup: true, attempt: Number(r[3]) || 1, type: r[4] });
        if (r[0] === code) { attempt++; if (r[1] === FAIL) hadFail = true; }
      }
    }
    const type = attempt === 1 ? 'ใหม่' : (hadFail ? 'ตัวแก้' : 'เทสซ้ำ');
    const date = String(d.date || '');
    const row = sh.getLastRow() + 1;
    const rg = sh.getRange(row, 1, 1, HEADERS.length);
    rg.setNumberFormat('@').setValues([[date, dayName_(date), String(d.time || ''), code,
      d.result || '', d.note || '', String(attempt), type, id]]);
    if (type === 'ตัวแก้') rg.setBackground('#fff4e5');
    return json_({ ok: true, attempt: attempt, type: type });
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
