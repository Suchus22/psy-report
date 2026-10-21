/****************************************************
 * MYLIFE MEDICAL CENTER - Psy Assessment Backend
 * วิธีใช้:
 * 1. สร้าง Google Sheet ใหม่
 * 2. Extensions > Apps Script > ลบโค้ดเดิม > วางไฟล์นี้ทั้งหมด > Save
 * 3. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy Web App URL ไปวางในหน้าเว็บ index.html แท็บ "ตั้งค่า"
 ****************************************************/

const SHEET_NAME = 'Assessments';
const HEADERS = ['id','timestamp','name','position','emotion_control','empathy','dilemma','notes','diagnosis','recommendation','psychiatrist','assess_date','qa_group','qa_json'];

// คลังคำถามรายกลุ่ม (แท็บ QuestionBank)
const QB_SHEET = 'QuestionBank';
const QB_HEADERS = ['id','group_name','topic','question','guideline','active'];

// ฐานข้อมูลชื่อ-ตำแหน่ง (แท็บ IC_Database)
const IC_SHEET = 'IC_Database';
const IC_HEADERS = ['id','ic_name','position','active'];

// แจ้งเตือน Discord (ไม่ใส่ = ไม่แจ้ง)
// วิธีเอา URL: เปิด Discord → ห้องที่ต้องการ → แก้ไขช่อง (Edit Channel) → การผสานรวม (Integrations) → เว็บฮุก (Webhooks) → เว็บฮุกใหม่ → คัดลอกลิงก์ → วางแทนที่ '' ด้านล่าง
const DISCORD_WEBHOOK_URL = '';
// เลือกเคสที่จะแจ้ง: 'ALL' = ทุกเคส, หรือระบุเฉพาะ เช่น ['FAILED'] หรือ ['FAILED','CONDITIONAL PASSED']
const NOTIFY_FILTER = 'ALL';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#0ea5a4').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  // ซ่อม header ถ้ายังไม่มี
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  // ย้ายคอลัมน์ใหม่ (qa_group, qa_json) เติมท้าย header เดิมให้อัตโนมัติ
  try {
    const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    if (h.length < HEADERS.length) {
      sh.getRange(1, h.length + 1, 1, HEADERS.length - h.length).setValues([HEADERS.slice(h.length)]);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0ea5a4').setFontColor('#ffffff');
    }
  } catch (e) {}
  return sh;
}

function getQbSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(QB_SHEET);
  if (!sh) {
    sh = ss.insertSheet(QB_SHEET);
    sh.appendRow(QB_HEADERS);
    sh.getRange(1,1,1,QB_HEADERS.length).setFontWeight('bold').setBackground('#ea580c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) sh.appendRow(QB_HEADERS);
  return sh;
}

function getICSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IC_SHEET);
    sh.appendRow(IC_HEADERS);
    sh.getRange(1,1,1,IC_HEADERS.length).setFontWeight('bold').setBackground('#16a34a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) sh.appendRow(IC_HEADERS);
  return sh;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'ping') {
    return ContentService.createTextOutput('pong MYLIFE-PSY OK').setMimeType(ContentService.MimeType.TEXT);
  }
  if (action === 'list') {
    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return jsonOut_([]);
    const head = values[0];
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const o = {};
      head.forEach((h, j) => o[String(h)] = values[i][j]);
      rows.push(o);
    }
    return jsonOut_(rows);
  }
  if (action === 'qb_list') {
    const sh = getQbSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return jsonOut_([]);
    const head = values[0].map(String);
    const rows = [];
    const stamp = [];
    for (let i = 1; i < values.length; i++) {
      const o = {};
      head.forEach((h, j) => o[h] = values[i][j]);
      ['id','group_name','topic','question','guideline'].forEach(k => { if (o[k] !== undefined) o[k] = String(o[k]).trim(); });
      if (!o.id) { o.id = 'Q' + Date.now() + i; stamp.push([i + 1, o.id]); }
      if (o.active === '' || o.active === undefined || o.active === null) o.active = '1';
      if (String(o.active) !== '0' && (o.group_name || o.question)) rows.push(o);
    }
    try { stamp.forEach(s => sh.getRange(s[0], 1).setValue(s[1])); } catch (e) {}
    return jsonOut_(rows);
  }
  if (action === 'ic_list') {
    const sh = getICSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return jsonOut_([]);
    const head = values[0].map(String);
    const rows = [];
    const stamp = [];
    for (let i = 1; i < values.length; i++) {
      const o = {};
      head.forEach((h, j) => o[h] = values[i][j]);
      ['id','ic_name','position'].forEach(k => { if (o[k] !== undefined) o[k] = String(o[k]).trim(); });
      if (!o.id) { o.id = 'IC' + Date.now() + i; stamp.push([i + 1, o.id]); }
      if (o.active === '' || o.active === undefined || o.active === null) o.active = '1';
      if (String(o.active) !== '0' && o.ic_name) rows.push(o);
    }
    // คืน id ที่สุ่มให้กลับเข้าชีต เพื่อให้ปุ่มแก้ไข/ลบใช้งานได้กับแถวที่พิมพ์เอง
    try { stamp.forEach(s => sh.getRange(s[0], 1).setValue(s[1])); } catch (e) {}
    return jsonOut_(rows);
  }
  if (action === 'ic_add') {
    const sh = getICSheet_();
    const p = e.parameter || {};
    const trim = v => String(v || '').trim();
    sh.appendRow([trim(p.id) || ('IC' + Date.now()), trim(p.ic_name), trim(p.position), '1']);
    return jsonOut_({ ok: true });
  }
  if (action === 'ic_upd') {
    const sh = getICSheet_();
    const p = e.parameter || {};
    const trim = v => String(v || '').trim();
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(p.id)) {
        sh.getRange(i + 1, 2).setValue(trim(p.ic_name));
        sh.getRange(i + 1, 3).setValue(trim(p.position));
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }
  if (action === 'ic_del') {
    const sh = getICSheet_();
    const p = e.parameter || {};
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(p.id)) {
        sh.getRange(i + 1, 4).setValue('0');
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }
  if (action === 'qb_add') {
    const sh = getQbSheet_();
    const p = e.parameter || {};
    const trim = v => String(v || '').trim();
    sh.appendRow([trim(p.id) || ('Q' + Date.now()), trim(p.group_name), trim(p.topic), trim(p.question), trim(p.guideline), '1']);
    return jsonOut_({ ok: true });
  }
  if (action === 'qb_del') {
    const sh = getQbSheet_();
    const p = e.parameter || {};
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(p.id)) {
        sh.getRange(i + 1, 6).setValue('0');
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }
  if (action === 'sh_del') {
    const sh = getSheet_();
    const p = e.parameter || {};
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(p.id)) {
        sh.deleteRow(i + 1);
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }
  return ContentService.createTextOutput('MYLIFE PSY API running. Use ?action=list or ?action=ping').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const sh = getSheet_();
    let data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }
    const row = HEADERS.map(h => data[h] !== undefined ? data[h] : '');
    sh.appendRow(row);
    try { notifyDiscord_(data); } catch (e) { /* แจ้งไม่สำเร็จก็ไม่ให้บันทึกพัง */ }
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function shouldNotify_(diagnosis) {
  if (!DISCORD_WEBHOOK_URL) return false;
  if (NOTIFY_FILTER === 'ALL') return true;
  if (Array.isArray(NOTIFY_FILTER)) return NOTIFY_FILTER.indexOf(diagnosis) !== -1;
  return true;
}

function notifyDiscord_(d) {  if (!shouldNotify_(d.diagnosis)) return;
  const color = d.diagnosis === 'PASSED' ? 0x16a34a : d.diagnosis === 'CONDITIONAL PASSED' ? 0xca8a04 : 0xdc2626;
  const dot = d.diagnosis === 'PASSED' ? '🟢' : d.diagnosis === 'CONDITIONAL PASSED' ? '🟡' : '🔴';
  const payload = {
    username: 'MYLIFE Psy Assessment',
    embeds: [{
      title: dot + ' ' + (d.diagnosis || 'ผลประเมินใหม่') + ' — ' + (d.name || '-'),
      color: color,
      fields: [
        { name: 'ชื่อ (IC)', value: String(d.name || '-'), inline: true },
        { name: 'ตำแหน่ง', value: String(d.position || '-'), inline: true },
        { name: 'วันที่ประเมิน', value: String(d.assess_date || '-'), inline: true },
        { name: 'อารมณ์ / Empathy / Dilemma', value: [d.emotion_control, d.empathy, d.dilemma].filter(Boolean).join(' / ') || '-' },
        { name: 'บันทึกเพิ่มเติม', value: String(d.notes || '-').slice(0, 900) || '-' },
        { name: 'ข้อแนะนำ', value: String(d.recommendation || '-').slice(0, 900) || '-' },
        { name: 'จิตแพทย์ผู้ประเมิน', value: String(d.psychiatrist || '-'), inline: true }
      ],
      timestamp: new Date().toISOString()
    }]
  };
  UrlFetchApp.fetch(DISCORD_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// สร้างชีตย่อย QuestionBank + แถวตัวอย่าง (active=0 ไม่โชว์ในแอป ใช้ดูรูปแบบ)
// วิธีใช้: ใน Apps Script เลือก setupQuestionBank แล้วกด Run (▶) ครั้งเดียว
function setupQuestionBank() {
  const sh = getQbSheet_();
  if (sh.getLastRow() <= 1) {
    sh.appendRow(['TMP-1', 'ตัวอย่างกลุ่ม (ลบแถวนี้ได้)', 'ตัวอย่างหัวข้อ', 'ตัวอย่างคำถาม: เมื่อถูกกดดันจะทำอย่างไร', 'ผ่าน = ... / เฝ้าระวัง = ... / ไม่ผ่าน = ...', '0']);
  }
  return 'QuestionBank ready: ' + sh.getName();
}

// สร้างชีตย่อย IC_Database + แถวตัวอย่าง (active=0 ไม่โชว์ในแอป)
// วิธีใช้: ใน Apps Script เลือก setupICDatabase แล้วกด Run (▶) ครั้งเดียว
function setupICDatabase() {
  const sh = getICSheet_();
  if (sh.getLastRow() <= 1) {
    sh.appendRow(['IC-1', 'นายสมชาย ใจดี', 'นักเรียนตำรวจ', '0']);
    sh.appendRow(['IC-2', 'นางสมหญี รักดี', 'ตำรวจ', '0']);
    sh.appendRow(['IC-3', 'นายวิชัย แข็งขัน', 'แพทย์', '0']);
  }
  return 'IC_Database ready: ' + sh.getName();
}

// ฟังก์ชันทดสอบ: ใน Apps Script เลือก testDiscord แล้วกด Run (▶) ถ้า Discord ขึ้นข้อความ = Webhook ใช้ได้
function testDiscord() {
  notifyDiscord_({
    name: 'ทดสอบ ระบบ', position: 'แพทย์', emotion_control: 'ผ่าน',
    empathy: 'ปกติ', dilemma: 'ปกติ', notes: 'ทดสอบการเชื่อมต่อ',
    diagnosis: 'PASSED', recommendation: 'ทดสอบ', psychiatrist: 'ทดสอบ',
    assess_date: '2026-10-16'
  });
}
