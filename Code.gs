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
    for (let i = 1; i < values.length; i++) {
      const o = {};
      head.forEach((h, j) => o[h] = values[i][j]);
      if (String(o.active) !== '0') rows.push(o);
    }
    return jsonOut_(rows);
  }
  if (action === 'qb_add') {
    const sh = getQbSheet_();
    const p = e.parameter || {};
    sh.appendRow([p.id || ('Q' + Date.now()), p.group_name || '', p.topic || '', p.question || '', p.guideline || '', '1']);
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

// ฟังก์ชันทดสอบ: ใน Apps Script เลือก testDiscord แล้วกด Run (▶) ถ้า Discord ขึ้นข้อความ = Webhook ใช้ได้
function testDiscord() {
  notifyDiscord_({
    name: 'ทดสอบ ระบบ', position: 'แพทย์', emotion_control: 'ผ่าน',
    empathy: 'ปกติ', dilemma: 'ปกติ', notes: 'ทดสอบการเชื่อมต่อ',
    diagnosis: 'PASSED', recommendation: 'ทดสอบ', psychiatrist: 'ทดสอบ',
    assess_date: '2026-10-16'
  });
}
