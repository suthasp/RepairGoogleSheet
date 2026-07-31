/**
 * ระบบแจ้งซ่อม (Repair Request System) - Backend
 * ใช้ Google Sheet เป็นฐานข้อมูล + LINE Messaging API แจ้งเตือน
 *
 * วิธีติดตั้ง: ดูไฟล์ "คู่มือการติดตั้ง.md"
 */

// ============ CONFIG ============
const SHEET_REQUESTS = 'Requests';
const SHEET_TECHNICIANS = 'Technicians';
const SHEET_CATEGORIES = 'Categories';
const SHEET_ADMIN = 'AdminConfig';
const DRIVE_FOLDER_NAME = 'RepairSystem_Photos';

const STATUS_LIST = ['รอดำเนินการ', 'กำลังซ่อม', 'เสร็จสิ้น', 'ปิดงาน', 'ยกเลิก'];
const PRIORITY_LIST = ['ต่ำ', 'ปานกลาง', 'สูง', 'เร่งด่วน'];

// ============ SETUP (รันครั้งเดียวตอนติดตั้ง) ============
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Requests sheet ---
  let sh = ss.getSheetByName(SHEET_REQUESTS);
  if (!sh) sh = ss.insertSheet(SHEET_REQUESTS);
  sh.clear();
  sh.appendRow([
    'TicketID', 'Timestamp', 'ReporterType', 'ReporterName', 'Contact',
    'Department', 'Location', 'Category', 'Description', 'PhotoURL',
    'Status', 'Priority', 'AssignedTo', 'UpdatedAt', 'ClosedAt', 'Notes'
  ]);
  sh.setFrozenRows(1);
  sh.getRange('A1:P1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // --- Technicians sheet ---
  let shT = ss.getSheetByName(SHEET_TECHNICIANS);
  if (!shT) shT = ss.insertSheet(SHEET_TECHNICIANS);
  shT.clear();
  shT.appendRow(['Name', 'Phone', 'Active']);
  shT.appendRow(['ช่างสมชาย', '0812345678', true]);
  shT.setFrozenRows(1);
  shT.getRange('A1:C1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // --- Categories sheet ---
  let shC = ss.getSheetByName(SHEET_CATEGORIES);
  if (!shC) shC = ss.insertSheet(SHEET_CATEGORIES);
  shC.clear();
  shC.appendRow(['CategoryName']);
  ['ไฟฟ้า', 'ประปา', 'เครื่องปรับอากาศ', 'คอมพิวเตอร์/เน็ตเวิร์ก', 'เฟอร์นิเจอร์', 'อาคาร/โครงสร้าง', 'อื่นๆ']
    .forEach(c => shC.appendRow([c]));
  shC.setFrozenRows(1);
  shC.getRange('A1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  // --- AdminConfig sheet (เก็บ PIN สำหรับเข้าหน้าแอดมิน) ---
  let shA = ss.getSheetByName(SHEET_ADMIN);
  if (!shA) shA = ss.insertSheet(SHEET_ADMIN);
  shA.clear();
  shA.appendRow(['Key', 'Value']);
  shA.appendRow(['ADMIN_PIN', '1234']); // เปลี่ยนรหัสนี้ทันทีหลังติดตั้ง
  shA.setFrozenRows(1);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('ตั้งค่าฐานข้อมูลเรียบร้อย! อย่าลืมเปลี่ยน PIN ในชีท "AdminConfig"');
}

// ============ WEB APP ROUTING ============
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'form';
  let template;
  if (page === 'admin') {
    template = HtmlService.createTemplateFromFile('Admin');
  } else {
    template = HtmlService.createTemplateFromFile('Form');
  }
  const html = template.evaluate()
    .setTitle(page === 'admin' ? 'ระบบแจ้งซ่อม - Admin Dashboard' : 'แจ้งซ่อม')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============ HELPERS ============
function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getConfig_(key) {
  const sh = getSheet_(SHEET_ADMIN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function generateTicketId_() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd');
  const sh = getSheet_(SHEET_REQUESTS);
  const lastRow = sh.getLastRow();
  const seq = String(lastRow).padStart(4, '0'); // lastRow already accounts for header
  return `RP${datePart}-${seq}`;
}

// ============ PUBLIC API: FORM ============
function getFormOptions() {
  const shC = getSheet_(SHEET_CATEGORIES);
  const categories = shC.getDataRange().getValues().slice(1).map(r => r[0]).filter(String);
  return { categories: categories, priorities: PRIORITY_LIST };
}

/**
 * data = {
 *   reporterType, reporterName, contact, department, location,
 *   category, description, priority, photoBase64, photoMimeType, photoName
 * }
 */
function submitTicket(data) {
  if (!data.reporterName || !data.location || !data.category || !data.description) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน');
  }

  const sh = getSheet_(SHEET_REQUESTS);
  const ticketId = generateTicketId_();
  const now = new Date();

  let photoUrl = '';
  if (data.photoBase64) {
    photoUrl = uploadPhoto_(data.photoBase64, data.photoMimeType, data.photoName, ticketId);
  }

  sh.appendRow([
    ticketId, now, data.reporterType || 'บุคคลทั่วไป', data.reporterName,
    data.contact || '', data.department || '', data.location, data.category,
    data.description, photoUrl, STATUS_LIST[0], data.priority || PRIORITY_LIST[1],
    '', now, '', ''
  ]);

  try {
    notifyLine_(ticketId, data);
  } catch (err) {
    // ไม่ให้ error ของ LINE ทำให้การแจ้งซ่อมล้มเหลว
    console.error('LINE notify failed: ' + err);
  }

  return { ticketId: ticketId };
}

function uploadPhoto_(base64Data, mimeType, fileName, ticketId) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, `${ticketId}_${fileName}`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ============ LINE MESSAGING API ============
function notifyLine_(ticketId, data) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const targetId = PropertiesService.getScriptProperties().getProperty('LINE_TARGET_ID'); // group ID หรือ user ID
  if (!token || !targetId) return; // ยังไม่ได้ตั้งค่า LINE ให้ข้ามไป

  const message =
    `🔧 แจ้งซ่อมใหม่ [${ticketId}]\n` +
    `ผู้แจ้ง: ${data.reporterName}\n` +
    `สถานที่: ${data.location}\n` +
    `ประเภท: ${data.category}\n` +
    `รายละเอียด: ${data.description}\n` +
    `ความสำคัญ: ${data.priority || '-'}`;

  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: targetId,
    messages: [{ type: 'text', text: message }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}

// ============ ADMIN API ============
function checkAdminPin(pin) {
  const correctPin = String(getConfig_('ADMIN_PIN'));
  return String(pin) === correctPin;
}

function getAllTickets(pin) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_REQUESTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
    });
    return obj;
  });
  return rows.reverse(); // ใหม่สุดก่อน
}

function getTechnicians(pin) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_TECHNICIANS);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[2] === true)
    .map(r => r[0]);
}

function updateTicket(pin, ticketId, updates) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_REQUESTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('TicketID');
  const statusCol = headers.indexOf('Status');
  const assignedCol = headers.indexOf('AssignedTo');
  const updatedCol = headers.indexOf('UpdatedAt');
  const closedCol = headers.indexOf('ClosedAt');
  const notesCol = headers.indexOf('Notes');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === ticketId) {
      const rowNum = i + 1;
      const now = new Date();
      if (updates.status !== undefined) {
        sh.getRange(rowNum, statusCol + 1).setValue(updates.status);
        if (updates.status === 'เสร็จสิ้น' || updates.status === 'ปิดงาน') {
          sh.getRange(rowNum, closedCol + 1).setValue(now);
        }
      }
      if (updates.assignedTo !== undefined) {
        sh.getRange(rowNum, assignedCol + 1).setValue(updates.assignedTo);
      }
      if (updates.notes !== undefined) {
        sh.getRange(rowNum, notesCol + 1).setValue(updates.notes);
      }
      sh.getRange(rowNum, updatedCol + 1).setValue(now);
      return { success: true };
    }
  }
  throw new Error('ไม่พบ Ticket ID นี้');
}

function getDashboardStats(pin) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_REQUESTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('Status');
  const categoryCol = headers.indexOf('Category');
  const tsCol = headers.indexOf('Timestamp');

  const byStatus = {};
  const byCategory = {};
  let last7days = 0;
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusCol] || 'ไม่ระบุ';
    const category = row[categoryCol] || 'ไม่ระบุ';
    byStatus[status] = (byStatus[status] || 0) + 1;
    byCategory[category] = (byCategory[category] || 0) + 1;
    const ts = row[tsCol];
    if (ts instanceof Date && (now - ts) / (1000 * 60 * 60 * 24) <= 7) {
      last7days++;
    }
  }

  return {
    total: data.length - 1,
    byStatus: byStatus,
    byCategory: byCategory,
    last7days: last7days
  };
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const event = data.events[0];
  const id = event.source.groupId || event.source.roomId || event.source.userId;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AdminConfig')
    .getRange('D1').setValue(id);
  return ContentService.createTextOutput('OK');
}

