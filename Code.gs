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
const SHEET_LOCATIONS = 'Locations';
const SHEET_ADMIN = 'AdminConfig';
const DRIVE_FOLDER_NAME = 'RepairSystem_Photos';

// ลำดับคอลัมน์ในชีท Requests — SCCD คือเลขอ้างอิงงานจากระบบ SCCD/ITSM
const REQUEST_HEADERS = [
  'TicketID', 'SCCD', 'Timestamp', 'ReporterType', 'ReporterName', 'Contact',
  'Department', 'Location', 'Category', 'Description', 'PhotoURL',
  'Status', 'Priority', 'AssignedTo', 'UpdatedAt', 'ClosedAt', 'Notes'
];

const STATUS_LIST = ['รอดำเนินการ', 'กำลังซ่อม', 'เสร็จสิ้น', 'ปิดงาน', 'ยกเลิก'];
const PRIORITY_LIST = ['ต่ำ', 'ปานกลาง', 'สูง', 'เร่งด่วน'];

// รายชื่อสถานที่ (รหัสสาขา/หน่วยงาน) — แก้เพิ่ม/ลบได้ในชีท "Locations" ไม่ต้องแก้โค้ด
const LOCATION_LIST = [
  'AGN-ASD', 'AGN-BBT', 'AGN-BGC', 'AGN-BPL', 'AGN-CSW', 'AGN-DNM', 'AGN-LTP',
  'AGN-PKK', 'AGN-PSN', 'AGN-PSP', 'AGN-PTT', 'AGN-RBN', 'AGN-RIT', 'AGN-TMM', 'AGN-TYB',
  'CLS-SKA', 'CLS-STN',
  'CN-MTG', 'CN-PBI', 'CN-TTW1', 'CN-TYB', 'CN-TYN',
  'RN-AYT', 'RN-CBR', 'RN-CMI', 'RN-KKN', 'RN-NKR', 'RN-NKT', 'RN-PSN', 'RN-SKA', 'RN-SRT',
  'TOC-BSN', 'TOC-HYI', 'TOC-KKN', 'TOC-LPN', 'TOC-NMA', 'TOC-PLK', 'TOC-PSI',
  'TOC-RST', 'TOC-SNI', 'TOC-SNK'
];

// ============ SETUP (รันครั้งเดียวตอนติดตั้ง) ============
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Requests sheet ---
  let sh = ss.getSheetByName(SHEET_REQUESTS);
  if (!sh) sh = ss.insertSheet(SHEET_REQUESTS);
  sh.clear();
  sh.appendRow(REQUEST_HEADERS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, REQUEST_HEADERS.length)
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

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

  // --- Locations sheet ---
  ensureLocationsSheet_();

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

/**
 * สร้างชีท "Locations" + ใส่รายชื่อสถานที่ตั้งต้น
 * ปลอดภัยกับข้อมูลเดิม: ถ้ามีชีทและมีข้อมูลอยู่แล้วจะไม่เขียนทับ
 * ใช้สำหรับระบบที่ติดตั้งไปแล้ว — เลือกฟังก์ชันนี้แล้วกด Run (ห้ามรัน setupSpreadsheet ซ้ำ เพราะจะล้างข้อมูลแจ้งซ่อมทั้งหมด)
 */
function setupLocations() {
  const added = ensureLocationsSheet_();
  const msg = added > 0
    ? `เพิ่มสถานที่ ${added} รายการในชีท "Locations" เรียบร้อย`
    : 'ชีท "Locations" มีข้อมูลอยู่แล้ว ไม่มีการเปลี่ยนแปลง';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

/**
 * เพิ่มคอลัมน์ SCCD ในชีท Requests ที่ติดตั้งไปแล้ว (ไม่กระทบข้อมูลเดิม)
 * เลือกฟังก์ชันนี้แล้วกด Run ครั้งเดียว — รันซ้ำได้ ไม่เพิ่มคอลัมน์ซ้ำ
 */
function migrateAddSccdColumn() {
  const sh = getSheet_(SHEET_REQUESTS);
  if (!sh) throw new Error('ไม่พบชีท ' + SHEET_REQUESTS);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('SCCD') !== -1) {
    const msg = 'มีคอลัมน์ SCCD อยู่แล้ว ไม่มีการเปลี่ยนแปลง';
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
    return;
  }

  const ticketCol = headers.indexOf('TicketID');
  if (ticketCol === -1) throw new Error('ไม่พบคอลัมน์ TicketID');

  sh.insertColumnAfter(ticketCol + 1);          // แทรกคอลัมน์ว่างถัดจาก TicketID
  const newCol = ticketCol + 2;
  sh.getRange(1, newCol).setValue('SCCD')
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  sh.setColumnWidth(newCol, 110);
  SpreadsheetApp.flush();

  const msg = 'เพิ่มคอลัมน์ SCCD เรียบร้อย (คอลัมน์ ' + columnLetter_(newCol) + ') — งานเดิมจะเว้นว่างไว้ กรอกย้อนหลังในชีทได้';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

function columnLetter_(col) {
  let s = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

/** @return {number} จำนวนรายการที่เพิ่มเข้าไป (0 = มีข้อมูลอยู่แล้ว) */
function ensureLocationsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_LOCATIONS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOCATIONS);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(['LocationCode']);
  }
  sh.setFrozenRows(1);
  sh.getRange('A1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  if (sh.getLastRow() > 1) return 0; // มีข้อมูลแล้ว ไม่แตะต้อง

  sh.getRange(2, 1, LOCATION_LIST.length, 1)
    .setValues(LOCATION_LIST.map(code => [code]));
  sh.autoResizeColumn(1);
  SpreadsheetApp.flush();
  return LOCATION_LIST.length;
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
  return { categories: categories, locations: getLocations_(), priorities: PRIORITY_LIST };
}

/** รายชื่อสถานที่จากชีท Locations (ถ้ายังไม่มีชีท ใช้ค่าตั้งต้นในโค้ด) */
function getLocations_() {
  const sh = getSheet_(SHEET_LOCATIONS);
  if (!sh || sh.getLastRow() < 2) return LOCATION_LIST.slice();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(String);
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

  // สถานที่ต้องเป็นรหัสที่มีอยู่ในชีท Locations เท่านั้น (กันข้อมูลเพี้ยนตอนทำรายงาน)
  const locations = getLocations_();
  if (locations.indexOf(String(data.location).trim()) === -1) {
    throw new Error('ไม่พบสถานที่ "' + data.location + '" ในระบบ กรุณาเลือกจากรายการ');
  }

  const sh = getSheet_(SHEET_REQUESTS);
  const ticketId = generateTicketId_();
  const now = new Date();

  let photoUrl = '';
  if (data.photoBase64) {
    photoUrl = uploadPhoto_(data.photoBase64, data.photoMimeType, data.photoName, ticketId);
  }

  // เขียนตามชื่อหัวคอลัมน์จริงในชีท เพื่อไม่ให้พลาดถ้าลำดับคอลัมน์ต่างจากค่าตั้งต้น
  const values = {
    TicketID: ticketId,
    SCCD: data.sccd || '',
    Timestamp: now,
    ReporterType: data.reporterType || 'บุคคลทั่วไป',
    ReporterName: data.reporterName,
    Contact: data.contact || '',
    Department: data.department || '',
    Location: data.location,
    Category: data.category,
    Description: data.description,
    PhotoURL: photoUrl,
    Status: STATUS_LIST[0],
    Priority: data.priority || PRIORITY_LIST[1],
    AssignedTo: '',
    UpdatedAt: now,
    ClosedAt: '',
    Notes: ''
  };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => values[h] !== undefined ? values[h] : ''));

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
    (data.sccd ? `SCCD/ITSM: ${data.sccd}\n` : '') +
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

// รับค่า Active ได้ทั้งแบบ checkbox (true) และแบบพิมพ์ข้อความ (TRUE / ใช่ / 1)
function isActive_(v) {
  if (v === true) return true;
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'ใช่';
}

function getTechnicians(pin) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_TECHNICIANS);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() && isActive_(r[2]))
    .map(r => String(r[0]).trim());
}

// ============ ADMIN API: จัดการช่าง ============
/** คืนรายชื่อช่างทั้งหมด (รวมที่ปิดใช้งาน) พร้อมเลขแถวในชีท */
function getTechniciansAll(pin) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_TECHNICIANS);
  const values = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0]).trim();
    if (!name) continue;
    list.push({
      row: i + 1,
      name: name,
      phone: String(values[i][1] || '').trim(),
      active: isActive_(values[i][2])
    });
  }
  return list;
}

function addTechnician(pin, name, phone) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณากรอกชื่อช่าง');

  const sh = getSheet_(SHEET_TECHNICIANS);
  const existing = sh.getDataRange().getValues().slice(1)
    .map(r => String(r[0]).trim().toLowerCase());
  if (existing.indexOf(name.toLowerCase()) !== -1) {
    throw new Error('มีชื่อช่างนี้อยู่แล้ว');
  }

  sh.appendRow([name, String(phone || '').trim(), true]);
  sh.getRange(sh.getLastRow(), 3).insertCheckboxes();
  return { success: true };
}

/** updates = { name?, phone?, active? } — row คือเลขแถวจาก getTechniciansAll */
function updateTechnician(pin, row, updates) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_TECHNICIANS);
  row = Number(row);
  if (!(row >= 2) || row > sh.getLastRow()) throw new Error('ไม่พบข้อมูลช่างแถวนี้');

  const oldName = String(sh.getRange(row, 1).getValue()).trim();

  if (updates.name !== undefined) {
    const newName = String(updates.name).trim();
    if (!newName) throw new Error('กรุณากรอกชื่อช่าง');
    const dup = sh.getDataRange().getValues().slice(1).some((r, i) =>
      (i + 2) !== row && String(r[0]).trim().toLowerCase() === newName.toLowerCase());
    if (dup) throw new Error('มีชื่อช่างนี้อยู่แล้ว');
    sh.getRange(row, 1).setValue(newName);
    if (newName !== oldName) renameAssignedTo_(oldName, newName);
  }
  if (updates.phone !== undefined) {
    sh.getRange(row, 2).setValue(String(updates.phone).trim());
  }
  if (updates.active !== undefined) {
    sh.getRange(row, 3).setValue(updates.active === true || updates.active === 'true');
  }
  return { success: true };
}

function deleteTechnician(pin, row) {
  if (!checkAdminPin(pin)) throw new Error('PIN ไม่ถูกต้อง');
  const sh = getSheet_(SHEET_TECHNICIANS);
  row = Number(row);
  if (!(row >= 2) || row > sh.getLastRow()) throw new Error('ไม่พบข้อมูลช่างแถวนี้');
  sh.deleteRow(row);
  return { success: true };
}

/** เปลี่ยนชื่อช่างในงานที่มอบหมายไว้แล้ว ให้ตรงกับชื่อใหม่ */
function renameAssignedTo_(oldName, newName) {
  if (!oldName) return;
  const sh = getSheet_(SHEET_REQUESTS);
  const data = sh.getDataRange().getValues();
  const col = data[0].indexOf('AssignedTo');
  if (col === -1) return;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]).trim() === oldName) {
      sh.getRange(i + 1, col + 1).setValue(newName);
    }
  }
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

