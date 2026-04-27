// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

const DEFAULT_ADMIN_PIN = '123456';

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if (action === 'getConfig') {
    result = getConfig();
  } else if (action === 'getKnownFaces') {
    result = getKnownFaces();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result;

  if (action === 'registerUser') {
    result = verifyAdminPin(data.adminPin) || registerUser(data.name, data.faceDescriptor);
  } else if (action === 'logAttendance') {
    result = logAttendance(data.name, data.lat, data.lng);
  } else if (action === 'saveConfig') {
    result = verifyAdminPin(data.adminPin) || saveConfig(data.lat, data.lng, data.radius);
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, faceDescriptor) {
  name = String(name || '').trim();
  if (!name) return { error: 'กรุณากรอกชื่อพนักงาน' };
  if (!Array.isArray(faceDescriptor) || faceDescriptor.length !== 128) {
    return { error: 'ข้อมูลใบหน้าไม่ถูกต้อง' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users');

  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]);
  return { success: true, message: 'บันทึกข้อมูลหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const jsonStr = data[i][1];
    if (name && jsonStr) {
      try {
        users.push({ label: name, descriptor: JSON.parse(jsonStr) });
      } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, lat, lng) {
  name = String(name || '').trim();
  const latitude = parseFloat(lat);
  const longitude = parseFloat(lng);
  if (!name) return { error: 'ไม่พบชื่อพนักงาน' };
  if (!isFinite(latitude) || !isFinite(longitude)) return { error: 'ไม่พบพิกัด GPS' };

  const config = getConfig();
  if (config.lat && config.lng && config.radius > 0) {
    const distanceKm = haversineKm(latitude, longitude, config.lat, config.lng);
    if (distanceKm > config.radius) {
      return {
        error: 'อยู่นอกพื้นที่เช็คอิน ระยะห่าง ' + distanceKm.toFixed(3) + ' km',
        distanceKm: distanceKm
      };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Time', 'Date', 'Latitude', 'Longitude', 'Google Map Link']);
  }

  const now = new Date();
  const mapLink = `https://www.google.com/maps?q=${latitude},${longitude}`;
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  sheet.appendRow([
    name,
    timeStr,
    "'" + dateStr,
    latitude,
    longitude,
    mapLink
  ]);
  return { success: true, message: 'บันทึกเวลาสำเร็จ' };
}

// --- ส่วนจัดการ Config (GPS) ---
function saveConfig(lat, lng, radius) {
  lat = parseFloat(lat);
  lng = parseFloat(lng);
  radius = parseFloat(radius);
  if (!isFinite(lat) || lat < -90 || lat > 90) return { error: 'Latitude ไม่ถูกต้อง' };
  if (!isFinite(lng) || lng < -180 || lng > 180) return { error: 'Longitude ไม่ถูกต้อง' };
  if (!isFinite(radius) || radius < 0) return { error: 'รัศมีไม่ถูกต้อง' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.setColumnWidth(1, 150);
  }

  sheet.getRange('B2').setValue(lat);
  sheet.getRange('B3').setValue(lng);
  sheet.getRange('B4').setValue(radius);

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.5 };

  if (sheet) {
    const latVal = sheet.getRange('B2').getValue();
    const lngVal = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();

    if (latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
  }

  return config;
}

function verifyAdminPin(pin) {
  const expected = getAdminPin();
  const input = String(pin || '').trim();
  if (!/^\d{6}$/.test(input)) return { error: 'กรุณาใส่ Admin PIN 6 หลัก' };
  if (input !== expected) return { error: 'Admin PIN ไม่ถูกต้อง' };
  return null;
}

function getAdminPin() {
  const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  return /^\d{6}$/.test(String(pin || '')) ? String(pin) : DEFAULT_ADMIN_PIN;
}

function setAdminPin(pin) {
  pin = String(pin || '').trim();
  if (!/^\d{6}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 6 หลัก');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  return { success: true, message: 'ตั้งค่า Admin PIN เรียบร้อย' };
}

function haversineKm(lat1, lon1, lat2, lon2) {
  const earthRadiusKm = 6371;
  const dLat = deg2rad(lat2 - lat1);
  const dLon = deg2rad(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(deg2rad(lat1)) * Math.cos(deg2rad(lat2)) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return earthRadiusKm * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function deg2rad(deg) {
  return deg * Math.PI / 180;
}
