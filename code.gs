// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

const DEFAULT_ADMIN_PIN = '162823';

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

  if (action === 'verifyAdminPin') {
    result = verifyAdminPin(data.adminPin) || { success: true, message: 'PIN ถูกต้อง' };
  } else if (action === 'listUsers') {
    result = verifyAdminPin(data.adminPin) || listUsers();
  } else if (action === 'deleteUser') {
    result = verifyAdminPin(data.adminPin) || deleteUser(data.name);
  } else if (action === 'registerUser') {
    result = verifyAdminPin(data.adminPin) || registerUser(data.name, data.faceDescriptor);
  } else if (action === 'logAttendance') {
    result = logAttendance(data.name, data.lat, data.lng, data.attendanceType);
  } else if (action === 'getAttendanceStatus') {
    result = getAttendanceStatus(data.name);
  } else if (action === 'submitAttendanceNote') {
    result = submitAttendanceNote(data.name, data.date, data.endDate, data.noteType, data.note);
  } else if (action === 'saveConfig') {
    result = verifyAdminPin(data.adminPin) || saveConfig(data.lat, data.lng, data.radius, data.workDays, data.specialHolidays);
  } else if (action === 'getAttendanceReport') {
    result = verifyAdminPin(data.adminPin) || getAttendanceReport(data.fromDate, data.toDate, data.employeeName);
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, faceDescriptor) {
  name = normalizeEmployeeName(name);
  if (!name) return { error: 'กรุณากรอกชื่อพนักงาน' };
  if (!Array.isArray(faceDescriptor) || faceDescriptor.length !== 128) {
    return { error: 'ข้อมูลใบหน้าไม่ถูกต้อง' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users');

  const existingRow = findUserRowByName(sheet, name);
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, 3).setValues([[name, JSON.stringify(faceDescriptor), new Date()]]);
    removeDuplicateUserRows(sheet, name, existingRow);
    return { success: true, message: 'อัปเดตข้อมูลใบหน้าของ ' + name + ' เรียบร้อย' };
  }

  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]);
  return { success: true, message: 'บันทึกข้อมูลหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const usersByName = {};
  for (let i = 1; i < data.length; i++) {
    const name = normalizeEmployeeName(data[i][0]);
    const jsonStr = data[i][1];
    if (name && jsonStr) {
      try {
        usersByName[name] = { label: name, descriptor: JSON.parse(jsonStr) };
      } catch (e) {}
    }
  }
  return Object.keys(usersByName).sort().map(function(name) { return usersByName[name]; });
}

function listUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  const usersByName = {};
  data.forEach(function(row, index) {
    const name = normalizeEmployeeName(row[0]);
    if (!name) return;
    usersByName[name] = {
      name: name,
      registeredAt: row[2] || '',
      rowNumber: index + 2
    };
  });
  return Object.keys(usersByName).sort().map(function(name) { return usersByName[name]; });
}

function deleteUser(name) {
  name = normalizeEmployeeName(name);
  if (!name) return { error: 'ไม่พบชื่อพนักงานที่ต้องการลบ' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() < 2) return { error: 'ยังไม่มีข้อมูลพนักงาน' };

  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  let deletedCount = 0;
  for (let i = names.length - 1; i >= 0; i--) {
    if (normalizeEmployeeName(names[i][0]) === name) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }
  if (deletedCount > 0) return { success: true, message: 'ลบพนักงานเรียบร้อย: ' + name };
  return { error: 'ไม่พบพนักงาน: ' + name };
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, lat, lng, attendanceType) {
  name = normalizeEmployeeName(name);
  const latitude = parseFloat(lat);
  const longitude = parseFloat(lng);
  const type = String(attendanceType || 'in').toLowerCase() === 'out' ? 'out' : 'in';
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
  const sheet = ensureAttendanceSheet(ss);

  const now = new Date();
  const mapLink = `https://www.google.com/maps?q=${latitude},${longitude}`;
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  let rowNumber = findAttendanceRow(sheet, name, dateStr);
  const existing = rowNumber ? getAttendanceRowStatus(sheet, rowNumber) : { timeIn: '', timeOut: '' };

  if (type === 'in' && existing.timeIn) {
    if (existing.timeOut) return { error: 'วันนี้บันทึกเวลาเข้าและออกงานครบแล้ว' };
    return { error: 'วันนี้มีเวลาเข้างานแล้ว กรุณากดยืนยันออกงาน' };
  }

  if (type === 'out') {
    if (!rowNumber || !existing.timeIn) return { error: 'ยังไม่มีเวลาเข้างานของวันนี้ กรุณาเข้างานก่อน' };
    if (existing.timeOut) return { error: 'วันนี้บันทึกเวลาออกงานแล้ว' };
  }

  if (!rowNumber) {
    sheet.insertRowBefore(2);
    rowNumber = 2;
    sheet.getRange(rowNumber, 1, 1, 11).setValues([[
      name, "'" + dateStr, '', '', '', '', '', '', '', '', ''
    ]]);
  }

  if (type === 'out') {
    sheet.getRange(rowNumber, 4).setValue(timeStr);
    sheet.getRange(rowNumber, 7).setValue(latitude);
    sheet.getRange(rowNumber, 8).setValue(longitude);
    sheet.getRange(rowNumber, 10).setValue(mapLink);
    return { success: true, message: 'บันทึกเวลาออกงานสำเร็จ' };
  }

  sheet.getRange(rowNumber, 3).setValue(timeStr);
  sheet.getRange(rowNumber, 5).setValue(latitude);
  sheet.getRange(rowNumber, 6).setValue(longitude);
  sheet.getRange(rowNumber, 9).setValue(mapLink);
  return { success: true, message: 'บันทึกเวลาเข้างานสำเร็จ' };
}

function submitAttendanceNote(name, startDateValue, endDateValue, noteType, note) {
  name = normalizeEmployeeName(name);
  const startDate = parseIsoDate(startDateValue);
  const endDate = parseIsoDate(endDateValue || startDateValue);
  const type = String(noteType || '').trim();
  const detail = String(note || '').replace(/\s+/g, ' ').trim();

  if (!name) return { error: 'กรุณาเลือกหรือกรอกชื่อพนักงาน' };
  if (!startDate || !endDate) return { error: 'กรุณาเลือกวันที่ให้ถูกต้อง' };
  if (startDate.getTime() > endDate.getTime()) return { error: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น' };
  if (!type) return { error: 'กรุณาเลือกประเภทการแจ้ง' };
  if (!detail) return { error: 'กรุณากรอกรายละเอียด' };

  const dates = enumerateDates(startDate, endDate);
  if (dates.length > 5) return { error: 'แจ้งได้ไม่เกิน 5 วันต่อครั้ง' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAttendanceSheet(ss);
  const submittedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd/M/yyyy HH:mm:ss');
  const rangeText = dates.length > 1
    ? ' ช่วง ' + formatSheetDateFromDate(startDate) + ' - ' + formatSheetDateFromDate(endDate)
    : '';
  const noteText = '[' + type + ']' + rangeText + ' ' + detail + ' (แจ้งเมื่อ ' + submittedAt + ')';

  dates.forEach(function(dateObj) {
    const dateStr = formatSheetDateFromDate(dateObj);
    let rowNumber = findAttendanceRow(sheet, name, dateStr);

    if (!rowNumber) {
      sheet.insertRowBefore(2);
      rowNumber = 2;
      sheet.getRange(rowNumber, 1, 1, 11).setValues([[
        name, "'" + dateStr, '', '', '', '', '', '', '', '', ''
      ]]);
    }

    sheet.getRange(rowNumber, 11).setValue(noteText);
  });

  return {
    success: true,
    message: dates.length > 1 ? 'บันทึกหมายเหตุ ' + dates.length + ' วันเรียบร้อย' : 'บันทึกหมายเหตุเรียบร้อย',
    days: dates.length
  };
}

function getAttendanceStatus(name) {
  name = normalizeEmployeeName(name);
  if (!name) return { error: 'ไม่พบชื่อพนักงาน' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAttendanceSheet(ss);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const rowNumber = findAttendanceRow(sheet, name, dateStr);
  const status = rowNumber ? getAttendanceRowStatus(sheet, rowNumber) : { timeIn: '', timeOut: '' };
  let nextAttendanceType = 'in';
  let message = 'ยังไม่ได้เข้างานวันนี้';

  if (status.timeIn && !status.timeOut) {
    nextAttendanceType = 'out';
    message = 'วันนี้เข้างานแล้ว กรุณายืนยันออกงาน';
  } else if (status.timeIn && status.timeOut) {
    nextAttendanceType = 'done';
    message = 'วันนี้บันทึกเวลาเข้าและออกงานครบแล้ว';
  }

  return {
    success: true,
    name: name,
    date: dateStr,
    timeIn: status.timeIn,
    timeOut: status.timeOut,
    nextAttendanceType: nextAttendanceType,
    message: message
  };
}

function ensureAttendanceSheet(ss) {
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) sheet = ss.insertSheet('Attendance');
  const headers = [
    'Name', 'Date', 'Time In', 'Time Out',
    'Latitude In', 'Longitude In', 'Latitude Out', 'Longitude Out',
    'Map In', 'Map Out', 'Note'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function findAttendanceRow(sheet, name, dateStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const targetName = normalizeEmployeeName(name);
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeEmployeeName(values[i][0]) === targetName && String(values[i][1]).replace(/^'/, '').trim() === dateStr) {
      return i + 2;
    }
  }
  return 0;
}

function getAttendanceRowStatus(sheet, rowNumber) {
  const values = sheet.getRange(rowNumber, 3, 1, 2).getDisplayValues()[0];
  return {
    timeIn: String(values[0] || '').trim(),
    timeOut: String(values[1] || '').trim()
  };
}

// --- ส่วนจัดการ Config (GPS + Calendar) ---
function saveConfig(lat, lng, radius, workDays, specialHolidays) {
  lat = parseFloat(lat);
  lng = parseFloat(lng);
  radius = parseFloat(radius);
  const normalizedWorkDays = normalizeWorkDays(workDays);
  const normalizedSpecialHolidays = normalizeSpecialHolidays(specialHolidays);
  if (!isFinite(lat) || lat < -90 || lat > 90) return { error: 'Latitude ไม่ถูกต้อง' };
  if (!isFinite(lng) || lng < -180 || lng > 180) return { error: 'Longitude ไม่ถูกต้อง' };
  if (!isFinite(radius) || radius < 0) return { error: 'รัศมีไม่ถูกต้อง' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) sheet = ss.insertSheet('Config');

  const rows = [
    ['Parameter', 'Value'],
    ['Target Latitude', lat],
    ['Target Longitude', lng],
    ['Allowed Radius (KM)', radius],
    ['Work Days', normalizedWorkDays.join(',')],
    ['Special Holidays', normalizedSpecialHolidays.join(',')]
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 260);

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.5, workDays: [1, 2, 3, 4, 5, 6], specialHolidays: [] };

  if (sheet) {
    const values = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 0; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      if (key && key !== 'Parameter') map[key] = values[i][1];
    }
    const latVal = map['Target Latitude'];
    const lngVal = map['Target Longitude'];
    const radiusVal = map['Allowed Radius (KM)'];

    if (latVal !== undefined && latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== undefined && lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== undefined && radiusVal !== '') config.radius = parseFloat(radiusVal);
    config.workDays = normalizeWorkDays(map['Work Days']);
    config.specialHolidays = normalizeSpecialHolidays(map['Special Holidays']);
  }

  return config;
}

function getAttendanceReport(fromDate, toDate, employeeName) {
  const startDate = parseIsoDate(fromDate);
  const endDate = parseIsoDate(toDate);
  if (!startDate || !endDate) return { error: 'กรุณาเลือกวันที่เริ่มต้นและสิ้นสุด' };
  if (startDate.getTime() > endDate.getTime()) return { error: 'วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedEmployee = normalizeEmployeeName(employeeName);
  let employees = listUsers().map(function(user) { return user.name; }).sort();
  if (selectedEmployee) employees = employees.filter(function(name) { return normalizeEmployeeName(name) === selectedEmployee; });
  const config = getConfig();
  const workDaySet = {};
  config.workDays.forEach(function(day) { workDaySet[day] = true; });
  const holidaySet = {};
  config.specialHolidays.forEach(function(day) { holidaySet[day] = true; });
  const attendanceMap = buildAttendanceMap(ss);
  const dates = enumerateDates(startDate, endDate);
  const daily = [];
  const summaryMap = {};

  employees.forEach(function(name) {
    summaryMap[name] = {
      name: name,
      workedDays: 0,
      absentDays: 0,
      offDays: 0,
      specialHolidayDays: 0,
      incompleteDays: 0,
      totalHours: 0,
      absentDates: [],
      incompleteDates: []
    };
  });

  dates.forEach(function(dateObj) {
    const dateKey = formatDateKey(dateObj);
    const isSpecialHoliday = !!holidaySet[dateKey];
    const isWorkday = !!workDaySet[dateObj.getDay()] && !isSpecialHoliday;

    employees.forEach(function(name) {
      const record = attendanceMap[normalizeEmployeeName(name) + '|' + dateKey] || {};
      const timeIn = record.timeIn || '';
      const timeOut = record.timeOut || '';
      const note = record.note || '';
      const hours = calculateHours(timeIn, timeOut);
      let status;
      const summary = summaryMap[name];

      if (timeIn && timeOut) {
        status = isSpecialHoliday || !isWorkday ? 'มาทำงานวันหยุด' : 'ออกงานแล้ว';
        summary.workedDays++;
        summary.totalHours = round2(summary.totalHours + hours);
      } else if (timeIn && !timeOut) {
        status = 'ลืมออกงาน';
        summary.incompleteDays++;
        summary.incompleteDates.push(dateKey);
      } else if (!timeIn && timeOut) {
        status = 'มีเวลาออก ไม่มีเวลาเข้า';
        summary.incompleteDays++;
        summary.incompleteDates.push(dateKey);
      } else if (note.indexOf('[ลาหยุด]') === 0) {
        status = 'ลาหยุด';
        summary.offDays++;
      } else if (note.indexOf('[มาสาย]') === 0) {
        status = 'มาสาย/รอเข้างาน';
        summary.incompleteDays++;
        summary.incompleteDates.push(dateKey);
      } else if (isSpecialHoliday) {
        status = 'วันหยุดพิเศษ';
        summary.specialHolidayDays++;
      } else if (!isWorkday) {
        status = 'วันหยุดประจำ';
        summary.offDays++;
      } else {
        status = 'หยุด/ไม่ลงเวลา';
        summary.absentDays++;
        summary.absentDates.push(dateKey);
      }

      daily.push({
        date: dateKey,
        name: name,
        timeIn: timeIn,
        timeOut: timeOut,
        hours: hours,
        status: status,
        note: note
      });
    });
  });

  const summary = Object.keys(summaryMap).sort().map(function(name) {
    summaryMap[name].totalHours = round2(summaryMap[name].totalHours);
    return summaryMap[name];
  });

  return {
    success: true,
    fromDate: formatDateKey(startDate),
    toDate: formatDateKey(endDate),
    workDays: config.workDays,
    specialHolidays: config.specialHolidays,
    selectedEmployee: selectedEmployee,
    employeeCount: employees.length,
    summary: summary,
    daily: daily
  };
}

function buildAttendanceMap(ss) {
  const map = {};
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet || sheet.getLastRow() < 2) return map;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  const headerIndex = {};
  headers.forEach(function(header, index) { headerIndex[header] = index; });
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();

  rows.forEach(function(row) {
    const name = normalizeEmployeeName(row[headerIndex['Name']]);
    const dateKey = parseDateTextToKey(row[headerIndex['Date']]);
    if (!name || !dateKey) return;

    const oldTime = row[headerIndex['Time']] || '';
    const timeIn = row[headerIndex['Time In']] || oldTime || '';
    const timeOut = row[headerIndex['Time Out']] || '';
    const note = row[headerIndex['Note']] || '';
    map[normalizeEmployeeName(name) + '|' + dateKey] = {
      timeIn: String(timeIn || '').trim(),
      timeOut: String(timeOut || '').trim(),
      note: String(note || '').trim()
    };
  });
  return map;
}

function formatSheetDateFromDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
}

function normalizeEmployeeName(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function findUserRowByName(sheet, name) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const targetName = normalizeEmployeeName(name);
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < names.length; i++) {
    if (normalizeEmployeeName(names[i][0]) === targetName) return i + 2;
  }
  return 0;
}

function removeDuplicateUserRows(sheet, name, keepRow) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const targetName = normalizeEmployeeName(name);
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = names.length - 1; i >= 0; i--) {
    const rowNumber = i + 2;
    if (rowNumber !== keepRow && normalizeEmployeeName(names[i][0]) === targetName) {
      sheet.deleteRow(rowNumber);
    }
  }
}

function normalizeWorkDays(value) {
  let items = [];
  if (Array.isArray(value)) {
    items = value;
  } else if (value !== undefined && value !== null && String(value).trim() !== '') {
    items = String(value).split(',');
  }
  const days = items
    .map(function(day) { return parseInt(day, 10); })
    .filter(function(day) { return day >= 0 && day <= 6; });
  const unique = Array.from(new Set(days));
  return unique.length ? unique.sort(function(a, b) { return a - b; }) : [1, 2, 3, 4, 5, 6];
}

function normalizeSpecialHolidays(value) {
  let items = [];
  if (Array.isArray(value)) {
    items = value;
  } else if (value !== undefined && value !== null && String(value).trim() !== '') {
    items = String(value).split(/[\n,]/);
  }
  const dates = items
    .map(function(item) { return formatDateKey(parseIsoDate(String(item || '').trim())); })
    .filter(Boolean);
  return Array.from(new Set(dates)).sort();
}

function parseIsoDate(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  if (date.getFullYear() !== Number(match[1]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) {
    return null;
  }
  date.setHours(0, 0, 0, 0);
  return date;
}

function parseDateTextToKey(value) {
  const text = String(value || '').replace(/^'/, '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return '';
  const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  return formatDateKey(date);
}

function enumerateDates(startDate, endDate) {
  const dates = [];
  const cursor = new Date(startDate.getTime());
  while (cursor.getTime() <= endDate.getTime()) {
    dates.push(new Date(cursor.getTime()));
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

function formatDateKey(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return '';
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const day = String(dateObj.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

function calculateHours(timeIn, timeOut) {
  const start = parseTimeToMinutes(timeIn);
  let end = parseTimeToMinutes(timeOut);
  if (start === null || end === null) return 0;
  if (end < start) end += 24 * 60;
  return round2((end - start) / 60);
}

function parseTimeToMinutes(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > 23 || minute > 59) return null;
  return hour * 60 + minute;
}

function round2(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
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
