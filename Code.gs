function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle("My Measurement")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== PROPERTIES / SHEET HELPERS ====================
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });
}

function appendSheetData(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' not found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => data[header] !== undefined ? data[header] : '');
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function findColumnIdx(headers, possibleNames) {
  function normalize(s) { return String(s).trim().replace(/\s+/g, ' ').toLowerCase(); }
  for (var i = 0; i < possibleNames.length; i++) {
    var search = normalize(possibleNames[i]);
    for (var j = 0; j < headers.length; j++) {
      if (normalize(headers[j]) === search) return j;
    }
  }
  return -1;
}

function ensureUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(["Type", "Value", "Parent ID", "UserID", "Password", "Mobile", "Status", "GMail", "Password Generated Date", "Password Expiry Date"]);
    sheet.appendRow(["Owner", "Admin", "", "admin", "admin123", "", "Active", "", "", ""]);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Type", "Value", "Parent ID", "UserID", "Password", "Mobile", "Status", "GMail", "Password Generated Date", "Password Expiry Date"]);
    sheet.appendRow(["Owner", "Admin", "", "admin", "admin123", "", "Active", "", "", ""]);
    return sheet;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 10);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var requiredColumns = ["Password", "Status"];
  var nextCol = headers.length + 1;
  requiredColumns.forEach(function(colName) {
    var found = false;
    for (var j = 0; j < headers.length; j++) { if (headers[j].toLowerCase() === colName.toLowerCase()) { found = true; break; } }
    if (!found) { sheet.getRange(1, nextCol).setValue(colName); headers.push(colName); nextCol++; }
  });
  return sheet;
}

// ==================== SAVED WORK ====================
function saveWork(sor, janpad, engineer, panchayat, workName, ts, as, head, remark, status) {
  let nextSrNo = 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Saved Work');
  if (!sheet) {
    sheet = ss.insertSheet('Saved Work');
    sheet.appendRow(["Sr no", "SOR", "Janpad", "Engineer", "Panchayat", "Work", "TS Authority", "AS Authority", "Head", "Remark", "Status", "Created Date", "Total Amount", "Estimate Data", "Valuation Data"]);
  }
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existingHeaders.indexOf('Valuation Data') === -1) { sheet.getRange(1, existingHeaders.length + 1).setValue('Valuation Data'); }
  if (existingHeaders.indexOf('Estimate Data') === -1) { sheet.getRange(1, existingHeaders.length + 1).setValue('Estimate Data'); }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const workNameIndex = headers.indexOf('Work');
  if (workNameIndex !== -1) {
    for (let i = 1; i < data.length; i++) { if (data[i][workNameIndex] === workName) throw new Error('A work with the same name already exists!'); }
  }
  if (data.length > 1) {
    const lastRow = data[data.length - 1];
    const srNoIndex = headers.indexOf('Sr no');
    if (srNoIndex !== -1) { const v = lastRow[srNoIndex]; if (v) nextSrNo = parseInt(v) + 1; }
  }
  return appendSheetData('Saved Work', {
    'Sr no': nextSrNo, 'SOR': sor, 'Janpad': janpad, 'Engineer': engineer, 'Panchayat': panchayat, 'Work': workName, 'TS Authority': ts, 'AS Authority': as, 'Head': head, 'Remark': remark, 'Status': status, 'Created Date': new Date().toLocaleDateString('en-IN'), 'Total Amount': 0, 'Estimate Data': '', 'Valuation Data': ''
  });
}

function getSavedWorks() { return getSheetData('Saved Work'); }

function updateWork(index, sor, janpad, engineer, panchayat, workName, ts, as, head, remark) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Saved Work');
  if (!sheet) throw new Error('Sheet "Saved Work" not found');
  const row = index + 2;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const upd = { 'SOR': sor, 'Janpad': janpad, 'Engineer': engineer, 'Panchayat': panchayat, 'Work': workName, 'TS Authority': ts, 'AS Authority': as, 'Head': head, 'Remark': remark };
  headers.forEach((header, colIndex) => { if (upd[header] !== undefined) sheet.getRange(row, colIndex + 1).setValue(upd[header]); });
  return true;
}

function saveEstimateData(index, estimateDataJson, status, totalAmount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Saved Work');
  if (!sheet) throw new Error('Sheet "Saved Work" not found');
  const row = index + 2;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let targetColName = 'Estimate Data';
  if (status === 'Valuation') { targetColName = 'Valuation Data'; }
  let edi = headers.indexOf(targetColName);
  if (edi === -1) { edi = headers.length; sheet.getRange(1, edi + 1).setValue(targetColName); }
  sheet.getRange(row, edi + 1).setValue(estimateDataJson);
  const si = headers.indexOf('Status');
  if (si !== -1) sheet.getRange(row, si + 1).setValue(status);
  const ai = headers.indexOf('Total Amount');
  if (ai !== -1 && totalAmount !== undefined && totalAmount !== null) sheet.getRange(row, ai + 1).setValue(totalAmount);
  return true;
}

function deleteWork(index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Saved Work');
  if (!sheet) throw new Error('Sheet "Saved Work" not found');
  sheet.deleteRow(index + 2);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const srIdx = headers.indexOf('Sr no');
  if (srIdx !== -1) { for (let i = 1; i < data.length; i++) { sheet.getRange(i + 1, srIdx + 1).setValue(i); } }
  return true;
}

// ==================== SPECIFICATIONS ====================
function getSpecifications(sorType) {
  const sheetName = 'RES'; // MGNREGA removed
  const data = getSheetData(sheetName);
  return data.map(row => ({ 
    code: row['Code'], spec: row['Specification'], rate1: row['Work Rate'], rate2: row['Labour Rate'], unit: row['Unit'],
    m40: parseFloat(row['Metal 40mm (m³)']) || 0, m20: parseFloat(row['Metal 20mm (m³)']) || 0, m10: parseFloat(row['Metal 10mm (m³)']) || 0,
    cement: parseFloat(row['Cement (Bag)']) || 0, sand: parseFloat(row['Sand (m³)']) || 0, brick: parseFloat(row['Brick (Nos)']) || 0
  }));
}

function addSpecification(sorType, code, spec, rate1, rate2, unit, materials) {
  const sheetName = 'RES';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' not found');
  const data = sheet.getDataRange().getValues();
  const ci = data[0].indexOf('Code');
  if (ci !== -1) { for (let i = 1; i < data.length; i++) { if (data[i][ci] === code) throw new Error('A specification with the same code already exists!'); } }
  var rowData = {
    'Sr No': '', 'Code': code, 'Specification': spec, 'Work Rate': rate1, 'Labour Rate': rate2, 'Unit': unit,
    'Metal 40mm (m³)': materials.m40 || 0, 'Metal 20mm (m³)': materials.m20 || 0, 'Metal 10mm (m³)': materials.m10 || 0,
    'Cement (Bag)': materials.cement || 0, 'Sand (m³)': materials.sand || 0, 'Brick (Nos)': materials.brick || 0
  };
  return appendSheetData(sheetName, rowData);
}

function updateSpecification(sorType, rowIndex, code, spec, rate1, rate2, unit, materials) {
  const sheetName = 'RES';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' not found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updates = { 
    'Code': code, 'Specification': spec, 'Work Rate': rate1, 'Labour Rate': rate2, 'Unit': unit,
    'Metal 40mm (m³)': materials.m40 || 0, 'Metal 20mm (m³)': materials.m20 || 0, 'Metal 10mm (m³)': materials.m10 || 0,
    'Cement (Bag)': materials.cement || 0, 'Sand (m³)': materials.sand || 0, 'Brick (Nos)': materials.brick || 0
  };
  headers.forEach(function(header, colIndex) { if (updates[header] !== undefined) { sheet.getRange(rowIndex + 2, colIndex + 1).setValue(updates[header]); } });
  return true;
}

function deleteSpecification(sorType, rowIndex) {
  const sheetName = 'RES';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' not found');
  sheet.deleteRow(rowIndex + 2);
  return true;
}

// ==================== TEMPLATES ====================
function saveTemplate(templateType, name, description, sorType, templateData, userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Templates');
  if (!sheet) { sheet = ss.insertSheet('Templates'); sheet.appendRow(["Template Type", "Template Name", "Description", "SOR Type", "Template Data", "Created Date", "Created By"]); }
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var createdByIdx = headers.indexOf('Created By');
  if (createdByIdx === -1) { sheet.getRange(1, lastCol + 1).setValue('Created By'); }
  var existing = getSheetData('Templates');
  for (var i = 0; i < existing.length; i++) { if (existing[i]['Template Name'] === name && existing[i]['Template Type'] === templateType) throw new Error('A template with this name already exists!'); }
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");
  return appendSheetData('Templates', { 'Template Type': templateType || 'My', 'Template Name': name, 'Description': description || '', 'SOR Type': sorType, 'Template Data': JSON.stringify(templateData), 'Created Date': dateStr, 'Created By': userId || '' });
}

function getTemplates() {
  var data = getSheetData('Templates');
  return data.map(function(row) {
    var parsed = {}; try { parsed = JSON.parse(row['Template Data']); } catch(e) { parsed = { sections: [] }; }
    var cd = row['Created Date'];
    if (cd != null && cd != '' && typeof cd !== 'string') { try { cd = Utilities.formatDate(cd, Session.getScriptTimeZone(), "dd-MM-yyyy"); } catch(e) { cd = String(cd); } }
    return { templateType: row['Template Type'] || 'My', name: row['Template Name'] || '', description: row['Description'] || '', sorType: row['SOR Type'] || '', data: parsed, createdDate: cd || '', createdBy: row['Created By'] || '' };
  });
}

function updateTemplate(index, templateType, name, description, sorType, templateData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');
  if (!sheet) throw new Error('Sheet "Templates" not found');
  var row = index + 2;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var tti = headers.indexOf('Template Type'), ni = headers.indexOf('Template Name'), di = headers.indexOf('Description'), si = headers.indexOf('SOR Type'), tdi = headers.indexOf('Template Data'), cdi = headers.indexOf('Created Date');
  if (tti !== -1) sheet.getRange(row, tti + 1).setValue(templateType || 'My');
  if (ni !== -1) sheet.getRange(row, ni + 1).setValue(name);
  if (di !== -1) sheet.getRange(row, di + 1).setValue(description || '');
  if (si !== -1) sheet.getRange(row, si + 1).setValue(sorType);
  if (tdi !== -1) sheet.getRange(row, tdi + 1).setValue(JSON.stringify(templateData));
  if (cdi !== -1) sheet.getRange(row, cdi + 1).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy"));
  return true;
}

function deleteTemplate(index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');
  if (!sheet) throw new Error('Sheet "Templates" not found');
  sheet.deleteRow(index + 2);
  return true;
}

// ==================== DATABASE (Dropdowns) ====================
function getDataByType(type, sheetName) {
  sheetName = sheetName || "Database";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.filter(row => row[0] && String(row[0]).toLowerCase() === String(type).toLowerCase()).map(row => ({ value: row[1], parent: row[2], id: row[3] }));
}

function addDataToDatabase(type, value, parent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Database") || ss.insertSheet("Database");
  if (sheet.getLastRow() === 0) sheet.appendRow(["Type", "Value", "Parent", "ID"]);
  const id = Utilities.getUuid();
  sheet.appendRow([type, value, parent, id]);
  return id;
}

function addTsAuthority(a) { return addDataToDatabase("TS Authority", a, ""); }
function addAsAuthority(a) { return addDataToDatabase("AS Authority", a, ""); }
function addHead(h) { return addDataToDatabase("Head", h, ""); }
function addProvision(p) { return addDataToDatabase("Provision", p, ""); }

function getTsAuthorities() { return getDataByType("TS Authority").map(i => i.value); }
function getAsAuthorities() { return getDataByType("AS Authority").map(i => i.value); }
function getHeads() { return getDataByType("Head").map(i => i.value); }
function getProvisions() { try { const ss = SpreadsheetApp.getActiveSpreadsheet(); const s = ss.getSheetByName("Database"); if (!s || s.getLastRow() <= 1) return []; const d = s.getRange(2, 1, s.getLastRow() - 1, 4).getValues(); return d.filter(r => r[0] === "Provision").map(r => r[1]); } catch(e) { return []; } }

// ==================== USERS SHEET READERS ====================
function formatSheetDate(dateVal) {
  if (!dateVal) return "";
  if (typeof dateVal === 'string') {
    var s = dateVal.trim();
    if (s === '') return '';
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
    if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)) return s.replace(/-/g, '/');
    var isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) return isoMatch[3] + '/' + isoMatch[2] + '/' + isoMatch[1];
    try { var d = new Date(s); if (!isNaN(d.getTime())) { return (d.getDate() < 10 ? '0' : '') + d.getDate() + '/' + (d.getMonth() + 1 < 10 ? '0' : '') + (d.getMonth() + 1) + '/' + d.getFullYear(); } } catch(e) {}
    return s;
  }
  if (dateVal instanceof Date) {
    try { return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd/MM/yyyy"); } catch(e) { return (dateVal.getDate() < 10 ? '0' : '') + dateVal.getDate() + '/' + (dateVal.getMonth() + 1 < 10 ? '0' : '') + (dateVal.getMonth() + 1) + '/' + dateVal.getFullYear(); }
  }
  if (typeof dateVal === 'number') {
    try { var d = new Date((dateVal - 25569) * 86400 * 1000); return (d.getDate() < 10 ? '0' : '') + d.getDate() + '/' + (d.getMonth() + 1 < 10 ? '0' : '') + (d.getMonth() + 1) + '/' + d.getFullYear(); } catch(e) { return ''; }
  }
  return String(dateVal);
}

function readUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet || sheet.getLastRow() <= 1) return { data: [], headers: [], typeIdx: -1, valueIdx: -1, parentIdIdx: -1, userIdIdx: -1, passwordIdx: -1, statusIdx: -1, mobileIdx: -1, gmailIdx: -1 };
  const rawData = sheet.getDataRange().getValues();
  var headers = rawData[0].map(function(h) { return String(h).trim(); });
  const data = rawData.slice(1);
  return {
    data: data, headers: headers,
    typeIdx: findColumnIdx(headers, ["Type"]), valueIdx: findColumnIdx(headers, ["Value", "Name"]),
    parentIdIdx: findColumnIdx(headers, ["Parent ID", "ParentID", "Parent", "parentid"]),
    userIdIdx: findColumnIdx(headers, ["UserID", "User ID", "userid", "user_id"]),
    passwordIdx: findColumnIdx(headers, ["Password", "password", "pwd"]),
    statusIdx: findColumnIdx(headers, ["Status", "status"]),
    mobileIdx: findColumnIdx(headers, ["Mobile", "mobile", "Phone", "phone", "Contact"]),
    gmailIdx: findColumnIdx(headers, ["GMail", "Gmail", "gmail", "Email", "email", "E-Mail"]),
    pwdGenDateIdx: findColumnIdx(headers, ["Password Generated Date", "Password  Generated Date", "Pwd Gen Date", "PwdGenDate", "PasswordGeneratedDate"]),
    pwdExpDateIdx: findColumnIdx(headers, ["Password Expiry Date", "Password  Expiry Date", "Pwd Exp Date", "PwdExpDate", "PasswordExpiryDate", "Expiry Date"])
  };
}

// ==================== USER MANAGEMENT (SECURITY) ====================
function getUsersForAdmin() {
  var u = readUsersSheet();
  if (u.data.length === 0) return [];
  var users = [];
  for (var i = 0; i < u.data.length; i++) {
    users.push({
      index: i, type: u.typeIdx !== -1 ? String(u.data[i][u.typeIdx]).trim() : "", value: u.valueIdx !== -1 ? String(u.data[i][u.valueIdx]).trim() : "",
      parentId: u.parentIdIdx !== -1 ? String(u.data[i][u.parentIdIdx]).trim() : "", userId: u.userIdIdx !== -1 ? String(u.data[i][u.userIdIdx]).trim() : "",
      password: u.passwordIdx !== -1 ? String(u.data[i][u.passwordIdx]).trim() : "", mobile: u.mobileIdx !== -1 ? String(u.data[i][u.mobileIdx]).trim() : "",
      gmail: u.gmailIdx !== -1 ? String(u.data[i][u.gmailIdx]).trim() : "", status: u.statusIdx !== -1 ? String(u.data[i][u.statusIdx]).trim() : "Active"
    });
  }
  return users;
}

function saveUserData(data) {
  ensureUsersSheet();
  var u = readUsersSheet();
  if (data.userId) {
    for (var i = 0; i < u.data.length; i++) { if (String(u.data[i][u.userIdIdx]).trim() === data.userId && i !== data.index) { throw new Error("User ID already exists! Please choose a different ID."); } }
  }
  if (data.index !== undefined && data.index !== null && data.index !== "") {
    var row = parseInt(data.index) + 2;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (u.typeIdx !== -1) sheet.getRange(row, u.typeIdx + 1).setValue(data.type);
    if (u.valueIdx !== -1) sheet.getRange(row, u.valueIdx + 1).setValue(data.value);
    if (u.parentIdIdx !== -1) sheet.getRange(row, u.parentIdIdx + 1).setValue(data.parentId || '');
    if (u.userIdIdx !== -1) sheet.getRange(row, u.userIdIdx + 1).setValue(data.userId);
    if (u.mobileIdx !== -1) sheet.getRange(row, u.mobileIdx + 1).setValue(data.mobile || '');
    if (u.gmailIdx !== -1) sheet.getRange(row, u.gmailIdx + 1).setValue(data.gmail || '');
    if (u.statusIdx !== -1) sheet.getRange(row, u.statusIdx + 1).setValue(data.status || 'Active');
    if (u.passwordIdx !== -1 && data.password) { sheet.getRange(row, u.passwordIdx + 1).setValue(data.password); }
  } else {
    var genDate = new Date(); var expDate = new Date(genDate); expDate.setMonth(expDate.getMonth() + 3);
    var genDateStr = Utilities.formatDate(genDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var expDateStr = Utilities.formatDate(expDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    appendSheetData('Users', { 'Type': data.type, 'Value': data.value, 'Parent ID': data.parentId || '', 'UserID': data.userId, 'Password': data.password || '123456', 'Mobile': data.mobile || '', 'Status': data.status || 'Active', 'GMail': data.gmail || '', 'Password Generated Date': genDateStr, 'Password Expiry Date': expDateStr });
  }
  return true;
}

function deleteUser(index) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error("Users sheet not found");
  var u = readUsersSheet();
  if (u.data[index] && String(u.data[index][u.userIdIdx]).trim() === 'admin') { throw new Error("Cannot delete the main Admin account!"); }
  sheet.deleteRow(parseInt(index) + 2);
  return true;
}

function getJanpads() { 
  try {
    var u = readUsersSheet();
    if (u.typeIdx === -1 || u.valueIdx === -1) return [];
    var uniqueJanpads = [...new Set(u.data.filter(r => String(r[u.typeIdx]).trim() === "Janpad").map(r => String(r[u.valueIdx]).trim()).filter(v => v))];
    return uniqueJanpads;
  } catch(e) { return []; } 
}

function getEngineersByJanpad(janpad) {
  try { 
    var u = readUsersSheet();
    if (u.typeIdx === -1 || u.valueIdx === -1 || u.parentIdIdx === -1) return [];
    return u.data.filter(r => String(r[u.typeIdx]).trim() === "Engineer" && String(r[u.parentIdIdx]).trim() === janpad).map(r => String(r[u.valueIdx]).trim());
  } catch(e) { return []; }
}

function getPanchayatsByEngineer(janpad, engineerName) {
  try {
    var u = readUsersSheet();
    if (u.typeIdx === -1 || u.valueIdx === -1 || u.parentIdIdx === -1 || u.userIdIdx === -1) return [];
    var engineerUserId = "";
    for (var i = 0; i < u.data.length; i++) {
      if (String(u.data[i][u.typeIdx]).trim() === "Engineer" && String(u.data[i][u.valueIdx]).trim() === engineerName) { engineerUserId = String(u.data[i][u.userIdIdx]).trim(); break; }
    }
    if (!engineerUserId) return [];
    var result = [];
    for (var i = 0; i < u.data.length; i++) {
      if (String(u.data[i][u.typeIdx]).trim() === "Panchayat" && String(u.data[i][u.parentIdIdx]).trim() === engineerUserId) { result.push(String(u.data[i][u.valueIdx]).trim()); }
    }
    return result;
  } catch(e) { return []; }
}

function getEngineersForJanpad(j) { try { return getEngineersByJanpad(j); } catch(e) { return []; } }
function getPanchayatsForEngineer(j, e) { try { return getPanchayatsByEngineer(j, e); } catch(e) { return []; } }

function getDropdownData() {
  try {
    let janpads = getJanpads();
    let ts = [], as = [], heads = [], provisions = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sd = ss.getSheetByName("Database");
    if (sd && sd.getLastRow() > 1) { const dd = sd.getRange(2, 1, sd.getLastRow() - 1, 4).getValues(); ts = dd.filter(r => r[0] === "TS Authority").map(r => r[1]); as = dd.filter(r => r[0] === "AS Authority").map(r => r[1]); heads = dd.filter(r => r[0] === "Head").map(r => r[1]); provisions = dd.filter(r => r[0] === "Provision").map(r => r[1]); }
    return { janpads: janpads, engineers: {}, panchayats: {}, tsAuthorities: ts, asAuthorities: as, heads: heads, provisions: provisions };
  } catch(e) { return { janpads: [], engineers: {}, panchayats: {}, tsAuthorities: [], asAuthorities: [], heads: [], provisions: [] }; }
}

// ==================== LOGIN SYSTEM ====================
function authenticateUser(userId, password) {
  try { ensureUsersSheet(); } catch(e) { throw new Error('Users sheet setup failed: ' + e.message); }
  var u = readUsersSheet();
  if (u.data.length === 0) throw new Error('No user data found');
  if (u.userIdIdx === -1) throw new Error('UserID column not found!');
  if (u.passwordIdx === -1) throw new Error('Password column not found!');
  
  var mobileIdx = findColumnIdx(u.headers, ["Mobile", "mobile", "Phone", "phone", "Contact"]);
  var gmailIdx = findColumnIdx(u.headers, ["GMail", "Gmail", "gmail", "Email", "email", "E-Mail"]);
  var pwdGenDateIdx = findColumnIdx(u.headers, ["Password Generated Date", "Password  Generated Date", "Pwd Gen Date", "PwdGenDate", "PasswordGeneratedDate", "Password Gen Date", "Pwd Generated Date"]);
  var pwdExpDateIdx = findColumnIdx(u.headers, ["Password Expiry Date", "Password  Expiry Date", "Pwd Exp Date", "PwdExpDate", "PasswordExpiryDate", "Password Exp Date", "Pwd Expiry Date", "Expiry Date", "Password Expiration Date", "Expiry"]);
  
  for (var i = 0; i < u.data.length; i++) {
    var rowUserId = String(u.data[i][u.userIdIdx]).trim();
    var rowPassword = String(u.data[i][u.passwordIdx]).trim();
    var rowType = u.typeIdx !== -1 ? String(u.data[i][u.typeIdx]).trim() : "";
    var rowValue = u.valueIdx !== -1 ? String(u.data[i][u.valueIdx]).trim() : "";
    var rowStatus = u.statusIdx !== -1 ? String(u.data[i][u.statusIdx]).trim().toLowerCase() : "active";
    var rowParentId = u.parentIdIdx !== -1 ? String(u.data[i][u.parentIdIdx]).trim() : "";
    var rowMobile = mobileIdx !== -1 ? String(u.data[i][mobileIdx]).trim() : "";
    var rowGmail = gmailIdx !== -1 ? String(u.data[i][gmailIdx]).trim() : "";
    var rowPwdGenDate = pwdGenDateIdx !== -1 ? formatSheetDate(u.data[i][pwdGenDateIdx]) : "";
    var rowPwdExpDate = pwdExpDateIdx !== -1 ? formatSheetDate(u.data[i][pwdExpDateIdx]) : "";
    
    if (!rowUserId && !rowValue) continue;
    
    if (rowUserId === userId && rowPassword === password) {
      if (rowStatus !== "active") return { success: false, message: "Your account is inactive. Contact administrator." };
      if (pwdExpDateIdx !== -1) {
        var rawExpDate = u.data[i][pwdExpDateIdx];
        if (rawExpDate instanceof Date) {
          var today = new Date(); today.setHours(0, 0, 0, 0); var expDateOnly = new Date(rawExpDate); expDateOnly.setHours(0, 0, 0, 0);
          if (expDateOnly < today) { return { success: false, message: "Your password has expired on " + rowPwdExpDate + ". Contact administrator to reset." }; }
        }
      }
      var janpadName = ""; var engineerName = "";
      if (rowType === "Owner") { janpadName = rowParentId || "All"; engineerName = rowValue; } 
      else if (rowType === "Janpad") { janpadName = rowValue; } 
      else if (rowType === "Engineer") { janpadName = rowParentId; engineerName = rowValue; } 
      else if (rowType === "Panchayat") {
        for (var j = 0; j < u.data.length; j++) {
          var eType = u.typeIdx !== -1 ? String(u.data[j][u.typeIdx]).trim() : "";
          var eValue = u.valueIdx !== -1 ? String(u.data[j][u.valueIdx]).trim() : "";
          var eUserId = u.userIdIdx !== -1 ? String(u.data[j][u.userIdIdx]).trim() : "";
          var eParentId = u.parentIdIdx !== -1 ? String(u.data[j][u.parentIdIdx]).trim() : "";
          if ((eType === "Engineer" || eType === "Owner") && eUserId === rowParentId) { janpadName = eParentId; engineerName = eValue; break; }
        }
      }
      return { success: true, name: rowValue, janpad: janpadName, engineer: engineerName, userId: rowUserId, type: rowType, mobile: rowMobile, gmail: rowGmail, pwdGenDate: rowPwdGenDate, pwdExpDate: rowPwdExpDate };
    }
  }
  return { success: false, message: "Invalid User ID or Password." };
}

// ==================== OTP & PASSWORD RESET ====================
function sendOtpToUser(userId) {
  if (!userId) throw new Error("User ID is required");
  var u = readUsersSheet();
  if (u.data.length === 0 || u.userIdIdx === -1) throw new Error("User not found");
  var gmailIdx = findColumnIdx(u.headers, ["GMail", "Gmail", "gmail", "Email", "email", "E-Mail"]);
  for (var i = 0; i < u.data.length; i++) {
    if (String(u.data[i][u.userIdIdx]).trim() === userId) {
      var gmail = gmailIdx !== -1 ? String(u.data[i][gmailIdx]).trim() : "";
      if (!gmail) throw new Error("No email found for this User ID. Contact administrator.");
      var otp = Math.floor(1000 + Math.random() * 9000).toString();
      PropertiesService.getScriptProperties().setProperty('otp_' + userId, otp);
      var subject = "MeasurePro - Password Reset OTP";
      var body = "Your One Time Password (OTP) for password reset is: " + otp + "\n\nDo not share this OTP with anyone.\n\n- MeasurePro System";
      MailApp.sendEmail(gmail, subject, body);
      var emailParts = gmail.split('@');
      var maskedEmail = emailParts[0].substring(0, 2) + "***@" + emailParts[1];
      return { success: true, message: "OTP sent to " + maskedEmail };
    }
  }
  throw new Error("User ID not found");
}

function verifyOtpAndSendPassword(userId, otp) {
  if (!userId || !otp) throw new Error("User ID and OTP are required");
  var u = readUsersSheet();
  var storedOtp = PropertiesService.getScriptProperties().getProperty('otp_' + userId);
  if (!storedOtp) throw new Error("OTP expired or not requested. Please try again.");
  if (storedOtp !== otp.toString()) throw new Error("Invalid OTP. Please try again.");
  for (var i = 0; i < u.data.length; i++) {
    if (String(u.data[i][u.userIdIdx]).trim() === userId) {
      var gmailIdx = findColumnIdx(u.headers, ["GMail", "Gmail", "gmail", "Email", "email", "E-Mail"]);
      var password = u.passwordIdx !== -1 ? String(u.data[i][u.passwordIdx]).trim() : "";
      var gmail = gmailIdx !== -1 ? String(u.data[i][gmailIdx]).trim() : "";
      if (!gmail) throw new Error("No email found to send password.");
      var subject = "MeasurePro - Your Login Password";
      var body = "Your login credentials are:\n\nUser ID: " + userId + "\nPassword: " + password + "\n\nPlease delete this email after noting your password.\n\n- MeasurePro System";
      MailApp.sendEmail(gmail, subject, body);
      PropertiesService.getScriptProperties().deleteProperty('otp_' + userId);
      return { success: true, message: "Password sent to your email!" };
    }
  }
  throw new Error("User not found");
}

function submitEstimate(index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Saved Work');
  if (!sheet) throw new Error('Sheet "Saved Work" not found');
  const row = index + 2;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const si = headers.indexOf('Status');
  if (si === -1) throw new Error('Status column not found');
  sheet.getRange(row, si + 1).setValue('Estimated');
  return true;
}

function requestPasswordChangeOtp(userId, oldPassword) {
  if (!userId || !oldPassword) throw new Error("User ID and Old Password are required.");
  var u = readUsersSheet();
  if (!u.data || u.data.length === 0 || u.userIdIdx === -1) throw new Error("Users sheet data not found.");
  var userFound = false; var gmail = "";
  for (var i = 0; i < u.data.length; i++) {
    var rowUserId = String(u.data[i][u.userIdIdx]).trim();
    var rowPassword = String(u.data[i][u.passwordIdx]).trim();
    if (rowUserId === userId && rowPassword === oldPassword) {
      userFound = true;
      var gmailIdx = findColumnIdx(u.headers, ["GMail", "Gmail", "gmail", "Email", "email", "E-Mail"]);
      gmail = gmailIdx !== -1 ? String(u.data[i][gmailIdx]).trim() : "";
      break;
    }
  }
  if (!userFound) throw new Error("Old password is incorrect!");
  if (!gmail) throw new Error("No email registered with this account. Contact administrator.");
  var otp = Math.floor(1000 + Math.random() * 9000).toString();
  var now = new Date().getTime();
  PropertiesService.getScriptProperties().setProperty('chg_otp_' + userId, otp);
  PropertiesService.getScriptProperties().setProperty('chg_otp_time_' + userId, now.toString());
  var subject = "MeasurePro - Change Password OTP";
  var body = "Hello " + userId + ",\n\nYour OTP for changing password is: " + otp + "\n\nThis OTP is valid for 10 minutes only.\nDo not share this OTP with anyone.\n\n- MeasurePro System";
  MailApp.sendEmail(gmail, subject, body);
  var emailParts = gmail.split('@');
  var maskedEmail = emailParts[0].substring(0, 2) + "***@" + emailParts[1];
  return { success: true, message: "OTP sent to " + maskedEmail };
}

function changeUserPassword(userId, otp, newPassword) {
  if (!userId || !otp || !newPassword) throw new Error("All fields are required.");
  var props = PropertiesService.getScriptProperties();
  var storedOtp = props.getProperty('chg_otp_' + userId);
  var storedTime = props.getProperty('chg_otp_time_' + userId);
  if (!storedOtp || !storedTime) throw new Error("OTP expired or not requested. Please try again.");
  if (storedOtp !== otp.toString()) throw new Error("Invalid OTP.");
  var now = new Date().getTime();
  var otpTime = parseInt(storedTime);
  var diffMinutes = (now - otpTime) / (1000 * 60);
  if (diffMinutes > 10) { props.deleteProperty('chg_otp_' + userId); props.deleteProperty('chg_otp_time_' + userId); throw new Error("OTP has expired (10 min limit). Please request a new one."); }
  var u = readUsersSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  for (var i = 0; i < u.data.length; i++) {
    if (String(u.data[i][u.userIdIdx]).trim() === userId) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, u.passwordIdx + 1).setValue(newPassword);
      var genDate = new Date(); var expDate = new Date(genDate); expDate.setMonth(expDate.getMonth() + 3);
      var genDateStr = Utilities.formatDate(genDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      var expDateStr = Utilities.formatDate(expDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      if (u.pwdGenDateIdx !== -1) sheet.getRange(rowNum, u.pwdGenDateIdx + 1).setValue(genDateStr);
      if (u.pwdExpDateIdx !== -1) sheet.getRange(rowNum, u.pwdExpDateIdx + 1).setValue(expDateStr);
      break;
    }
  }
  props.deleteProperty('chg_otp_' + userId);
  props.deleteProperty('chg_otp_time_' + userId);
  return { success: true, message: "Password changed successfully!" };
}

function getDriveDrawings() {
  var folderId = "1rGNAgeZBluKPQ9K1qpaMQWv9-pR3dcfA"; 
  var filesData = [];
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      filesData.push({ name: file.getName(), id: file.getId(), mime: file.getMimeType() });
    }
  } catch(e) { Logger.log("Drive Error: " + e.message); }
  return filesData;
}

function resetUserPasswordDates(index) {
  if (index === null || index === undefined) throw new Error("Invalid user index");
  var u = readUsersSheet();
  var row = parseInt(index) + 2;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error("Users sheet not found");
  var genDate = new Date(); var expDate = new Date(genDate); expDate.setMonth(expDate.getMonth() + 3);
  var genDateStr = Utilities.formatDate(genDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var expDateStr = Utilities.formatDate(expDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
  if (u.pwdGenDateIdx !== -1) sheet.getRange(row, u.pwdGenDateIdx + 1).setValue(genDateStr);
  if (u.pwdExpDateIdx !== -1) sheet.getRange(row, u.pwdExpDateIdx + 1).setValue(expDateStr);
  return { genDate: genDateStr, expDate: expDateStr };
}
