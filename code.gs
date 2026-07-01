const SPREADSHEET_ID = '147dN1g-U_rDnRLXsvWODFQ_ROrl06yt8BylC1CAj6Kc';

const SHEETS = {
  lead: 'Leads',
  revenue: 'Revenue',
  adSpend: 'AdSpend',
  campaign: 'Campaigns',
  product: 'Products',
  followup: 'Followups'
};

const SCHEMAS = {
  USER: ['Username', 'Password', 'Name', 'Role', 'Active'],
  Leads: ['id', 'date', 'name', 'phone', 'customId', 'type', 'channel', 'campaign', 'campaignId', 'adSet', 'adName', 'product', 'status', 'owner', 'note'],
  Revenue: ['leadId', 'name', 'phone', 'product', 'date', 'paymentStatus', 'amount', 'grossProfit', 'receipt', 'note'],
  AdSpend: ['date', 'channel', 'product', 'campaign', 'adSet', 'adName', 'amount', 'impressions', 'clicks', 'messages', 'leads'],
  Campaigns: ['id', 'channel', 'product', 'campaign', 'adSet', 'adName', 'objective', 'startDate', 'endDate', 'status', 'imageFileId', 'imageUrl', 'note'],
  Products: ['id', 'name', 'group', 'department', 'serviceLine', 'channel', 'status', 'note'],
  Followups: ['leadId', 'customerName', 'phone', 'product', 'channel', 'campaign', 'leadStatus', 'date', 'time', 'result', 'nextDate', 'updatedBy', 'note']
};

function doGet(e) {
  const action = e.parameter.action || 'read';
  if (action === 'updateLeadStatus') {
    return withScriptLock(() => updateLeadStatus({ leadId: e.parameter.leadId, status: e.parameter.status }));
  }
  if (action === 'updateLead') return withScriptLock(() => updateLead(JSON.parse(e.parameter.payload || '{}')));
  if (action === 'deleteLead') return withScriptLock(() => deleteLead(e.parameter.leadId));
  if (action === 'updateCampaign') return withScriptLock(() => updateCampaign(JSON.parse(e.parameter.payload || '{}')));
  if (action !== 'read') return json({ ok: false, error: 'Unknown action' });
  withScriptLock(() => ensureMasterIds());
  return json({
    ok: true,
    leads: readObjects('Leads'),
    revenue: readObjects('Revenue'),
    adSpend: readObjects('AdSpend'),
    campaigns: readObjects('Campaigns'),
    products: readObjects('Products'),
    followups: readObjects('Followups')
  });
}

function setupSystemSheets() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SCHEMAS).forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    const headers = SCHEMAS[sheetName];
    if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    else {
      const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
      const normalizedCurrent = current.map(normalizeHeader);
      const merged = current.filter(Boolean).concat(headers.filter(header => !normalizedCurrent.includes(normalizeHeader(header))));
      sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#dff3fb');
    if (sheetName === 'Leads') formatLeadTextColumns(sheet);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
  ensureMasterIds();
  return 'สร้าง Sheet หัวคอลัมน์ และรหัสอัตโนมัติเรียบร้อยแล้ว';
}

function formatLeadTextColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader);
  ['id', 'leadid', 'phone', 'customid'].forEach(columnName => {
    const column = headers.indexOf(columnName);
    if (column >= 0 && sheet.getMaxRows() > 1) sheet.getRange(2, column + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  });
}

function repairLeadIdsAndDates() {
  const sheet = getSheet('Leads');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 'ไม่มีข้อมูล Lead ที่ต้องซ่อม';
  const headers = data[0].map(normalizeHeader);
  const idColumn = headers.findIndex(header => ['id', 'leadid'].includes(header));
  const dateColumn = headers.findIndex(header => ['date', 'วันที่'].includes(header));
  const phoneColumn = headers.findIndex(header => ['phone', 'เบอร์โทร', 'เบอร์โทรศัพท์'].includes(header));
  const customIdColumn = headers.findIndex(header => ['customid'].includes(header));
  if (idColumn < 0 || dateColumn < 0) throw new Error('ไม่พบคอลัมน์ id หรือ date ใน Sheet Leads');

  const now = new Date();
  const prefix = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMM');
  let repaired = 0;
  for (let index = 1; index < data.length; index++) {
    const rowNumber = index + 1;
    let phone = phoneColumn >= 0 ? String(data[index][phoneColumn] || '').replace(/\.0$/, '') : '';
    if (/^\d{9}$/.test(phone)) phone = `0${phone}`;
    if (!data[index][idColumn]) {
      sheet.getRange(rowNumber, idColumn + 1).setValue(`LD-${prefix}-${String(index).padStart(4, '0')}`);
      repaired++;
    }
    if (!data[index][dateColumn]) sheet.getRange(rowNumber, dateColumn + 1).setValue(now).setNumberFormat('yyyy-mm-dd');
    if (phoneColumn >= 0 && phone) sheet.getRange(rowNumber, phoneColumn + 1).setNumberFormat('@').setValue(phone);
    if (customIdColumn >= 0 && phone && (!data[index][customIdColumn] || String(data[index][customIdColumn]).startsWith('Auto'))) {
      sheet.getRange(rowNumber, customIdColumn + 1).setNumberFormat('@').setValue(`CUS-${phone}`);
    }
  }
  return `ซ่อมข้อมูล Lead แล้ว ${repaired} รายการ`;
}

function updateLead(payload) {
  const result = findLeadRow(payload.id);
  if (!result) return json({ ok: false, error: 'ไม่พบลีดที่ต้องการแก้ไข' });
  const headers = result.headers;
  result.sheet.getRange(result.row, 1, 1, headers.length).setValues([headers.map(header => {
    const payloadKey = Object.keys(payload).find(key => normalizeHeader(key) === normalizeHeader(header));
    return payloadKey ? payload[payloadKey] : result.values[header];
  })]);
  return json({ ok: true });
}

function deleteLead(leadId) {
  const result = findLeadRow(leadId);
  if (!result) return json({ ok: false, error: 'ไม่พบลีดที่ต้องการลบ' });
  result.sheet.deleteRow(result.row);
  return json({ ok: true });
}

function updateCampaign(payload) {
  const result = findRowById('Campaigns', payload.id, ['id', 'campaign_id', 'campaignid', 'campaign id']);
  if (!result) return json({ ok: false, error: 'ไม่พบแคมเปญที่ต้องการแก้ไข' });
  result.sheet.getRange(result.row, 1, 1, result.headers.length).setValues([result.headers.map(header => {
    const payloadKey = Object.keys(payload).find(key => normalizeHeader(key) === normalizeHeader(header));
    return payloadKey ? payload[payloadKey] : result.values[header];
  })]);
  return json({ ok: true });
}

function findRowById(sheetName, id, idAliases) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(String);
  const normalizedAliases = idAliases.map(normalizeHeader);
  const idColumn = headers.findIndex(header => normalizedAliases.includes(normalizeHeader(header)));
  if (idColumn < 0) return null;
  const rowIndex = data.findIndex((row, index) => index > 0 && String(row[idColumn]) === String(id));
  if (rowIndex < 0) return null;
  const values = {};
  headers.forEach((header, index) => values[header] = data[rowIndex][index]);
  return { sheet, headers, values, row: rowIndex + 1 };
}

function findLeadRow(leadId) {
  const sheet = getSheet('Leads');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(String);
  const normalized = headers.map(header => header.trim().toLowerCase());
  const idColumn = normalized.findIndex(header => ['id', 'lead_id', 'leadid', 'lead id'].includes(header));
  if (idColumn < 0) return null;
  const rowIndex = data.findIndex((row, index) => index > 0 && String(row[idColumn]) === String(leadId));
  if (rowIndex < 0) return null;
  const values = {};
  headers.forEach((header, index) => values[header] = data[rowIndex][index]);
  return { sheet, headers, values, row: rowIndex + 1 };
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9ก-๙]/g, '');
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.action === 'login') return login(body);
  if (body.action === 'uploadCampaignImage') return uploadCampaignImage(body);
  if (body.action === 'updateLeadStatus') return withScriptLock(() => updateLeadStatus(body));
  if (body.action !== 'append') return json({ ok: false, error: 'Unknown action' });
  const sheetName = SHEETS[body.type];
  if (!sheetName) return json({ ok: false, error: 'Unknown data type' });
  withScriptLock(() => appendObject(sheetName, body.payload || {}));
  return json({ ok: true });
}

function withScriptLock(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function uploadCampaignImage(body) {
  const allowedTypes = ['image/jpeg', 'image/png', 'image/webp'];
  if (!allowedTypes.includes(body.mimeType)) return json({ ok: false, error: 'รองรับเฉพาะ JPG, PNG และ WebP' });
  const bytes = Utilities.base64Decode(body.base64 || '');
  if (bytes.length > 5 * 1024 * 1024) return json({ ok: false, error: 'ไฟล์มีขนาดเกิน 5 MB' });
  const folders = DriveApp.getFoldersByName('Central Lead Log - Campaign Assets');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Central Lead Log - Campaign Assets');
  const safeName = `${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss')}-${String(body.fileName || 'campaign-image').replace(/[^a-zA-Z0-9ก-๙._-]/g, '-')}`;
  const file = folder.createFile(Utilities.newBlob(bytes, body.mimeType, safeName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return json({ ok: true, fileId: file.getId(), url: `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1200` });
}

function updateLeadStatus(body) {
  const sheet = getSheet('Leads');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return json({ ok: false, error: 'ไม่พบข้อมูลลีด' });
  const headers = values[0].map(value => String(value).trim().toLowerCase());
  const idColumn = headers.findIndex(header => ['id', 'lead_id', 'leadid', 'lead id'].includes(header));
  const statusColumn = headers.findIndex(header => ['status', 'lead_status', 'lead status', 'สถานะ'].includes(header));
  if (idColumn < 0 || statusColumn < 0) return json({ ok: false, error: 'ไม่พบคอลัมน์ id หรือ status' });
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[idColumn]) === String(body.leadId));
  if (rowIndex < 0) return json({ ok: false, error: 'ไม่พบ Lead_ID ที่ต้องการ' });
  sheet.getRange(rowIndex + 1, statusColumn + 1).setValue(body.status);
  return json({ ok: true });
}

function login(body) {
  const username = String(body.username || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!username || !password) return json({ ok: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' });

  const users = readObjects('USER');
  const user = users.find(row => String(valueOf(row, 'Username')).trim().toLowerCase() === username);
  if (!user) return json({ ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });

  const active = String(valueOf(user, 'Active') || 'TRUE').trim().toLowerCase();
  if (['false', '0', 'inactive', 'no'].includes(active)) return json({ ok: false, error: 'บัญชีนี้ถูกปิดใช้งาน' });

  const storedPassword = String(valueOf(user, 'Password'));
  const passwordHash = sha256(password);
  if (storedPassword !== password && storedPassword.toLowerCase() !== passwordHash) {
    return json({ ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
  }

  return json({
    ok: true,
    user: {
      username: valueOf(user, 'Username'),
      name: valueOf(user, 'Name') || valueOf(user, 'Username'),
      role: valueOf(user, 'Role') || 'User'
    }
  });
}

function valueOf(object, key) {
  const actualKey = Object.keys(object).find(item => item.toLowerCase() === key.toLowerCase());
  return actualKey ? object[actualKey] : '';
}

function sha256(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'))
    .join('');
}

function setupUserSheet() {
  const sheet = getSheet('USER');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Username', 'Password', 'Name', 'Role', 'Active']);
  }
}

function readObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(Boolean)).map(row => {
    const object = {};
    headers.forEach((header, index) => object[header] = row[index]);
    return object;
  });
}

function appendObject(sheetName, object) {
  const sheet = getSheet(sheetName);
  object = Object.assign({}, object);
  if ((sheetName === 'Products' || sheetName === 'Campaigns') && !String(object.id || '').trim()) {
    object.id = nextSheetMasterId(sheetName, object);
  }
  let headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String) : [];
  if (sheetName === 'Products' || sheetName === 'Campaigns') {
    const idColumn = masterIdColumnIndex(headers, sheetName);
    if (idColumn >= 0) {
      const idHeader = headers[idColumn];
      object[idHeader] = object.id;
      if (normalizeHeader(idHeader) !== 'id') delete object.id;
    }
  }
  const incoming = Object.keys(object);
  const missing = incoming.filter(key => !headers.includes(key));
  if (!headers.length) {
    headers = incoming;
    sheet.appendRow(headers);
  } else if (missing.length) {
    headers = headers.concat(missing);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.appendRow(headers.map(header => object[header] ?? ''));
}

function ensureMasterIds() {
  ['Products', 'Campaigns'].forEach(sheetName => {
    const sheet = getSheet(sheetName);
    if (sheet.getLastRow() < 2) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const idColumn = masterIdColumnIndex(headers, sheetName);
    if (idColumn < 0) return;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    rows.forEach((row, index) => {
      if (String(row[idColumn] || '').trim() || !row.some(Boolean)) return;
      const payload = {};
      headers.forEach((header, column) => payload[header] = row[column]);
      sheet.getRange(index + 2, idColumn + 1).setValue(nextSheetMasterId(sheetName, payload));
    });
  });
}

function nextSheetMasterId(sheetName, payload) {
  const sheet = getSheet(sheetName);
  const periodSource = payload.startDate || payload['Start Date'] || new Date();
  const parsedDate = periodSource instanceof Date ? periodSource : new Date(`${periodSource}T00:00:00`);
  const periodDate = Number.isNaN(parsedDate.getTime()) ? new Date() : parsedDate;
  const period = Utilities.formatDate(periodDate, 'Asia/Bangkok', 'yyyyMM');
  const prefix = sheetName === 'Products' ? 'PROD-' : `CAM-${period}-`;
  const digits = sheetName === 'Products' ? 3 : 4;
  if (sheet.getLastRow() < 2) return `${prefix}${String(1).padStart(digits, '0')}`;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idColumn = masterIdColumnIndex(headers, sheetName);
  if (idColumn < 0) return `${prefix}${String(1).padStart(digits, '0')}`;
  const ids = sheet.getRange(2, idColumn + 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
  const highest = ids.reduce((max, id) => {
    if (!id.startsWith(prefix)) return max;
    const sequence = Number(id.slice(prefix.length));
    return Number.isFinite(sequence) ? Math.max(max, sequence) : max;
  }, 0);
  return `${prefix}${String(highest + 1).padStart(digits, '0')}`;
}

function masterIdColumnIndex(headers, sheetName) {
  const preferred = sheetName === 'Products' ? 'productid' : 'campaignid';
  const preferredIndex = headers.findIndex(header => normalizeHeader(header) === preferred);
  if (preferredIndex >= 0) return preferredIndex;
  return headers.findIndex(header => normalizeHeader(header) === 'id');
}

function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
