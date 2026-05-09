const SHEET_NAME = 'HRV紀錄';

const HEADERS = [
  'recordId',
  'timestamp',
  'email',
  'patientName',
  'patientBirthDate',
  'patientGender',
  'measureDate',
  'stressIndex',
  'emotionalFlexibility',
  'relaxationIndex',
  'vitalityIndex',
  'averageHeartRate',
  'duration',
  'note',
  'assessmentTitle',
  'assessmentLevel',
  'assessmentObservation',
  'assessmentRecommendation',
  'brsrSleep',
  'brsrAnxiety',
  'brsrAnger',
  'brsrDepression',
  'brsrInferiority',
  'brsrSuicide',
  'brsrTotal',
  'brsrLevel',
];

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const sheet = getSheet_();
    ensureHeaders_(sheet);
    if (payload.recordId && recordExists_(sheet, payload.recordId)) return jsonOutput_({ ok: true, deduped: true });
    const writeResult = appendPayload_(sheet, payload);
    return jsonOutput_({ ok: true, payloadEmail: payload.email || '', row: writeResult.row, headers: writeResult.headers });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function parsePayload_(e) {
  const params = (e && e.parameter) || {};
  if (params.payload) return JSON.parse(params.payload);
  if (params.email) return params;

  const contents = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!contents) return {};

  try {
    return JSON.parse(contents);
  } catch (error) {
    const decoded = decodeURIComponent(contents.replace(/\+/g, ' '));
    const match = decoded.match(/(?:^|&)payload=(.*)$/);
    if (match) return JSON.parse(match[1]);
    throw error;
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};

  try {
    if (params.action === 'save') {
      const payload = parsePayload_(e);
      const sheet = getSheet_();
      ensureHeaders_(sheet);
      if (payload.recordId && recordExists_(sheet, payload.recordId)) return scriptOrJson_(params, { ok: true, deduped: true });
      const writeResult = appendPayload_(sheet, payload);
      return scriptOrJson_(params, { ok: true, payloadEmail: payload.email || '', row: writeResult.row, headers: writeResult.headers });
    }

    if (params.action === 'debug') {
      const sheet = getSheet_();
      ensureHeaders_(sheet);
      const values = sheet.getDataRange().getValues();
      const headers = values[0] || [];
      const lastRows = values.slice(Math.max(1, values.length - 5)).map((row) =>
        row.map((value) => (value instanceof Date ? Utilities.formatDate(value, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : value)),
      );
      return scriptOrJson_(params, {
        ok: true,
        sheetName: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn(),
        headers,
        lastRows,
      });
    }

    if (params.action === 'history') {
      const email = String(params.email || '').trim().toLowerCase();
      if (!email) return scriptOrJson_(params, { ok: false, error: 'missing email' });

      const sheet = getSheet_();
      ensureHeaders_(sheet);
      const records = readRecords_(sheet)
        .filter((record) => String(record.email || '').trim().toLowerCase() === email)
        .sort((a, b) => String(a.measureDate || a.timestamp || '').localeCompare(String(b.measureDate || b.timestamp || '')));

      return scriptOrJson_(params, { ok: true, records });
    }

    return scriptOrJson_(params, { ok: true, message: 'HRV web app is running.' });
  } catch (error) {
    return scriptOrJson_(params, { ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function appendPayload_(sheet, payload) {
  const headers = getUniqueHeaders_(sheet);
  const row = headers.map((header) => {
    if (!header) return '';
    if (header === 'timestamp') return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    return Object.prototype.hasOwnProperty.call(payload, header) ? String(payload[header]) : '';
  });
  const nextRow = sheet.getLastRow() + 1;
  const range = sheet.getRange(nextRow, 1, 1, headers.length);
  range.setNumberFormat('@');
  range.setValues([row]);
  return { row: nextRow, headers };
}

function recordExists_(sheet, recordId) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return false;
  const headers = values[0].map((header) => String(header || '').trim());
  const index = headers.indexOf('recordId');
  if (index === -1) return false;
  return values.slice(1).some((row) => String(row[index] || '') === String(recordId));
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((header) => String(header || '').trim());
  const hasHeaders = firstRow.some((value) => value);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  const missingHeaders = HEADERS.filter((header) => !firstRow.includes(header));
  if (missingHeaders.length === 0) return;

  const nextColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, nextColumn, 1, missingHeaders.length).setValues([missingHeaders]);
}

function readRecords_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = getUniqueHeaders_(sheet);
  return values.slice(1).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      if (!header) return;
      const value = row[index];
      record[header] = value instanceof Date ? Utilities.formatDate(value, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : value;
    });
    return record;
  });
}

function getUniqueHeaders_(sheet) {
  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((header) => String(header || '').trim());
  const seen = {};
  return rawHeaders.map((header) => {
    if (!header) return '';
    seen[header] = (seen[header] || 0) + 1;
    return seen[header] === 1 ? header : '';
  });
}

function scriptOrJson_(params, data) {
  if (params.callback) {
    const callback = String(params.callback).replace(/[^\w$.]/g, '');
    return ContentService.createTextOutput(`${callback}(${JSON.stringify(data)});`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonOutput_(data);
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function repairHrvSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const backupName = `${SHEET_NAME}_備份_${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss')}`;
  sheet.copyTo(spreadsheet).setName(backupName);

  const records = migrateRows_(values);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  if (records.length > 0) {
    const rows = records.map((record) => HEADERS.map((header) => record[header] || ''));
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
  return `完成整理：已備份為 ${backupName}，保留 ${records.length} 筆資料。`;
}

function migrateRows_(values) {
  if (!values || values.length < 2) return [];

  const rawHeaders = values[0].map((header) => String(header || '').trim());
  const records = [];
  const seenRecordIds = {};

  values.slice(1).forEach((row, rowIndex) => {
    const record = {};

    rawHeaders.forEach((header, columnIndex) => {
      const key = normalizeHeader_(header);
      if (!key || record[key]) return;
      const value = normalizeCellValue_(row[columnIndex], key);
      if (value !== '') record[key] = value;
    });

    if (!hasAnyRecordValue_(record)) return;
    if (!record.recordId) record.recordId = `migrated-${rowIndex + 2}-${new Date().getTime()}`;
    if (seenRecordIds[record.recordId]) return;
    seenRecordIds[record.recordId] = true;
    records.push(record);
  });

  return records;
}

function normalizeHeader_(header) {
  const key = String(header || '').trim();
  const map = {
    上傳時間: 'timestamp',
    測量日期: 'measureDate',
    量測日期: 'measureDate',
    Email: 'email',
    email: 'email',
    壓力指數: 'stressIndex',
    情緒反應靈活度: 'emotionalFlexibility',
    情緒靈活度: 'emotionalFlexibility',
    放鬆指數: 'relaxationIndex',
    活力指數: 'vitalityIndex',
    平均心跳: 'averageHeartRate',
    測量時間: 'duration',
    量測時間: 'duration',
    備註: 'note',
  };
  return map[key] || (HEADERS.includes(key) ? key : '');
}

function normalizeCellValue_(value, key) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (key === 'duration') return Utilities.formatDate(value, 'Asia/Taipei', 'H:mm');
    if (key === 'measureDate' || key === 'patientBirthDate') return Utilities.formatDate(value, 'Asia/Taipei', 'yyyy/MM/dd');
    return Utilities.formatDate(value, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  }
  return value;
}

function hasAnyRecordValue_(record) {
  return ['email', 'measureDate', 'stressIndex', 'emotionalFlexibility', 'relaxationIndex', 'vitalityIndex', 'note'].some((key) => record[key] !== undefined && record[key] !== '');
}
