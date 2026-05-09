const SHEET_NAME = 'HRV紀錄';

const HEADERS = [
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
    const payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
    const sheet = getSheet_();
    ensureHeaders_(sheet);
    const row = HEADERS.map((header) => (header === 'timestamp' ? new Date() : payload[header] || ''));
    sheet.appendRow(row);
    return jsonOutput_({ ok: true });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};

  try {
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

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
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

  const headers = values[0].map((header) => String(header || '').trim());
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
