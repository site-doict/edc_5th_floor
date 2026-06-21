const SPREADSHEET_NAME = 'SET Centre Inspection Responses';
const SHEET_NAME = 'Responses';
const SPREADSHEET_ID_KEY = 'SET_CENTRE_INSPECTION_SPREADSHEET_ID';

const HEADERS = [
  'received_at',
  'submitted_at',
  'completion_status',
  'ap_name',
  'ap_id',
  'mobile',
  'district',
  'upazila',
  'inspection_date',
  'centre_name',
  'visit_type',
  'overall_status',
  'usable',
  'overall_comment',
  'civil_comment',
  'elec_general_comment',
  'elec_durable_comment',
  'lightning_comment',
  'solar_panel_brand',
  'solar_inverter_brand',
  'solar_comment',
  'drive_civil',
  'drive_elec',
  'drive_solar',
  'drive_warranty',
  'quality',
  'major_issues',
  'recommendation',
  'issue_summary',
  'elec_bill',
  'elec_bill_detail',
  'ap_remarks',
  'civil_checklist_json',
  'elec_general_checklist_json',
  'elec_durable_checklist_json',
  'lightning_checklist_json',
  'solar_checklist_json'
];

function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

  if (action === 'latest') {
    const payload = JSON.stringify(getLatestRow_());
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.parent.postMessage({source:"set-centre-inspection-load", data:' + payload + '}, "*");' +
      'document.body.textContent="loaded";' +
      '</script></body></html>'
    );
  }

  return ContentService.createTextOutput('Google Sheets save endpoint is running.');
}

function doPost(e) {
  let lock = null;
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const sheet = getOrCreateSheet_();
    const values = HEADERS.map((header) => getValue_(e, header));
    values[0] = new Date().toISOString();
    sheet.appendRow(values);

    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.parent.postMessage({source:"set-centre-inspection-save", ok:true}, "*");' +
      'document.body.textContent="saved";' +
      '</script></body></html>'
    );
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.parent.postMessage({source:"set-centre-inspection-save", ok:false, error:' + JSON.stringify(error && error.message ? error.message : String(error)) + '}, "*");' +
      'document.body.textContent="failed";' +
      '</script></body></html>'
    );
  } finally {
    try {
      if (lock) lock.releaseLock();
    } catch (e) {
      // ignore release errors when lock acquisition failed
    }
  }
}

function getOrCreateSheet_() {
  const spreadsheet = getOrCreateSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(SPREADSHEET_ID_KEY);

  if (storedId) {
    return SpreadsheetApp.openById(storedId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty(SPREADSHEET_ID_KEY, active.getId());
    shareSpreadsheetReadOnly_(active.getId());
    return active;
  }

  const created = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty(SPREADSHEET_ID_KEY, created.getId());
  shareSpreadsheetReadOnly_(created.getId());
  return created;
}

function shareSpreadsheetReadOnly_(spreadsheetId) {
  DriveApp.getFileById(spreadsheetId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

function getValue_(e, key) {
  if (!e || !e.parameter || typeof e.parameter[key] === 'undefined') {
    return '';
  }
  return e.parameter[key];
}

function getLatestRow_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { ok: true, hasData: false, row: null };
  }

  const values = sheet.getRange(lastRow, 1, 1, HEADERS.length).getValues()[0];
  const row = {};

  HEADERS.forEach((header, index) => {
    row[header] = values[index] || '';
  });

  return { ok: true, hasData: true, row: row };
}