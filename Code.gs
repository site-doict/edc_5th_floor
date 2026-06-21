const SHEET_NAME = 'Responses';

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

function doGet() {
  return ContentService.createTextOutput('Google Sheets save endpoint is running.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getOrCreateSheet_();
    const values = HEADERS.map((header) => getValue_(e, header));
    values[0] = new Date().toISOString();
    sheet.appendRow(values);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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

function getValue_(e, key) {
  if (!e || !e.parameter || typeof e.parameter[key] === 'undefined') {
    return '';
  }
  return e.parameter[key];
}