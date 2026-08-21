/**
 * WA SIMPLE — Masterclass lead capture endpoint.
 *
 * Receives the registration form POST from index.html and appends one row
 * per lead to the "Leads" sheet, creating the sheet + header row on first run.
 *
 * Deploy: Extensions > Apps Script > paste this file > Deploy > New deployment
 *         > Web app > Execute as: Me > Who has access: Anyone > Deploy.
 * Then copy the /exec URL into SHEET_ENDPOINT in index.html.
 */

/** Leave empty to use the spreadsheet this script is bound to, or paste a spreadsheet ID. */
var SPREADSHEET_ID = '';
var SHEET_NAME = 'Leads';

/** Column order of the sheet. Add a key here and it starts getting written — no other change needed. */
var FIELDS = [
  'received_at',
  'full_name',
  'whatsapp_number',
  'email',
  'business_industry',
  'annual_turnover',
  'current_setup',
  'biggest_priority',
  'monthly_volume',
  'implementation_timeline',
  'consent',
  'utm_source',
  'utm_medium',
  'utm_campaign',
  'utm_campaign_id',
  'utm_adset',
  'utm_adset_id',
  'utm_ad',
  'utm_ad_id',
  'utm_content',
  'utm_term',
  'utm_placement',
  'device_platform',
  'gclid',
  'fbclid',
  'landing_page_url',
  'first_page_url',
  'current_page_url',
  'referrer',
  'device',
  'browser',
  'os',
  'timestamp'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var data = parseBody_(e);
    var clean = validate_(data);
    if (clean.error) return json_({ status: 'error', message: clean.error });

    var sheet = getSheet_();
    sheet.appendRow(FIELDS.map(function (key) { return clean.row[key] || ''; }));

    return json_({ status: 'ok' });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Lets you open the /exec URL in a browser to confirm the deployment is live. */
function doGet() {
  return json_({ status: 'ok', message: 'WA SIMPLE lead endpoint is live' });
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  var raw = e.postData.contents;
  try {
    return JSON.parse(raw);
  } catch (err) {
    // form-encoded fallback
    var out = {};
    raw.split('&').forEach(function (pair) {
      var kv = pair.split('=');
      if (kv[0]) out[decodeURIComponent(kv[0])] = decodeURIComponent((kv[1] || '').replace(/\+/g, ' '));
    });
    return out;
  }
}

/** Server-side mirror of the browser validation — never trust the client alone. */
function validate_(d) {
  var name = String(d.full_name || '').trim().replace(/\s+/g, ' ');
  var phone = String(d.whatsapp_number || '').replace(/\D/g, '').replace(/^0+/, '').replace(/^91(?=\d{10}$)/, '');
  var email = String(d.email || '').trim().toLowerCase();

  if (name.length < 3) return { error: 'Invalid name' };
  if (!/^[6-9]\d{9}$/.test(phone)) return { error: 'Invalid WhatsApp number' };
  if (!/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(email)) return { error: 'Invalid email' };
  if (!d.business_industry) return { error: 'Industry missing' };
  if (String(d.consent).toLowerCase() !== 'yes') return { error: 'Consent missing' };

  var row = {};
  FIELDS.forEach(function (key) { row[key] = String(d[key] == null ? '' : d[key]).slice(0, 500); });
  row.received_at = new Date();
  row.full_name = name;
  row.whatsapp_number = "'" + phone; // leading quote keeps Sheets from dropping a leading digit / using sci-notation
  row.email = email;

  return { row: row };
}

function getSheet_() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FIELDS);
    sheet.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
