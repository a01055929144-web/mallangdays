// Mallangdays Web App
// Script Properties:
// - ANTHROPIC_API_KEY: Claude API key
// - SHEET_ID: Google Sheet ID

const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const DEFAULT_MODEL = 'claude-sonnet-4-6';
const DEFAULT_SHEET_ID = '17JnY3k5j9ySX6TGbBAr3LNAubjGYY1erwMbNXGT2CPE';

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  if (action === 'read') return json_(readRows_());
  if (action === 'append') return json_(appendRow_(parseRow_(e.parameter.row)));
  if (action === 'updateAppend') {
    return json_(updateAppend_(Number(e.parameter.rowIdx), e.parameter.col, e.parameter.value || ''));
  }
  return json_({ ok: true, service: 'mallangdays' });
}

function doPost(e) {
  const params = e.parameter || {};
  const action = (params.action || '').trim();
  try {
    if (action === 'claude') return json_(callClaude_(params.payload));
    if (action === 'claudeStatus') return json_(claudeStatus_());
    if (action === 'update') {
      return json_(updateCell_(Number(params.rowIdx), params.col, params.value || ''));
    }
    return json_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function callClaude_(payloadText) {
  const apiKey = getAnthropicKey_();
  const payload = JSON.parse(payloadText || '{}');
  payload.model = payload.model || DEFAULT_MODEL;

  const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  const data = JSON.parse(text || '{}');
  if (status < 200 || status >= 300) {
    throw new Error((data.error && data.error.message) || text || ('Anthropic API ' + status));
  }
  return { ok: true, data: data };
}

function claudeStatus_() {
  getAnthropicKey_();
  return { ok: true };
}

function readRows_() {
  const sheet = getSheet_();
  return { ok: true, data: sheet.getDataRange().getValues() };
}

function appendRow_(row) {
  const sheet = getSheet_();
  sheet.appendRow(row);
  return { ok: true, rowIdx: sheet.getLastRow() };
}

function updateCell_(rowIdx, col, value) {
  if (!rowIdx || !col) throw new Error('rowIdx와 col이 필요합니다.');
  getSheet_().getRange(rowIdx, colToIndex_(col)).setValue(value);
  return { ok: true };
}

function updateAppend_(rowIdx, col, value) {
  if (!rowIdx || !col) throw new Error('rowIdx와 col이 필요합니다.');
  const sheet = getSheet_();
  const cell = sheet.getRange(rowIdx, colToIndex_(col));
  const prev = cell.getValue();
  cell.setValue(String(prev || '') + String(value || ''));
  return { ok: true };
}

function parseRow_(text) {
  const row = JSON.parse(text || '[]');
  if (!Array.isArray(row)) throw new Error('row는 배열이어야 합니다.');
  return row;
}

function getSheet_() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || DEFAULT_SHEET_ID;
  return SpreadsheetApp.openById(sheetId).getSheets()[0];
}

function getAnthropicKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('Script Properties에 ANTHROPIC_API_KEY가 없습니다.');
  return key;
}

function colToIndex_(col) {
  if (/^\d+$/.test(String(col))) return Number(col);
  return String(col).toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
