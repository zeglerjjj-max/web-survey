/**
 * RESPON-12 survey backend for Google Apps Script.
 * Setup:
 * 1. Run setupSpreadsheet() once and authorize access.
 * 2. Copy the Sheet URL from the execution log.
 * 3. Deploy > New deployment > Web app; Execute as: Me; Who has access: Anyone.
 * 4. Paste the deployment URL ending in /exec into index.html.
 *
 * This script records only survey answers, a response ID, consent, and server timestamp.
 * It does not intentionally collect names, email addresses, or IP addresses.
 */

const SHEET_NAME = 'Jawaban Survei';
const PROPERTY_KEY = 'RESPON12_SPREADSHEET_ID';

function setupSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROPERTY_KEY);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('RESPON-12 — Jawaban Survei');
    props.setProperty(PROPERTY_KEY, ss.getId());
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp server', 'Response ID', 'Persetujuan', ...Array.from({ length: 20 }, (_, i) => `Q${i + 1}`)];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
  Logger.log('Spreadsheet dibuat/siap: ' + ss.getUrl());
  return ss.getUrl();
}

function doPost(e) {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY);
    if (!id) return jsonResponse({ ok: false, error: 'Spreadsheet belum disiapkan.' });

    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    const payload = JSON.parse(raw);
    if (!payload || payload.consent !== true || !payload.responseId || typeof payload.answers !== 'object') {
      return jsonResponse({ ok: false, error: 'Payload tidak valid.' });
    }

    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ ok: false, error: 'Sheet jawaban tidak ditemukan.' });

    // Avoid accidental duplicate writes if a client retries the same response ID.
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const found = sheet.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(String(payload.responseId))
        .matchEntireCell(true)
        .findNext();
      if (found) return jsonResponse({ ok: true, duplicate: true });
    }

    const a = payload.answers;
    const row = [new Date(), safeCell(payload.responseId), 'Setuju'];
    for (let i = 1; i <= 20; i++) {
      const value = a[`q${i}`];
      row.push(safeCell(Array.isArray(value) ? value.join(' | ') : (value == null ? '' : String(value))));
    }
    sheet.appendRow(row);
    return jsonResponse({ ok: true });
  } catch (err) {
    console.error(err);
    return jsonResponse({ ok: false, error: 'Gagal memproses jawaban.' });
  }
}

// Reduce the risk of spreadsheet formula injection from free-text survey responses.
function safeCell(value) {
  let text = String(value == null ? '' : value).slice(0, 5000);
  if (/^[\s]*[=+@\-]/.test(text)) text = "'" + text;
  return text;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
