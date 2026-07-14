/**
 * London University Developments Map - usage analytics receiver.
 *
 * Setup:
 * 1. Create a Google Sheet for website usage logs.
 * 2. Open Extensions > Apps Script and paste this file.
 * 3. Set SPREADSHEET_ID below.
 * 4. Deploy > New deployment > Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web app URL into index.html:
 *    window.LUDM_ANALYTICS_ENDPOINT = 'YOUR_WEB_APP_URL';
 */

const SPREADSHEET_ID = 'PASTE_GOOGLE_SHEET_ID_HERE';
const RAW_SHEET_NAME = 'Raw Events';
const DAILY_SHEET_NAME = 'Daily Summary';

const RAW_HEADERS = [
  'receivedAt',
  'event',
  'timestamp',
  'sessionId',
  'page',
  'query',
  'selectedSchools',
  'selectedCount',
  'visibleCount',
  'details',
  'language',
  'viewport',
  'referrer',
  'userAgent',
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const sheet = getOrCreateSheet_(RAW_SHEET_NAME, RAW_HEADERS);
    sheet.appendRow([
      new Date(),
      payload.event || '',
      payload.timestamp || '',
      payload.sessionId || '',
      payload.page || '',
      payload.query || '',
      payload.selectedSchools || '',
      payload.selectedCount ?? '',
      payload.visibleCount ?? '',
      JSON.stringify(payload.details || {}),
      payload.language || '',
      payload.viewport || '',
      payload.referrer || '',
      payload.userAgent || '',
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createDailySummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const raw = ss.getSheetByName(RAW_SHEET_NAME);
  if (!raw || raw.getLastRow() < 2) return;

  const rows = raw.getRange(2, 1, raw.getLastRow() - 1, RAW_HEADERS.length).getValues();
  const eventIndex = RAW_HEADERS.indexOf('event');
  const receivedIndex = RAW_HEADERS.indexOf('receivedAt');
  const sessionIndex = RAW_HEADERS.indexOf('sessionId');
  const selectedSchoolsIndex = RAW_HEADERS.indexOf('selectedSchools');
  const detailsIndex = RAW_HEADERS.indexOf('details');

  const byDate = new Map();
  rows.forEach(row => {
    const received = row[receivedIndex];
    if (!(received instanceof Date)) return;
    const dateKey = Utilities.formatDate(received, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!byDate.has(dateKey)) {
      byDate.set(dateKey, {
        date: dateKey,
        sessions: new Set(),
        pageViews: 0,
        detailOpens: 0,
        selections: 0,
        exportsStarted: 0,
        exportsConfirmed: 0,
        clientRequirementUses: 0,
        schoolClicks: {},
        developmentClicks: {},
      });
    }

    const item = byDate.get(dateKey);
    const eventName = row[eventIndex];
    if (row[sessionIndex]) item.sessions.add(row[sessionIndex]);
    if (eventName === 'page_view') item.pageViews += 1;
    if (eventName === 'development_detail_opened') item.detailOpens += 1;
    if (eventName === 'development_selected') item.selections += 1;
    if (eventName === 'pdf_export_started') item.exportsStarted += 1;
    if (eventName === 'pdf_export_confirmed') item.exportsConfirmed += 1;
    if (eventName === 'client_requirements_applied') item.clientRequirementUses += 1;

    String(row[selectedSchoolsIndex] || '').split(',').filter(Boolean).forEach(school => {
      item.schoolClicks[school] = (item.schoolClicks[school] || 0) + 1;
    });

    try {
      const details = JSON.parse(row[detailsIndex] || '{}');
      if (details.development) {
        item.developmentClicks[details.development] = (item.developmentClicks[details.development] || 0) + 1;
      }
    } catch (_) {}
  });

  const summaryHeaders = [
    'date',
    'uniqueSessions',
    'pageViews',
    'detailOpens',
    'developmentSelections',
    'exportsStarted',
    'exportsConfirmed',
    'clientRequirementUses',
    'topSchools',
    'topDevelopments',
  ];
  const summary = getOrCreateSheet_(DAILY_SHEET_NAME, summaryHeaders);
  summary.clearContents();
  summary.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);

  const output = [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date)).map(item => [
    item.date,
    item.sessions.size,
    item.pageViews,
    item.detailOpens,
    item.selections,
    item.exportsStarted,
    item.exportsConfirmed,
    item.clientRequirementUses,
    topItems_(item.schoolClicks),
    topItems_(item.developmentClicks),
  ]);

  if (output.length) summary.getRange(2, 1, output.length, summaryHeaders.length).setValues(output);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function topItems_(counts) {
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => `${name} (${count})`)
    .join(', ');
}
