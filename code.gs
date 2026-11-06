/**
 * Expense Hermawan Family — London Expense Tracker
 * Database: Google Spreadsheet "Hermawan Fam's"
 *
 * FIRST STEP: run setupSpreadsheet() once from the Apps Script editor.
 * It creates the spreadsheet, the receipts folder in Drive and the monthly recap trigger.
 */

const SPREADSHEET_NAME = "Hermawan Fam's";
const SHEET_SETTINGS = 'Settings';
const SHEET_TX = 'Transactions';
const SHEET_RECAP = 'Monthly_Recap';
const TZ_LONDON = 'Europe/London';
const CATEGORIES = [
  'Housing & Utilities',
  'Groceries & Food',
  'Transport & Travel',
  'Academic & Misc',
  'Lifestyle & Personal',
  'Others & Miscellaneous'
];

/* ============================================================
 * SETUP (run once)
 * ============================================================ */

function setupSpreadsheet() {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty('SPREADSHEET_ID')) {
    const existing = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
    console.log('Spreadsheet already exists: ' + existing.getUrl());
    return existing.getUrl();
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  ss.setSpreadsheetTimeZone(TZ_LONDON);

  // ---- Settings tab ----
  const sSet = ss.getSheets()[0].setName(SHEET_SETTINGS);
  sSet.getRange('A1:C1').setValues([['Key', 'Value', 'Description']]);
  sSet.getRange('B5').setNumberFormat('@'); // keep the PIN as text so leading zeros survive
  sSet.getRange('A2:C6').setValues([
    ['BASE_BUDGET_GBP', 5000, 'Monthly base budget (GBP), resets on the 1st of every month'],
    ['EXCHANGE_RATE_MODE', 'AUTO', 'AUTO = live exchange rate, MANUAL = use FIXED_RATE_IDR'],
    ['FIXED_RATE_IDR', 24441, 'Manual GBP to IDR rate / fallback if the rate API fails'],
    ['APP_PIN', '123456', 'App login PIN (6 digits) - CHANGE THIS IMMEDIATELY'],
    ['NOTIFICATION_EMAILS', Session.getEffectiveUser().getEmail(), 'Separate multiple recipients with commas']
  ]);
  sSet.getRange('B2:B6').setHorizontalAlignment('left');
  sSet.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['AUTO', 'MANUAL'], true).build()
  );
  styleHeader_(sSet.getRange('A1:C1'));
  sSet.setColumnWidths(1, 1, 190);
  sSet.setColumnWidths(2, 1, 220);
  sSet.setColumnWidths(3, 1, 430);
  sSet.setFrozenRows(1);

  // ---- Transactions tab ----
  const sTx = ss.insertSheet(SHEET_TX);
  const txHeader = ['Transaction_ID', 'Time_London', 'Time_WIB', 'Time_WITA', 'Month_Year',
    'Category', 'Description', 'Amount_GBP', 'Rate_IDR', 'Amount_IDR', 'Receipt_Link'];
  sTx.getRange(1, 1, 1, txHeader.length).setValues([txHeader]);
  sTx.getRange('E:E').setNumberFormat('@'); // stop Sheets turning "2026-11" into a date
  sTx.getRange('H2:H').setNumberFormat('£#,##0.00');
  sTx.getRange('I2:I').setNumberFormat('#,##0.00');
  sTx.getRange('J2:J').setNumberFormat('"Rp" #,##0');
  styleHeader_(sTx.getRange(1, 1, 1, txHeader.length));
  sTx.setFrozenRows(1);
  sTx.setColumnWidths(1, 1, 150);
  sTx.setColumnWidths(2, 3, 190);
  sTx.setColumnWidths(6, 2, 200);

  // ---- Monthly recap tab ----
  const sRecap = ss.insertSheet(SHEET_RECAP);
  const recapHeader = ['Month', 'Base Budget', 'Total Spent', 'Remaining', 'Percentage', 'Status'];
  sRecap.getRange(1, 1, 1, recapHeader.length).setValues([recapHeader]);
  sRecap.getRange('A:A').setNumberFormat('@');
  sRecap.getRange('B2:D').setNumberFormat('£#,##0.00');
  sRecap.getRange('E2:E').setNumberFormat('0.0%');
  styleHeader_(sRecap.getRange(1, 1, 1, recapHeader.length));
  sRecap.setFrozenRows(1);
  sRecap.setColumnWidths(1, 6, 140);

  // ---- Receipts folder in Drive ----
  const folder = DriveApp.createFolder("Receipts Hermawan Fam's");

  props.setProperties({
    SPREADSHEET_ID: ss.getId(),
    FOLDER_STRUK_ID: folder.getId()
  });

  setupMonthlyTrigger();

  console.log('Spreadsheet created: ' + ss.getUrl());
  console.log('Receipts folder: ' + folder.getUrl());
  return ss.getUrl();
}

/** Monthly recap on the 1st at 07:00 (project time zone; set it to Asia/Jakarta for 07:00 WIB). */
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyRecap') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendMonthlyRecap').timeBased().onMonthDay(1).atHour(7).create();
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
}

/* ============================================================
 * HELPERS
 * ============================================================ */

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet not created yet. Run setupSpreadsheet() once from the Apps Script editor.');
  return SpreadsheetApp.openById(id);
}

function getSettings() {
  const data = getSpreadsheet_().getSheetByName(SHEET_SETTINGS).getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (key) settings[key] = data[i][1];
  }
  return settings;
}

/** Normalise the Month_Year column to 'yyyy-MM' (safe even if Sheets turned it into a date). */
function monthKey_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, TZ_LONDON, 'yyyy-MM');
  return String(value).trim().slice(0, 7);
}

function round2_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function fmtInt_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function parseEmails_(raw) {
  return String(raw || '').split(',').map(e => e.trim()).filter(e => e !== '');
}

function aggregateMonth_(data, monthKey) {
  const breakdown = {};
  CATEGORIES.forEach(c => breakdown[c] = 0);
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || monthKey_(row[4]) !== monthKey) continue;
    const amount = parseFloat(row[7]) || 0;
    total += amount;
    const cat = breakdown.hasOwnProperty(row[5]) ? row[5] : 'Others & Miscellaneous';
    breakdown[cat] += amount;
  }
  return { total: round2_(total), breakdown: breakdown };
}

/* ============================================================
 * WEB APP
 * ============================================================ */

function doGet() {
  // The HTML file in the Apps Script editor must be named exactly "Index"
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Expense Hermawan Family')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verifyPin(inputPin) {
  const savedPin = String(getSettings().APP_PIN || '123456').trim();
  if (String(inputPin).trim() === savedPin) {
    return { success: true, summary: getDashboardSummary() };
  }
  return { success: false };
}

function getExchangeRate(settingsOpt) {
  const settings = settingsOpt || getSettings();
  const manualRate = parseFloat(settings.FIXED_RATE_IDR) || 24441;

  if (String(settings.EXCHANGE_RATE_MODE).trim().toUpperCase() === 'MANUAL') return manualRate;

  const cache = CacheService.getScriptCache();
  const cached = cache.get('GBP_IDR');
  if (cached) return parseFloat(cached);

  try {
    const res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/GBP', { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const rate = parseFloat(JSON.parse(res.getContentText()).rates.IDR);
      if (rate > 0) {
        cache.put('GBP_IDR', String(rate), 3600); // cache for 1 hour
        return rate;
      }
    }
  } catch (e) {
    console.error('Failed to fetch exchange rate: ' + e);
  }
  return manualRate; // fallback
}

function getDashboardSummary() {
  const settings = getSettings();
  const baseBudget = parseFloat(settings.BASE_BUDGET_GBP) || 5000;
  const data = getSpreadsheet_().getSheetByName(SHEET_TX).getDataRange().getValues();

  const currentMonth = Utilities.formatDate(new Date(), TZ_LONDON, 'yyyy-MM');
  const month = aggregateMonth_(data, currentMonth);

  const recentTransactions = [];
  for (let i = data.length - 1; i >= 1 && recentTransactions.length < 5; i--) {
    const row = data[i];
    if (!row[0]) continue;
    recentTransactions.push({
      waktuLondon: String(row[1]),
      waktuWIB: String(row[2]),
      waktuWITA: String(row[3]),
      kategori: row[5],
      keterangan: row[6],
      nominalGBP: parseFloat(row[7]) || 0,
      nominalIDR: parseFloat(row[9]) || 0,
      linkStruk: row[10] || '-'
    });
  }

  return {
    baseBudget: baseBudget,
    totalSpent: month.total,
    remainingBudget: round2_(baseBudget - month.total),
    exchangeRate: getExchangeRate(settings),
    recentTransactions: recentTransactions,
    chartData: {
      labels: Object.keys(month.breakdown),
      values: Object.values(month.breakdown).map(round2_)
    }
  };
}

function processTransaction(formObject) {
  const settings = getSettings();
  const rate = getExchangeRate(settings);

  // ---- Validate input: two amounts (GBP and IDR), either one is enough ----
  let amountGBP = parseFloat(formObject.nominalGBP);
  let amountIDR = parseFloat(formObject.nominalIDR);
  if (!(amountGBP > 0) && amountIDR > 0) amountGBP = amountIDR / rate;
  if (!(amountGBP > 0)) throw new Error('Amount must be greater than 0.');
  amountGBP = round2_(amountGBP);

  // Keep the IDR amount the user typed if it matches the GBP amount (within 1%); otherwise recalculate
  const computedIDR = Math.round(amountGBP * rate);
  amountIDR = (amountIDR > 0 && Math.abs(amountIDR - computedIDR) <= computedIDR * 0.01)
    ? Math.round(amountIDR) : computedIDR;

  const description = String(formObject.keterangan || '').trim();
  if (!description) throw new Error('Description is required.');
  const category = CATEGORIES.indexOf(formObject.kategori) >= 0 ? formObject.kategori : 'Others & Miscellaneous';

  const now = new Date();
  const timeLondon = Utilities.formatDate(now, TZ_LONDON, "dd MMM yyyy, HH:mm z");
  const timeWIB = Utilities.formatDate(now, 'Asia/Jakarta', "dd MMM yyyy, HH:mm 'WIB'");
  const timeWITA = Utilities.formatDate(now, 'Asia/Makassar', "dd MMM yyyy, HH:mm 'WITA'");
  const monthKey = Utilities.formatDate(now, TZ_LONDON, 'yyyy-MM');
  const idTx = 'TX-' + now.getTime();

  // ---- Receipt upload (optional) ----
  let receiptLink = '-';
  const blob = formObject.fileStruk;
  const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_STRUK_ID');
  if (blob && typeof blob.getBytes === 'function' && blob.getBytes().length > 0 && folderId) {
    try {
      const file = DriveApp.getFolderById(folderId).createFile(blob.setName(idTx + '_' + blob.getName()));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      receiptLink = file.getUrl();
    } catch (e) {
      console.error('Receipt upload failed: ' + e);
    }
  }

  // ---- Save (locked so two people saving at once do not collide) ----
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    getSpreadsheet_().getSheetByName(SHEET_TX).appendRow([
      idTx, timeLondon, timeWIB, timeWITA, monthKey, category, description,
      amountGBP, rate, amountIDR, receiptLink
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // Compute the summary AFTER the transaction is saved (used by the email and the dashboard)
  const summary = getDashboardSummary();

  // ---- Instant email notification (an email failure never cancels the transaction) ----
  try {
    const recipients = parseEmails_(settings.NOTIFICATION_EMAILS);
    if (recipients.length > 0) {
      const subject = '[London Expense] New Expense: £' + amountGBP.toFixed(2) +
        ' (' + category + ') - Remaining: £' + summary.remainingBudget.toFixed(2);
      const body = `
        <div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px;">
          <h2 style="color: #4f46e5;">🇬🇧 New Expense Recorded</h2>
          <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <tr><td style="padding: 8px; font-weight: bold;">London Time:</td><td style="padding: 8px;">${escapeHtml_(timeLondon)}</td></tr>
            <tr style="background: #f8fafc;"><td style="padding: 8px; font-weight: bold;">WIB / WITA Time:</td><td style="padding: 8px;">${escapeHtml_(timeWIB)} / ${escapeHtml_(timeWITA)}</td></tr>
            <tr><td style="padding: 8px; font-weight: bold;">Category:</td><td style="padding: 8px;"><strong>${escapeHtml_(category)}</strong></td></tr>
            <tr style="background: #f8fafc;"><td style="padding: 8px; font-weight: bold;">Description:</td><td style="padding: 8px;">${escapeHtml_(description)}</td></tr>
            <tr><td style="padding: 8px; font-weight: bold;">Amount (GBP):</td><td style="padding: 8px; color: #dc2626; font-size: 16px;"><strong>£${amountGBP.toFixed(2)}</strong></td></tr>
            <tr style="background: #f8fafc;"><td style="padding: 8px; font-weight: bold;">Amount (IDR):</td><td style="padding: 8px; font-size: 16px;"><strong>Rp ${fmtInt_(amountIDR)}</strong> <span style="color:#64748b;font-size:12px;">(rate Rp ${fmtInt_(rate)} / £)</span></td></tr>
            <tr><td style="padding: 8px; font-weight: bold;">Receipt:</td><td style="padding: 8px;">${receiptLink !== '-' ? '<a href="' + receiptLink + '">View receipt</a>' : 'None'}</td></tr>
          </table>
          <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;">
          <p style="font-size: 15px;"><b>Remaining Budget This Month:</b> <span style="color: #16a34a; font-size: 18px; font-weight: bold;">£${summary.remainingBudget.toFixed(2)}</span></p>
        </div>`;
      MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: body });
    }
  } catch (err) {
    console.error('Failed to send email: ' + err);
  }

  return { status: 'success', summary: summary };
}

/* ============================================================
 * MONTHLY RECAP (trigger: 1st of the month, 07:00)
 * ============================================================ */

function sendMonthlyRecap() {
  const settings = getSettings();
  const baseBudget = parseFloat(settings.BASE_BUDGET_GBP) || 5000;
  const ss = getSpreadsheet_();

  // On the 1st, going back 2 days always lands in the previous month
  const prevMonth = Utilities.formatDate(new Date(Date.now() - 2 * 24 * 3600 * 1000), TZ_LONDON, 'yyyy-MM');
  const data = ss.getSheetByName(SHEET_TX).getDataRange().getValues();
  const month = aggregateMonth_(data, prevMonth);

  const remaining = round2_(baseBudget - month.total);
  const pct = baseBudget > 0 ? month.total / baseBudget : 0;
  const status = pct > 1 ? 'Over Budget' : (pct >= 0.9 ? 'Warning' : 'Safe');

  // Log to the Monthly_Recap tab (skip if this month is already there)
  const sRecap = ss.getSheetByName(SHEET_RECAP);
  const existing = sRecap.getDataRange().getValues().slice(1).map(r => monthKey_(r[0]));
  if (existing.indexOf(prevMonth) === -1) {
    sRecap.appendRow([prevMonth, baseBudget, month.total, remaining, pct, status]);
  }

  const recipients = parseEmails_(settings.NOTIFICATION_EMAILS);
  if (recipients.length === 0) return;

  const rows = Object.keys(month.breakdown).map(cat =>
    '<tr><td style="padding:6px 8px;">' + escapeHtml_(cat) + '</td>' +
    '<td style="padding:6px 8px;text-align:right;">£' + round2_(month.breakdown[cat]).toFixed(2) + '</td></tr>'
  ).join('');

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: '[London Expense] ' + prevMonth + ' Recap: £' + month.total.toFixed(2) + ' spent (' + status + ')',
    htmlBody: `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px;">
        <h2 style="color: #4f46e5;">📊 Spending Recap ${prevMonth}</h2>
        <p>Base budget: <b>£${baseBudget.toFixed(2)}</b><br>
           Total spent: <b style="color:#dc2626;">£${month.total.toFixed(2)}</b> (${(pct * 100).toFixed(1)}%)<br>
           Remaining: <b style="color:#16a34a;">£${remaining.toFixed(2)}</b><br>
           Status: <b>${status}</b></p>
        <table style="width:100%;border-collapse:collapse;">${rows}</table>
        <p style="font-size:12px;color:#64748b;">The budget resets fully on the 1st with no rollover.</p>
      </div>`
  });
}