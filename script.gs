
/**
 * UNIFIED INDIA MART + WHATSAPP CRM + CONTACT SYNC
 * Flow:
 * 1) Pull IndiaMART leads into Customer sheet
 * 2) Reuse the same processing logic as onEdit (categorize, naming, group-200 logic)
 * 3) Append to Processed_data
 * 4) Sync to Google Contacts (background/manual)
 * 5) Export WhatsApp CSV by category/group
 */

// ==========================================
// 1. GLOBAL SETTINGS
// ==========================================

const S_CUSTOMER = 'customer';
const S_PROCESSED = 'Processed_data';
const S_CONFIG = 'Config';
const S_LOGS = 'Logs';

const TRIGGER_HANDLE_ON_EDIT = 'handleOnEdit';
const TRIGGER_BACKGROUND_SYNC = 'processBackgroundSync';
const TRIGGER_IMPORT = 'fetchIndiaMartBatch';

const IMPORT_PROP_START_DATE = 'START_DATE';
const IMPORT_PROP_END_DATE = 'END_DATE';
const IMPORT_PROP_OVERALL_START_DATE = 'OVERALL_START_DATE';
const IMPORT_PROP_OVERALL_END_DATE = 'OVERALL_END_DATE';
const IMPORT_PROP_PAGE = 'PAGE';
const IMPORT_PROP_TOTAL_IMPORTED = 'TOTAL_IMPORTED';
const IMPORT_PROP_STAGNANT_PAGE_COUNT = 'STAGNANT_PAGE_COUNT';
const IMPORT_PROP_LAST_PAGE_FINGERPRINT = 'LAST_PAGE_FINGERPRINT';
const IMPORT_PROP_BG_SYNC_QUEUED = 'BG_SYNC_QUEUED';
const IMPORT_PROP_STOP_REQUESTED = 'IMPORT_STOP_REQUESTED';
const IMPORT_PROP_LAST_STATUS = 'IMPORT_LAST_STATUS';
const IMPORT_PROP_EMPTY_RETRY_COUNT = 'EMPTY_RETRY_COUNT';
const IMPORT_PROP_LAST_API_HIT_MS = 'LAST_API_HIT_MS';
const IMPORT_PROP_INTERVAL_DAYS = 'IMPORT_INTERVAL_DAYS';
const IMPORT_PROP_LAST_RUN_DATE = 'IMPORT_LAST_RUN_DATE';
const IMPORT_PROP_RANGE_END_DATE = 'IMPORT_RANGE_END_DATE';
const IMPORT_PROP_IS_RECURRING = 'IMPORT_IS_RECURRING';

// 🛑 [USER CONFIGURATION REQUIRED] 🛑
// Replace this with your actual IndiaMART CRM Key
const INDIA_MART_API_KEY = 'YOUR_INDIAMART_API_KEY_HERE'; 

const IMPORT_EMPTY_MAX_RETRIES = 10;
const IMPORT_RATE_LIMIT_WINDOW_MS = (5 * 60 * 1000) + 15000; // 5m + safety buffer

// ==========================================
// 2. MENU + TRIGGERS
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('IndiaMART + WhatsApp CRM')
    .addItem('📥 Open Lead Puller', 'showIndiaMartSidebar')
    .addItem('🚀 Start Import (Saved Dates)', 'startLeadImport')
    .addItem('🗓️ Setup Recurring Import', 'showRecurringImportSetup')
    .addItem('📊 View Recurring Status', 'showRecurringStatus')
    .addItem('⏹️ Stop Import', 'stopIndiaMartImport')
    .addItem('🧯 Emergency Stop (Clear Triggers)', 'emergencyStopAllAutomation')
    .addSeparator()
    .addItem('📁 Export Selected to WhatsApp CSVs', 'showCsvPopup')
    .addItem('👥 Manual Sync to Google Contacts', 'showSyncPopup')
    .addSeparator()
    .addItem('⚙️ Initialize CRM Automation', 'setupTriggers')
    .addItem('🧪 Debug: Test Recurring Setup', 'testRecurringSetup')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  clearTriggersByHandlers_([TRIGGER_HANDLE_ON_EDIT, TRIGGER_BACKGROUND_SYNC]);

  ScriptApp.newTrigger(TRIGGER_HANDLE_ON_EDIT)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger(TRIGGER_BACKGROUND_SYNC)
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Automation Active!\n\n1. Customer edits/import are auto-processed.\n2. Contact sync runs every 5 minutes.'
  );
}

function clearTriggersByHandlers_(handlerNames) {
  const allowed = {};
  for (let i = 0; i < handlerNames.length; i++) {
    allowed[handlerNames[i]] = true;
  }

  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (allowed[fn]) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function ensureBackgroundSyncTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_BACKGROUND_SYNC) return;
  }
  ScriptApp.newTrigger(TRIGGER_BACKGROUND_SYNC)
    .timeBased()
    .everyMinutes(5)
    .create();
}

function queueBackgroundSyncSoon_(delayMs) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(IMPORT_PROP_BG_SYNC_QUEUED) === '1') return;

  const safeDelay = Math.max(5000, parseInt(delayMs, 10) || 15000);
  ScriptApp.newTrigger(TRIGGER_BACKGROUND_SYNC)
    .timeBased()
    .after(safeDelay)
    .create();
  props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '1');
}

// ==========================================
// 3. CUSTOMER -> PROCESSED ENGINE (onEdit + Import)
// ==========================================

function handleOnEdit(e) {
  if (!e || !e.range || !e.source) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (!isCustomerSheet_(sheet)) return;
  if (range.getRow() === 1) return;

  const editedStartCol = range.getColumn();
  const editedEndCol = range.getLastColumn();
  if (editedStartCol > 5 || editedEndCol < 1) return;

  processCustomerRows_(sheet, range.getRow(), range.getNumRows(), { source: 'onEdit' });
}

function processCustomerRows_(customerSheet, startRow, numRows, options) {
  if (!customerSheet || numRows <= 0) return 0;
  const opts = options || {};
  const shouldSyncNow = opts.syncNow !== false;
  const shouldExportCsvNow = opts.exportCsvNow !== false;
  const maxImmediateSyncRows = parseInt(opts.maxImmediateSyncRows, 10) || 0;
  const startedMs = new Date().getTime();
  const maxRunMs = parseInt(opts.maxRunMs, 10) || 210000; 
  let stoppedBySafety = false;

  const ss = customerSheet.getParent();
  const procSheet = getOrCreateSheetCaseInsensitive_(ss, S_PROCESSED);
  const configSheet = getSheetCaseInsensitive_(ss, S_CONFIG);

  ensureProcessedHeader_(procSheet);

  const rules = getCategoryRules_(configSheet);
  const existingPhones = getExistingProcessedPhoneSet_(procSheet);
  const fallbackCounters = {};

  const lastDataRow = customerSheet.getLastRow();
  if (startRow > lastDataRow) return 0;
  const effectiveNumRows = Math.min(numRows, (lastDataRow - startRow + 1));
  const values = customerSheet.getRange(startRow, 1, effectiveNumRows, 5).getValues();
  const pendingRows = [];
  let missingCount = 0;
  let invalidPhoneCount = 0;
  let duplicateCount = 0;
  let excludedCount = 0;

  for (let i = 0; i < values.length; i++) {
    if (i > 0 && i % 50 === 0) {
      if (isImportStopRequested_()) {
        stoppedBySafety = true;
        writeLog_('WARN', 'Processing interrupted by stop request. Row offset: ' + i);
        break;
      }
      if (new Date().getTime() - startedMs > maxRunMs) {
        stoppedBySafety = true;
        writeLog_('WARN', 'Processing interrupted by safety timeout. Row offset: ' + i);
        break;
      }
    }

    const row = values[i];

    const custName = safeCellString_(row[0]);
    const city = safeCellString_(row[1]);
    const orderTime = row[2];
    const prodName = safeCellString_(row[3]);
    const rawPhone = safeCellString_(row[4]);

    if (!prodName || !rawPhone || prodName === '-' || rawPhone === '-') {
      missingCount++;
      continue;
    }

    const cleanPhone = normalizeIndianPhone_(rawPhone);
    if (!cleanPhone) {
      invalidPhoneCount++;
      continue;
    }

    if (existingPhones[cleanPhone]) {
      duplicateCount++;
      continue;
    }

    const categoryMatch = matchCategory_(prodName, rules);
    if (categoryMatch.category === 'EXCLUDED') {
      excludedCount++;
      continue;
    }

    const counter = getNextGroupCounter_(configSheet, categoryMatch, fallbackCounters);

    const categoryInitial = categoryMatch.initial || 'OTH';
    const safeName = (custName && custName !== '-') ? custName : 'NoName';
    const cleanName = sanitizeNameForBroadcast_(safeName);
    const broadcastName = categoryInitial + '_G' + counter.groupNum + '_C' + counter.recCount + '_' + cleanName;

    pendingRows.push([
      safeName,
      city,
      orderTime,
      prodName,
      toDisplayIndianPhone_(cleanPhone),
      categoryMatch.category,
      broadcastName,
      cleanPhone,
      'Pending'
    ]);

    existingPhones[cleanPhone] = true;
  }

  if (pendingRows.length > 0) {
    writeLog_('DEBUG', 'Processing ' + pendingRows.length + ' valid rows into Processed_data');
    const appendStartRow = procSheet.getLastRow() + 1;
    procSheet.getRange(appendStartRow, 1, pendingRows.length, 9).setValues(pendingRows);
    ensureBackgroundSyncTrigger_();

    if (shouldSyncNow) {
      try {
        syncProcessedRowsImmediate_(procSheet, appendStartRow, pendingRows.length, maxImmediateSyncRows);
      } catch (e) {
        writeLog_('ERROR', 'Immediate sync failed: ' + e.message);
      }
    }

    if (shouldExportCsvNow) {
      try {
        autoExportCsvForProcessedRows_(procSheet, appendStartRow, pendingRows.length);
      } catch (e) {
        writeLog_('ERROR', 'Auto CSV export failed: ' + e.message);
      }
    }
  }

  writeLog_(
    'DEBUG',
    'processCustomerRows summary: scanned=' + values.length +
    ', added=' + pendingRows.length +
    ', missingItemOrPhone=' + missingCount +
    ', invalidPhone=' + invalidPhoneCount +
    ', duplicatePhone=' + duplicateCount +
    ', excluded=' + excludedCount +
    ', interrupted=' + (stoppedBySafety ? 'yes' : 'no')
  );

  return pendingRows.length;
}

function isCustomerSheet_(sheet) {
  if (!sheet) return false;
  return safeCellString_(sheet.getName()).toLowerCase() === S_CUSTOMER;
}

function ensureProcessedHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow([
    'Customer_Name',
    'City',
    'Order_Time',
    'Product_Name',
    'Phone_Number',
    'category',
    'Broadcast_Name',
    'WhatsApp_Number',
    'Synced_to_Contacts'
  ]);
}

function getCategoryRules_(configSheet) {
  if (!configSheet || configSheet.getLastRow() < 2) return [];

  const data = configSheet.getDataRange().getValues();
  const rules = [];

  for (let i = 1; i < data.length; i++) {
    const catName = safeCellString_(data[i][0]);
    if (!catName) continue;

    const initialRaw = safeCellString_(data[i][1]);
    const initial = (initialRaw || catName.substring(0, 3)).toUpperCase().replace(/\s+/g, '');
    const keywordsRaw = safeCellString_(data[i][2]).toLowerCase();

    if (!keywordsRaw) continue;

    const keywords = keywordsRaw
      .split('|')
      .map(function(k) { return k.trim(); })
      .filter(function(k) { return !!k; });

    if (!keywords.length) continue;

    const escaped = keywords.map(function(k) {
      return k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    });

    const pattern = '\\b(' + escaped.join('|') + ')\\b';

    rules.push({
      category: catName,
      initial: initial,
      regex: new RegExp(pattern, 'i'),
      configRow: i + 1
    });
  }

  return rules;
}

function matchCategory_(productName, rules) {
  const defaultCategory = {
    category: 'Others',
    initial: 'OTH',
    configRow: -1
  };

  if (!productName || !rules || !rules.length) return defaultCategory;

  for (let i = 0; i < rules.length; i++) {
    if (rules[i].regex.test(productName)) {
      if (safeCellString_(rules[i].category).toLowerCase() === 'exclude') {
        return { category: 'EXCLUDED', initial: 'EXC', configRow: rules[i].configRow };
      }
      return {
        category: rules[i].category,
        initial: rules[i].initial,
        configRow: rules[i].configRow
      };
    }
  }

  return defaultCategory;
}

function getNextGroupCounter_(configSheet, categoryMatch, fallbackCounters) {
  if (configSheet && categoryMatch.configRow > 0) {
    let groupNum = parseInt(configSheet.getRange(categoryMatch.configRow, 4).getValue(), 10) || 1;
    let recCount = parseInt(configSheet.getRange(categoryMatch.configRow, 5).getValue(), 10) || 0;

    if (recCount >= 200) {
      groupNum++;
      recCount = 0;
    }

    recCount++;

    configSheet.getRange(categoryMatch.configRow, 4, 1, 2).setValues([[groupNum, recCount]]);

    return { groupNum: groupNum, recCount: recCount };
  }

  const key = categoryMatch.initial || 'OTH';
  if (!fallbackCounters[key]) {
    fallbackCounters[key] = { groupNum: 1, recCount: 0 };
  }

  if (fallbackCounters[key].recCount >= 200) {
    fallbackCounters[key].groupNum++;
    fallbackCounters[key].recCount = 0;
  }

  fallbackCounters[key].recCount++;

  return {
    groupNum: fallbackCounters[key].groupNum,
    recCount: fallbackCounters[key].recCount
  };
}

function getExistingProcessedPhoneSet_(procSheet) {
  const result = {};
  if (!procSheet || procSheet.getLastRow() < 2) return result;

  const rowCount = procSheet.getLastRow() - 1;
  const rawPhoneCol = procSheet.getRange(2, 5, rowCount, 1).getValues();
  const waPhoneCol = procSheet.getRange(2, 8, rowCount, 1).getValues();

  for (let i = 0; i < rowCount; i++) {
    const p1 = normalizeIndianPhone_(rawPhoneCol[i][0]);
    const p2 = normalizeIndianPhone_(waPhoneCol[i][0]);
    if (p1) result[p1] = true;
    if (p2) result[p2] = true;
  }

  return result;
}

function normalizeIndianPhone_(raw) {
  const digitsOnly = safeCellString_(raw).replace(/\D/g, '');
  if (!digitsOnly) return '';

  let digits = digitsOnly;

  if (digits.length === 12 && digits.indexOf('91') === 0) {
    digits = digits.substring(2);
  }

  if (digits.length > 10) {
    digits = digits.slice(-10);
  }

  if (digits.length !== 10) return '';

  return digits;
}

function toDisplayIndianPhone_(tenDigitPhone) {
  if (!tenDigitPhone) return '';
  return '+91' + tenDigitPhone;
}

function toE164Indian_(rawPhone) {
  const ten = normalizeIndianPhone_(rawPhone);
  if (!ten) return '';
  return '+91' + ten;
}

function sanitizeNameForBroadcast_(name) {
  const clean = safeCellString_(name).replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '');
  return clean || 'NoName';
}

function safeCellString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

// ==========================================
// 4. INDIA MART IMPORT (INTEGRATED)
// ==========================================

function showSidebar() {
  showIndiaMartSidebar();
}

function showIndiaMartSidebar() {
  const html = HtmlService
    .createHtmlOutput(`
      <div style="font-family: 'Segoe UI', Tahoma, sans-serif; padding:20px; background:linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); min-height:100vh; box-sizing:border-box; margin:0;">
        <div style="background:#ffffff; border-radius:16px; box-shadow:0 10px 25px -5px rgba(0,0,0,0.1); padding:24px; border:1px solid #f1f5f9;">
          <h2 style="margin:0 0 15px; color:#0f172a; font-size: 20px;">📥 Lead Puller</h2>
          
          <div style="font-size:13px; color:#475569; margin-bottom:20px; background:#f0f9ff; border-left:4px solid #0ea5e9; padding:12px; border-radius: 4px; line-height: 1.5;">
            Pull any date range. The script auto-splits into 7-day chunks and keeps going until complete.
          </div>

          <div style="margin-bottom: 15px;">
            <label style="display:block; margin:0 0 6px; color:#334155; font-weight:600; font-size:14px;">📅 From Date</label>
            <input type="date" id="startDate" style="width:100%; padding:12px; border:1px solid #cbd5e1; border-radius:8px; font-size:14px; box-sizing:border-box; outline:none; transition:border 0.2s;" onfocus="this.style.borderColor='#0ea5e9'" onblur="this.style.borderColor='#cbd5e1'">
          </div>

          <div style="margin-bottom: 20px;">
            <label style="display:block; margin:0 0 6px; color:#334155; font-weight:600; font-size:14px;">📅 To Date</label>
            <input type="date" id="endDate" style="width:100%; padding:12px; border:1px solid #cbd5e1; border-radius:8px; font-size:14px; box-sizing:border-box; outline:none; transition:border 0.2s;" onfocus="this.style.borderColor='#0ea5e9'" onblur="this.style.borderColor='#cbd5e1'">
          </div>

          <div style="display:flex; flex-direction: column; gap:10px;">
            <button id="startBtn" onclick="startImport()" style="width:100%; padding:14px; border:none; border-radius:10px; background:linear-gradient(135deg,#0ea5e9,#0284c7); color:#fff; font-size:15px; font-weight:bold; cursor:pointer; box-shadow: 0 4px 6px -1px rgba(14,165,233,0.2); transition:transform 0.1s;">
              🚀 Start Import
            </button>
            <button id="stopBtn" onclick="stopImport()" style="width:100%; padding:14px; border:none; border-radius:10px; background:linear-gradient(135deg,#ef4444,#dc2626); color:#fff; font-size:15px; font-weight:bold; cursor:pointer; box-shadow: 0 4px 6px -1px rgba(239,68,68,0.2); transition:transform 0.1s;">
              🛑 Stop
            </button>
          </div>

          <div id="status" style="margin-top:20px; padding:12px; border-radius:8px; background:#f1f5f9; color:#334155; font-size:13px; border: 1px solid #e2e8f0; text-align:center;">Loading status...</div>
        </div>

        <script>
          function setStatus(msg, isError) {
            const el = document.getElementById("status");
            el.textContent = msg || "";
            el.style.background = isError ? "#fef2f2" : "#f1f5f9";
            el.style.color = isError ? "#b91c1c" : "#334155";
            el.style.border = isError ? "1px solid #fecaca" : "1px solid #e2e8f0";
          }

          function setBusy(isBusy) {
            document.getElementById("startBtn").disabled = isBusy;
            document.getElementById("stopBtn").disabled = isBusy;
            document.getElementById("startBtn").style.opacity = isBusy ? "0.7" : "1";
            document.getElementById("stopBtn").style.opacity = isBusy ? "0.7" : "1";
          }

          google.script.run
            .withSuccessHandler(function(msg) { setStatus(msg, false); })
            .withFailureHandler(function(err) { setStatus("❌ " + (err && err.message ? err.message : String(err)), true); })
            .getImportStatusMessage();

          function startImport() {
            const rawStart = document.getElementById("startDate").value;
            const rawEnd = document.getElementById("endDate").value;
            if (!rawStart || !rawEnd) {
              setStatus("❌ Please select both dates.", true);
              return;
            }
            if (rawStart > rawEnd) {
              setStatus("❌ From Date cannot be after To Date.", true);
              return;
            }

            setBusy(true);
            setStatus("⏳ Starting import...", false);
            google.script.run
              .withSuccessHandler(function(msg) {
                setBusy(false);
                setStatus("✅ " + msg, false);
              })
              .withFailureHandler(function(err) {
                setBusy(false);
                setStatus("❌ " + (err && err.message ? err.message : String(err)), true);
              })
              .initializeLeadImport(rawStart, rawEnd);
          }

          function stopImport() {
            setBusy(true);
            setStatus("⏳ Stopping import...", false);
            google.script.run
              .withSuccessHandler(function(msg) {
                setBusy(false);
                setStatus("🛑 " + msg, false);
              })
              .withFailureHandler(function(err) {
                setBusy(false);
                setStatus("❌ " + (err && err.message ? err.message : String(err)), true);
              })
              .stopIndiaMartImportFromUi();
          }
        </script>
      </div>
    `)
    .setTitle('IndiaMART Lead Manager')
    .setWidth(340);

  SpreadsheetApp.getUi().showSidebar(html);
}

function initializeLeadImport(startDate, endDate) {
  stopIndiaMartImport(false);

  const startObj = parseImportDate_(startDate);
  const endObj = parseImportDate_(endDate);
  if (!startObj || !endObj) {
    throw new Error('Invalid date format. Use YYYY-MM-DD.');
  }
  if (startObj.getTime() > endObj.getTime()) {
    throw new Error('From Date cannot be after To Date.');
  }

  const props = PropertiesService.getScriptProperties();
  setImportChunkRange_(props, startObj, endObj);
  props.setProperty(IMPORT_PROP_PAGE, '1');
  props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
  props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT, '');
  props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  props.setProperty(IMPORT_PROP_IS_RECURRING, '0');
  props.setProperty(
    IMPORT_PROP_LAST_STATUS,
    'Import queued. Overall range: ' +
      props.getProperty(IMPORT_PROP_OVERALL_START_DATE) + ' to ' +
      props.getProperty(IMPORT_PROP_OVERALL_END_DATE) + '.'
  );

  createIndiaMartImportTrigger_(1000);

  return 'Import started in background. Large date ranges will be auto-split into 7-day chunks until fully completed.';
}

function startLeadImport() {
  const props = PropertiesService.getScriptProperties();
  const startDate = props.getProperty(IMPORT_PROP_START_DATE) || props.getProperty(IMPORT_PROP_OVERALL_START_DATE);
  const endDate = props.getProperty(IMPORT_PROP_END_DATE) || props.getProperty(IMPORT_PROP_OVERALL_END_DATE);

  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Please open "IndiaMART: Open Lead Puller" and choose date range first.');
    return;
  }

  const startObj = parseImportDate_(startDate);
  const endObj = parseImportDate_(props.getProperty(IMPORT_PROP_OVERALL_END_DATE) || endDate);
  if (!startObj || !endObj) {
    SpreadsheetApp.getUi().alert('Saved date range is invalid. Please set dates again from Lead Puller.');
    return;
  }

  stopIndiaMartImport(false);
  if (!props.getProperty(IMPORT_PROP_OVERALL_START_DATE) || !props.getProperty(IMPORT_PROP_OVERALL_END_DATE)) {
    setImportChunkRange_(props, startObj, endObj);
  }
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
  props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT, '');
  props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import restarted from saved dates.');
  createIndiaMartImportTrigger_(1000);
  SpreadsheetApp.getUi().alert('IndiaMART import restarted with saved date range.');
}

function testRecurringSetup() {
  try {
    SpreadsheetApp.getUi().alert('Testing recurring setup with 7 days...');
    startRecurringImportWithInterval(7);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Test Failed:\n\n' + e.message);
  }
}

function showRecurringImportSetup() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, sans-serif; padding: 20px; background: linear-gradient(150deg,#f0f9ff 0%,#e0f2fe 100%); margin: 0; min-height: 100vh;}
          .container { max-width: 420px; margin: 0 auto; background: #fff; border-radius: 16px; border: 1px solid #bae6fd; box-shadow: 0 10px 25px rgba(0,0,0,0.08); padding: 24px; }
          h2 { color: #0369a1; margin: 0 0 15px 0; font-size: 22px;}
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 8px; font-weight: 600; color: #334155; }
          input { width: 100%; padding: 12px; box-sizing: border-box; font-size: 15px; border: 1px solid #cbd5e1; border-radius: 8px; transition: border 0.3s;}
          input:focus { border: 1px solid #0ea5e9; outline: none; }
          button { width: 100%; padding: 14px; background: linear-gradient(135deg,#0ea5e9,#0284c7); color: white; border: none; cursor: pointer; font-weight: bold; border-radius: 10px; font-size: 16px; box-shadow: 0 4px 6px -1px rgba(14, 165, 233, 0.2); transition: transform 0.1s;}
          button:hover { filter: brightness(1.05); }
          button:active { transform: scale(0.98); }
          .info { background: #f8fafc; padding: 12px; border-left: 4px solid #0284c7; margin: 15px 0; font-size: 13px; color: #475569; border-radius: 6px; line-height: 1.5; }
          .status { margin-top: 15px; padding: 12px; border-radius: 8px; background: #f1f5f9; font-size: 14px; color: #1e293b; border: 1px solid #e2e8f0; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🗓️ Weekly Recurring Import</h2>
          
          <div class="info">
            <strong>How it works:</strong><br>
            • Automatically runs every N days<br>
            • Each run completes full selected window in 7-day chunks<br>
            • Uses retry + no-progress detection to avoid stuck loops
          </div>

          <div class="form-group">
            <label for="interval">⏱️ Import Interval (Days):</label>
            <input type="number" id="interval" min="1" max="365" value="7" placeholder="e.g., 7 for weekly">
          </div>

          <div class="info">
            Examples:<br>
            • 7 = Weekly<br>
            • 14 = Bi-weekly<br>
            • 30 = Monthly
          </div>

          <button id="startRecurringBtn" onclick="startRecurring()">✅ Start Recurring Import</button>
          <div id="status" class="status">Ready to configure recurring import.</div>
          
          <script>
            function setStatus(msg, isError) {
              const el = document.getElementById('status');
              el.textContent = msg || '';
              el.style.background = isError ? '#fef2f2' : '#f1f5f9';
              el.style.color = isError ? '#b91c1c' : '#1e293b';
              el.style.border = isError ? '1px solid #fecaca' : '1px solid #e2e8f0';
            }

            function setBusy(isBusy) {
              const btn = document.getElementById('startRecurringBtn');
              btn.disabled = isBusy;
              btn.style.opacity = isBusy ? '0.7' : '1';
            }

            function startRecurring() {
              const intervalInput = document.getElementById('interval').value;
              const interval = parseFloat(intervalInput);
              
              if (!intervalInput || isNaN(interval)) {
                setStatus('❌ Enter a valid number.', true);
                return;
              }
              if (interval < 1) {
                setStatus('❌ Minimum interval is 1 day.', true);
                return;
              }
              if (interval > 365) {
                setStatus('❌ Maximum interval is 365 days.', true);
                return;
              }
              if (!Number.isInteger(interval)) {
                setStatus('⚠️ Rounded ' + interval + ' to ' + Math.round(interval) + ' days.', false);
              }

              setBusy(true);
              setStatus('⏳ Configuring recurring import...', false);
              
              google.script.run
                .withSuccessHandler(function(msg) {
                  google.script.run.showColorfulAlert('Recurring Setup Successful!', msg);
                  google.script.host.close();
                })
                .withFailureHandler(function(err) {
                  setBusy(false);
                  setStatus('❌ Setup failed: ' + (err && err.message ? err.message : String(err)), true);
                })
                .startRecurringImportWithInterval(Math.round(interval));
            }
          </script>
        </div>
      </body>
    </html>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Setup Weekly Import');
}

function startRecurringImportWithInterval(intervalDays) {
  try {
    const validatedDays = parseInt(intervalDays, 10);
    if (!validatedDays || isNaN(validatedDays) || validatedDays < 1 || validatedDays > 365) {
      throw new Error('Invalid interval: Must be between 1 and 365 days. You entered: ' + intervalDays);
    }

    stopIndiaMartImport(false);

    const props = PropertiesService.getScriptProperties();
    const now = new Date();

    props.setProperty(IMPORT_PROP_INTERVAL_DAYS, String(validatedDays));
    props.setProperty(IMPORT_PROP_IS_RECURRING, '1');
    props.setProperty(IMPORT_PROP_LAST_RUN_DATE, formatDateForImport_(now));

    const overallEndObj = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const overallStartObj = new Date(now.getTime() - (validatedDays * 24 * 60 * 60 * 1000));
    setImportChunkRange_(props, overallStartObj, overallEndObj);
    const startDate = props.getProperty(IMPORT_PROP_START_DATE);
    const endDate = props.getProperty(IMPORT_PROP_END_DATE);

    props.setProperty(IMPORT_PROP_RANGE_END_DATE, props.getProperty(IMPORT_PROP_OVERALL_END_DATE));
    props.setProperty(IMPORT_PROP_PAGE, '1');
    props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
    props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
    props.setProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT, '');
    props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');
    props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
    props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
    props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Recurring import started. Loading ' + validatedDays + ' days of data.');

    writeLog_('INFO', 'RECURRING IMPORT STARTED: ' + validatedDays + ' day interval. Range: ' + startDate + ' to ' + endDate);

    createIndiaMartImportTrigger_(1000);
    return (
      'Your import has been fully automated and will scan ' + validatedDays + ' days backward. ' +
      'It will continuously pull leads without manual intervention.'
    );
  } catch (error) {
    writeLog_('ERROR', 'startRecurringImportWithInterval failed: ' + error.message);
    throw error;
  }
}

function formatDateForImport_(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

function parseImportDate_(dateText) {
  const raw = safeCellString_(dateText);
  if (!raw) return null;

  const ymd = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) {
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  }

  const dmy = raw.match(/^(\d{2})-([A-Za-z]{3})-(\d{4})$/);
  if (dmy) {
    const monthMap = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };
    const m = monthMap[dmy[2].toLowerCase()];
    if (m === undefined) return null;
    return new Date(Number(dmy[3]), m, Number(dmy[1]));
  }

  const parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function addDays_(dateObj, days) {
  return new Date(dateObj.getTime() + (days * 24 * 60 * 60 * 1000));
}

function minDate_(a, b) {
  return (a.getTime() <= b.getTime()) ? a : b;
}

function buildLeadPageFingerprint_(rows) {
  if (!rows || !rows.length) return '';
  const parts = [];
  const limit = Math.min(rows.length, 20);
  for (let i = 0; i < limit; i++) {
    const lead = rows[i] || {};
    parts.push(
      safeCellString_(lead.SENDER_MOBILE) + '|' +
      safeCellString_(lead.QUERY_TIME) + '|' +
      safeCellString_(lead.QUERY_PRODUCT_NAME)
    );
  }
  return String(rows.length) + '::' + parts.join('||');
}

function setImportChunkRange_(props, overallStartDateObj, overallEndDateObj) {
  const chunkStart = new Date(
    overallStartDateObj.getFullYear(),
    overallStartDateObj.getMonth(),
    overallStartDateObj.getDate()
  );
  const chunkEnd = minDate_(addDays_(chunkStart, 6), overallEndDateObj);

  props.setProperty(IMPORT_PROP_OVERALL_START_DATE, formatDateForImport_(overallStartDateObj));
  props.setProperty(IMPORT_PROP_OVERALL_END_DATE, formatDateForImport_(overallEndDateObj));
  props.setProperty(IMPORT_PROP_START_DATE, formatDateForImport_(chunkStart));
  props.setProperty(IMPORT_PROP_END_DATE, formatDateForImport_(chunkEnd));
  props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT, '');
}

function advanceToNextChunk_(props) {
  const currentEnd = parseImportDate_(props.getProperty(IMPORT_PROP_END_DATE));
  const overallEnd = parseImportDate_(props.getProperty(IMPORT_PROP_OVERALL_END_DATE));
  if (!currentEnd || !overallEnd) return false;

  const nextStart = addDays_(currentEnd, 1);
  if (nextStart.getTime() > overallEnd.getTime()) return false;

  const nextEnd = minDate_(addDays_(nextStart, 6), overallEnd);
  props.setProperty(IMPORT_PROP_START_DATE, formatDateForImport_(nextStart));
  props.setProperty(IMPORT_PROP_END_DATE, formatDateForImport_(nextEnd));
  props.setProperty(IMPORT_PROP_PAGE, '1');
  props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  return true;
}

function finalizeOverallImportRun_(props, totalImported, fallbackEndDateText) {
  const isRecurring = props.getProperty(IMPORT_PROP_IS_RECURRING) === '1';
  const intervalDays = parseInt(props.getProperty(IMPORT_PROP_INTERVAL_DAYS), 10) || 0;

  if (isRecurring && intervalDays > 0) {
    const completedOverallEnd = parseImportDate_(props.getProperty(IMPORT_PROP_OVERALL_END_DATE) || fallbackEndDateText);
    const nextStart = completedOverallEnd ? addDays_(completedOverallEnd, 1) : addDays_(new Date(), 1);
    const nextEnd = addDays_(nextStart, intervalDays);

    setImportChunkRange_(props, nextStart, nextEnd);
    props.setProperty(IMPORT_PROP_RANGE_END_DATE, formatDateForImport_(nextEnd));
    props.setProperty(IMPORT_PROP_PAGE, '1');
    props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
    props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
    props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
    props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');

    const delayMs = (intervalDays * 24 * 60 * 60 * 1000);
    const nextRangeMsg =
      'Current overall range complete (' + totalImported + ' leads imported). ' +
      'Next recurring run scheduled in ' + intervalDays + ' days. Next overall range: ' +
      formatDateForImport_(nextStart) + ' to ' + formatDateForImport_(nextEnd) +
      '. First chunk: ' + props.getProperty(IMPORT_PROP_START_DATE) + ' to ' + props.getProperty(IMPORT_PROP_END_DATE) + '.';

    writeLog_('INFO', nextRangeMsg);
    props.setProperty(IMPORT_PROP_LAST_STATUS, nextRangeMsg);
    createIndiaMartImportTrigger_(delayMs);
    writeLog_('INFO', 'RECURRING: Waiting ' + intervalDays + ' days before next batch.');
    return;
  }

  const completeMsg =
    'Import completed for full requested range (' +
    props.getProperty(IMPORT_PROP_OVERALL_START_DATE) + ' to ' +
    props.getProperty(IMPORT_PROP_OVERALL_END_DATE) + '). Total imported: ' +
    totalImported + '.';
  props.setProperty(IMPORT_PROP_LAST_STATUS, completeMsg);
  stopIndiaMartImport(false);
  writeLog_('INFO', completeMsg);
}

function showRecurringStatus() {
  const props = PropertiesService.getScriptProperties();
  const isRecurring = props.getProperty(IMPORT_PROP_IS_RECURRING) === '1';
  
  if (!isRecurring) {
    const emptyHtml = HtmlService.createHtmlOutput(`
      <div style="font-family:'Segoe UI',sans-serif; text-align:center; padding:25px; background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); height: 100%; box-sizing: border-box; border-radius: 12px;">
        <h2 style="color:#334155; margin-top:10px; font-size: 20px;">No Active Recurring Import</h2>
        <p style="color:#64748b; font-size:14px; margin-bottom: 25px;">You can set up a recurring import schedule from the custom menu.</p>
        <button onclick="google.script.host.close()" style="background: #e2e8f0; color:#334155; border:none; padding:12px 24px; border-radius:8px; font-weight:bold; font-size: 15px; cursor:pointer; width:100%; transition: background 0.2s;">Close Window</button>
      </div>
    `).setWidth(350).setHeight(220);
    SpreadsheetApp.getUi().showModelessDialog(emptyHtml, 'Status');
    return;
  }

  const intervalDays = props.getProperty(IMPORT_PROP_INTERVAL_DAYS);
  const startDate = props.getProperty(IMPORT_PROP_START_DATE);
  const endDate = props.getProperty(IMPORT_PROP_END_DATE);
  const totalImported = props.getProperty(IMPORT_PROP_TOTAL_IMPORTED);
  const lastStatus = props.getProperty(IMPORT_PROP_LAST_STATUS);

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, sans-serif; padding: 20px; background: linear-gradient(150deg,#f8fafc 0%,#e2e8f0 100%); margin: 0; min-height: 100vh; box-sizing:border-box;}
          .card { background: white; border-radius: 16px; padding: 24px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); }
          .stat-box { background: #f1f5f9; padding: 14px; border-radius: 10px; margin-bottom: 12px; border-left: 4px solid #3b82f6;}
          .stat-title { font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: 700; margin-bottom: 6px; letter-spacing: 0.5px;}
          .stat-value { font-size: 16px; color: #0f172a; font-weight: 600; }
          .status-text { background: #e0f2fe; color: #0369a1; padding: 14px; border-radius: 10px; font-size: 13px; margin-bottom: 20px; border: 1px solid #bae6fd; line-height: 1.5;}
          button { width: 100%; padding: 14px; border: none; cursor: pointer; font-weight: bold; border-radius: 10px; font-size: 15px; margin-bottom: 10px; transition: transform 0.1s, box-shadow 0.2s;}
          button:active { transform: scale(0.98); }
          .btn-close { background: #f1f5f9; color: #475569; }
          .btn-close:hover { background: #e2e8f0; }
          .btn-stop { background: linear-gradient(135deg,#ef4444,#dc2626); color: white; box-shadow: 0 4px 6px -1px rgba(239,68,68,0.2); }
          .btn-stop:hover { filter: brightness(1.05); }
        </style>
      </head>
      <body>
        <div class="card">
          <h2 style="margin-top:0; color:#1e293b; font-size:22px; margin-bottom:18px;">📊 Import Status</h2>
          
          <div class="stat-box">
            <div class="stat-title">Schedule Interval</div>
            <div class="stat-value">Every ${intervalDays} Days</div>
          </div>
          
          <div class="stat-box">
            <div class="stat-title">Current Batch Range</div>
            <div class="stat-value">${startDate} to ${endDate}</div>
          </div>

          <div class="stat-box">
            <div class="stat-title">Total Imported Leads</div>
            <div class="stat-value">${totalImported || '0'}</div>
          </div>

          <div class="status-text">
            <strong>System Status:</strong><br>
            ${lastStatus || 'Running smoothly in background...'}
          </div>

          <button class="btn-stop" id="stopBtn" onclick="stopImport()">🛑 Stop & Delete Trigger</button>
          <button class="btn-close" onclick="google.script.host.close()">Close Window</button>
        </div>

        <script>
          function stopImport() {
            const btn = document.getElementById('stopBtn');
            btn.innerText = 'Stopping...';
            btn.disabled = true;
            
            google.script.run
              .withSuccessHandler(function(msg) {
                google.script.run.showColorfulAlert('Trigger Stopped', msg);
                google.script.host.close();
              })
              .withFailureHandler(function(err) {
                alert('Error: ' + err.message);
                btn.innerText = '🛑 Stop & Delete Trigger';
                btn.disabled = false;
              })
              .stopIndiaMartImportFromUi();
          }
        </script>
      </body>
    </html>
  `).setWidth(400).setHeight(600);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Recurring Status');
}

function fetchIndiaMartBatch() {
  try {
    writeLog_('DEBUG', '=== FETCH START ===');
    
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1)) {
      writeLog_('WARN', 'Lock acquisition failed - another process running');
      return;
    }

    writeLog_('DEBUG', 'Lock acquired');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();

    const startDate = props.getProperty(IMPORT_PROP_START_DATE);
    const endDate = props.getProperty(IMPORT_PROP_END_DATE);
    const overallStartDate = props.getProperty(IMPORT_PROP_OVERALL_START_DATE) || startDate;
    const overallEndDate = props.getProperty(IMPORT_PROP_OVERALL_END_DATE) || endDate;
    let page = parseInt(props.getProperty(IMPORT_PROP_PAGE), 10) || 1;
    let totalImported = parseInt(props.getProperty(IMPORT_PROP_TOTAL_IMPORTED), 10) || 0;
    let emptyRetryCount = parseInt(props.getProperty(IMPORT_PROP_EMPTY_RETRY_COUNT), 10) || 0;
    const lastApiHitMs = parseInt(props.getProperty(IMPORT_PROP_LAST_API_HIT_MS), 10) || 0;
    const nowMs = new Date().getTime();

    writeLog_('DEBUG', 'Loaded props: page=' + page + ', total=' + totalImported + ', retries=' + emptyRetryCount);

    if (isImportStopRequested_()) {
      writeLog_('INFO', 'Import stop was requested. No further pages scheduled.');
      return;
    }

    if (!startDate || !endDate) {
      writeLog_('ERROR', 'Import aborted. Missing start/end date.');
      props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import aborted: missing dates.');
      stopIndiaMartImport(false);
      return;
    }
    if (!props.getProperty(IMPORT_PROP_OVERALL_START_DATE) || !props.getProperty(IMPORT_PROP_OVERALL_END_DATE)) {
      props.setProperty(IMPORT_PROP_OVERALL_START_DATE, overallStartDate);
      props.setProperty(IMPORT_PROP_OVERALL_END_DATE, overallEndDate);
    }

    if (lastApiHitMs > 0) {
      const elapsedMs = nowMs - lastApiHitMs;
      if (elapsedMs < IMPORT_RATE_LIMIT_WINDOW_MS) {
        const waitMs = IMPORT_RATE_LIMIT_WINDOW_MS - elapsedMs;
        const waitSec = Math.ceil(waitMs / 1000);
        const waitMsg = 'Waiting ' + waitSec + 's before next API hit to respect IndiaMART 5-minute limit.';
        writeLog_('INFO', waitMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, waitMsg);
        createIndiaMartImportTrigger_(waitMs);
        return;
      }
    }

    const customerSheet = getOrCreateSheetCaseInsensitive_(ss, S_CUSTOMER);
    ensureCustomerHeader_(customerSheet);

    writeLog_('INFO', 'Starting IndiaMART page ' + page + ' | Range: ' + startDate + ' to ' + endDate);

    const url =
      'https://mapi.indiamart.com/wservce/crm/crmListing/v2/' +
      '?glusr_crm_key=' + encodeURIComponent(INDIA_MART_API_KEY) +
      '&start_time=' + encodeURIComponent(startDate) +
      '&end_time=' + encodeURIComponent(endDate) +
      '&page=' + encodeURIComponent(page);

    const options = {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json'
      }
    };

    writeLog_('DEBUG', 'Fetching URL...');
    props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, String(new Date().getTime()));
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const text = response.getContentText();

    writeLog_('API', 'HTTP ' + responseCode + ' | Page ' + page + ' | Bytes: ' + text.length);

    if (responseCode !== 200) {
      writeLog_('ERROR', 'HTTP ' + responseCode + ' - ' + text.substring(0, 500));
      if (responseCode === 429) {
        const extMsg = 'Rate-limit hit (HTTP 429). Another process may be using the same CRM key. Waiting before retry.';
        writeLog_('INFO', extMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, extMsg);
      }
      const delayMs = responseCode === 429 ? IMPORT_RATE_LIMIT_WINDOW_MS : 30000;
      createIndiaMartImportTrigger_(delayMs);
      return;
    }

    writeLog_('DEBUG', 'Parsing JSON...');
    const json = JSON.parse(text);
    const responseLength = json.RESPONSE ? json.RESPONSE.length : 0;
    const currentFingerprint = buildLeadPageFingerprint_(json.RESPONSE || []);
    const previousFingerprint = props.getProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT) || '';

    writeLog_('DEBUG', 'Response rows: ' + responseLength);

    if (!json.RESPONSE || json.RESPONSE.length === 0) {
      const apiMessage = safeCellString_(json.MESSAGE || json.message || json.STATUS || '');
      const rateLimitDetected = isRateLimitMessage_(apiMessage);
      const rangeLimitDetected = isDateRangeLimitMessage_(apiMessage);

      writeLog_('DEBUG', 'Empty response. Retries: ' + emptyRetryCount + '/' + IMPORT_EMPTY_MAX_RETRIES);

      if (rangeLimitDetected) {
        const sDate = parseImportDate_(startDate);
        const eDate = parseImportDate_(endDate);
        if (sDate && eDate && sDate.getTime() < eDate.getTime()) {
          const fixedEnd = minDate_(addDays_(sDate, 6), eDate);
          props.setProperty(IMPORT_PROP_END_DATE, formatDateForImport_(fixedEnd));
          props.setProperty(IMPORT_PROP_PAGE, '1');
          props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
          const fixMsg = 'Range auto-corrected to 7-day chunk: ' + startDate + ' to ' + formatDateForImport_(fixedEnd) + '. Retrying.';
          writeLog_('WARN', fixMsg + ' API message: ' + apiMessage);
          props.setProperty(IMPORT_PROP_LAST_STATUS, fixMsg);
          createIndiaMartImportTrigger_(1000);
          return;
        }
      }

      if (rateLimitDetected && (lastApiHitMs === 0 || (nowMs - lastApiHitMs) > IMPORT_RATE_LIMIT_WINDOW_MS)) {
        writeLog_(
          'INFO',
          'Rate-limit appears to come from another script/integration using the same CRM key. This run will wait and retry.'
        );
      }

      if (emptyRetryCount < IMPORT_EMPTY_MAX_RETRIES) {
        emptyRetryCount++;
        props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, String(emptyRetryCount));

        const delayMs = rateLimitDetected
          ? getRateLimitRetryDelayMs_(apiMessage)
          : Math.min(120000, 15000 * emptyRetryCount);
        const delaySec = Math.round(delayMs / 1000);
        const retryMsg =
          'API returned no RESPONSE rows (page ' + page + '). ' +
          'Retrying ' + emptyRetryCount + '/' + IMPORT_EMPTY_MAX_RETRIES + ' in ' + delaySec + 's.' +
          (apiMessage ? ' API message: ' + apiMessage : '');

        writeLog_('INFO', retryMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, retryMsg);
        createIndiaMartImportTrigger_(delayMs);
        return;
      }

      writeLog_('DONE', 'Chunk completed after retries. Total leads imported so far: ' + totalImported);
      writeLog_(
        'INFO',
        'API returned no RESPONSE rows for range ' + startDate + ' to ' + endDate + ' (page ' + page + ')' +
        (apiMessage ? '. API message: ' + apiMessage : '.') +
        ' Retries exhausted.'
      );

      if (advanceToNextChunk_(props)) {
        const nextChunkMsg =
          'Current chunk completed. Continuing next chunk: ' +
          props.getProperty(IMPORT_PROP_START_DATE) + ' to ' + props.getProperty(IMPORT_PROP_END_DATE) +
          ' (overall target: ' + props.getProperty(IMPORT_PROP_OVERALL_START_DATE) + ' to ' + props.getProperty(IMPORT_PROP_OVERALL_END_DATE) + ').';
        writeLog_('INFO', nextChunkMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, nextChunkMsg);
        createIndiaMartImportTrigger_(1000);
        return;
      }
      
      finalizeOverallImportRun_(props, totalImported, endDate);
      return;
    }

    if (page > 1 && currentFingerprint && previousFingerprint && currentFingerprint === previousFingerprint) {
      writeLog_(
        'WARN',
        'Detected same payload on page ' + page + ' for chunk ' + startDate + ' to ' + endDate +
        '. Skipping repeated page and advancing chunk.'
      );
      if (advanceToNextChunk_(props)) {
        const nextChunkMsg =
          'Auto-advanced chunk after repeated payload: ' +
          props.getProperty(IMPORT_PROP_START_DATE) + ' to ' + props.getProperty(IMPORT_PROP_END_DATE) + '.';
        props.setProperty(IMPORT_PROP_LAST_STATUS, nextChunkMsg);
        writeLog_('INFO', nextChunkMsg);
        createIndiaMartImportTrigger_(1000);
        return;
      }
      finalizeOverallImportRun_(props, totalImported, endDate);
      return;
    }

    if (emptyRetryCount !== 0) {
      props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
    }

    let insertedCount = 0;

    writeLog_('DEBUG', 'Inserting ' + json.RESPONSE.length + ' records (batch mode)...');

    const rowsToInsert = [];
    for (let i = 0; i < json.RESPONSE.length; i++) {
      if (isImportStopRequested_()) {
        writeLog_('INFO', 'Import stopped by user during batch insert. Imported so far: ' + totalImported);
        props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import stopped by user at page ' + page + '. Total leads imported: ' + totalImported + '.');
        stopIndiaMartImport(false);
        return;
      }

      const lead = json.RESPONSE[i];
      const row = [
        safeCellString_(lead.SENDER_NAME),
        safeCellString_(lead.SENDER_CITY),
        safeCellString_(lead.QUERY_TIME),
        safeCellString_(lead.QUERY_PRODUCT_NAME),
        safeCellString_(lead.SENDER_MOBILE)
      ];
      rowsToInsert.push(row);
      insertedCount++;
      totalImported++;
    }

    let processedCount = 0;

    if (rowsToInsert.length > 0) {
      const firstInsertRow = customerSheet.getLastRow() + 1;
      customerSheet.getRange(firstInsertRow, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
      writeLog_('DEBUG', 'Batch inserted ' + rowsToInsert.length + ' rows into Customer sheet');

      processedCount = processCustomerRows_(
        customerSheet,
        firstInsertRow,
        rowsToInsert.length,
        { syncNow: true, exportCsvNow: true, maxImmediateSyncRows: 120 }
      );
      writeLog_(
        'SUCCESS',
        'Processed ' + processedCount + ' valid rows into Processed_data (dedupe + validation applied).'
      );
    }

    let stagnantPages = parseInt(props.getProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT), 10) || 0;
    if (processedCount === 0) {
      stagnantPages++;
      props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, String(stagnantPages));
      writeLog_('WARN', 'No new valid rows on page ' + page + '. Stagnant page count: ' + stagnantPages + '.');
    } else {
      props.setProperty(IMPORT_PROP_STAGNANT_PAGE_COUNT, '0');
      stagnantPages = 0;
    }

    if (stagnantPages >= 2) {
      writeLog_(
        'WARN',
        'Detected repeated duplicate pages for chunk ' + startDate + ' to ' + endDate +
        '. Advancing to next chunk to avoid same-data loop.'
      );
      if (advanceToNextChunk_(props)) {
        const nextChunkMsg =
          'Advanced to next chunk after duplicate pages: ' +
          props.getProperty(IMPORT_PROP_START_DATE) + ' to ' + props.getProperty(IMPORT_PROP_END_DATE) + '.';
        props.setProperty(IMPORT_PROP_LAST_STATUS, nextChunkMsg);
        writeLog_('INFO', nextChunkMsg);
        createIndiaMartImportTrigger_(1000);
        return;
      }
      finalizeOverallImportRun_(props, totalImported, endDate);
      return;
    }

    props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, String(totalImported));
    props.setProperty(IMPORT_PROP_LAST_PAGE_FINGERPRINT, currentFingerprint);

    writeLog_('SUCCESS', 'Imported ' + insertedCount + ' leads (page ' + page + ')');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Running page ' + (page + 1) + '. Total imported: ' + totalImported + '.');
    props.setProperty(IMPORT_PROP_PAGE, String(page + 1));

    writeLog_('DEBUG', 'Creating trigger for next page (' + (page + 1) + ')...');
    createIndiaMartImportTrigger_(IMPORT_RATE_LIMIT_WINDOW_MS);
    writeLog_('DEBUG', 'Trigger created successfully');
  } catch (e) {
    writeLog_('ERROR', 'EXCEPTION in fetchIndiaMartBatch: ' + e.toString() + ' | Line: ' + e.stack);
    const props = PropertiesService.getScriptProperties();
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Error: ' + e.toString());
    createIndiaMartImportTrigger_(30000);
  } finally {
    try {
      const lock = LockService.getScriptLock();
      lock.releaseLock();
      writeLog_('DEBUG', 'Lock released');
    } catch (releaseErr) {
      writeLog_('ERROR', 'Lock release failed: ' + releaseErr.toString());
    }
  }
}

function ensureCustomerHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(['Customer_Name', 'City', 'Order_Time', 'Product_Name', 'Phone_Number']);
}

function createIndiaMartImportTrigger_(delayMs) {
  try {
    if (isImportStopRequested_()) {
      writeLog_('INFO', 'Not creating trigger - import stop requested');
      return;
    }
    
    const safeDelay = Math.max(1000, parseInt(delayMs, 10) || 1000);
    
    writeLog_('DEBUG', 'Clearing old triggers...');
    clearTriggersByHandlers_([TRIGGER_IMPORT]);
    
    writeLog_('DEBUG', 'Creating new trigger with delay: ' + safeDelay + 'ms');
    const trigger = ScriptApp.newTrigger(TRIGGER_IMPORT)
      .timeBased()
      .after(safeDelay)
      .create();
    
    writeLog_('DEBUG', 'Trigger created: ' + trigger.getUniqueId());
  } catch (e) {
    writeLog_('ERROR', 'FAILED TO CREATE TRIGGER: ' + e.toString());
  }
}

function stopIndiaMartImport(markStopped) {
  const props = PropertiesService.getScriptProperties();
  const shouldMarkStopped = markStopped !== false;

  if (shouldMarkStopped) {
    props.setProperty(IMPORT_PROP_STOP_REQUESTED, '1');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import stopped by user.');
    props.setProperty(IMPORT_PROP_IS_RECURRING, '0');
    props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');
  }

  clearTriggersByHandlers_([TRIGGER_IMPORT]);
}

function emergencyStopAllAutomation() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '1');
  props.setProperty(IMPORT_PROP_IS_RECURRING, '0');
  props.setProperty(IMPORT_PROP_LAST_STATUS, 'Emergency stop requested. All automation triggers cleared.');
  props.setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');

  clearTriggersByHandlers_([
    TRIGGER_IMPORT,
    TRIGGER_HANDLE_ON_EDIT,
    TRIGGER_BACKGROUND_SYNC
  ]);

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:'Segoe UI',sans-serif; text-align:center; padding:25px; background: linear-gradient(135deg, #fff1f2 0%, #fee2e2 100%); height: 100%; box-sizing: border-box; border-radius: 12px;">
      <div style="font-size:50px; margin-bottom:15px;">🛑</div>
      <h2 style="color:#991b1b; margin-top:0; font-size: 22px;">Automation Stopped</h2>
      <p style="color:#7f1d1d; font-size:14px; line-height:1.5; margin-bottom: 20px;">All background processes have been successfully deleted. Your CRM is now fully manual.</p>
      
      <div style="background:#fef2f2; border:1px solid #fecaca; padding:12px; border-radius:8px; font-size:13px; color:#991b1b; margin-bottom:20px; text-align:left;">
        <strong>Deleted Triggers:</strong><br><br>
        ✓ IndiaMART API Import Trigger<br>
        ✓ On-Edit Customer Processing<br>
        ✓ 5-Minute Background Contact Sync
      </div>
      
      <button onclick="google.script.host.close()" style="background: linear-gradient(135deg, #ef4444, #dc2626); color:white; border:none; padding:12px 24px; border-radius:8px; font-weight:bold; font-size: 15px; cursor:pointer; box-shadow: 0 4px 6px -1px rgba(239, 68, 68, 0.3); width:100%; transition: transform 0.1s;">Understood</button>
    </div>
  `).setWidth(380).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'System Halted');
}

function deleteTriggers() {
  stopIndiaMartImport();
}

function stopIndiaMartImportFromUi() {
  stopIndiaMartImport(true);
  return 'Import stopped and triggers cleared.';
}

function isImportStopRequested_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(IMPORT_PROP_STOP_REQUESTED) === '1';
}

function getImportStatusMessage() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(IMPORT_PROP_LAST_STATUS) || 'Idle.';
}

function isRateLimitMessage_(message) {
  const m = safeCellString_(message).toLowerCase();
  if (!m) return false;
  return m.indexOf('once in every 5 minutes') !== -1 ||
    (m.indexOf('crossed this limit') !== -1 && m.indexOf('try again after 5 minutes') !== -1);
}

function isDateRangeLimitMessage_(message) {
  const m = safeCellString_(message).toLowerCase();
  if (!m) return false;
  return m.indexOf('maximum allowed difference between start_time and end_time is 7 days') !== -1 ||
    (m.indexOf('maximum allowed difference') !== -1 && m.indexOf('7 days') !== -1);
}

function getRateLimitRetryDelayMs_(message) {
  const m = safeCellString_(message).toLowerCase();
  let minutes = 5;
  const match = m.match(/after\s+(\d+)\s+minute/);
  if (match && match[1]) {
    minutes = parseInt(match[1], 10) || 5;
  }

  return (minutes * 60 * 1000) + 15000;
}

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(S_LOGS);

  if (!logSheet) {
    logSheet = ss.insertSheet(S_LOGS);
    logSheet.appendRow(['Timestamp', 'Type', 'Message']);
  }

  return logSheet;
}

function writeLog_(type, message) {
  const logSheet = getLogSheet_();
  logSheet.appendRow([new Date(), type, message]);
}

// ==========================================
// 5. BACKGROUND SYNC ENGINE (UPGRADED FOR THREAD SAFETY)
// ==========================================

function processBackgroundSync() {
  PropertiesService.getScriptProperties().setProperty(IMPORT_PROP_BG_SYNC_QUEUED, '0');
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    writeLog_('WARN', 'Background sync skipped - locked by another process.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const procSheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
    if (!procSheet) return;

    const data = procSheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const startTime = new Date().getTime();
    const maxExecutionTime = 240000;

    const headers = data[0];
    const bNameIdx = headers.indexOf('Broadcast_Name');
    const phoneIdx = headers.indexOf('WhatsApp_Number');
    const syncIdx = headers.indexOf('Synced_to_Contacts');

    if (bNameIdx === -1 || phoneIdx === -1 || syncIdx === -1) return;

    let groupCache = {};
    try {
      const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
      const existingGroups = groupsResponse.contactGroups || [];

      for (let i = 0; i < existingGroups.length; i++) {
        groupCache[existingGroups[i].name] = existingGroups[i].resourceName;
      }
    } catch (e) {
      console.error('People API Error: ' + e.message);
      return;
    }

    for (let i = 1; i < data.length; i++) {
      if (new Date().getTime() - startTime > maxExecutionTime) {
        console.log('Time limit reached. Pausing until next cycle.');
        break;
      }

      const isSynced = safeCellString_(data[i][syncIdx]);
      if (isSynced !== 'Pending') continue;

      const broadcastName = safeCellString_(data[i][bNameIdx]);
      const phone = toE164Indian_(data[i][phoneIdx]);

      if (!broadcastName || !phone) continue;

      procSheet.getRange(i + 1, syncIdx + 1).setValue('Syncing...');
      SpreadsheetApp.flush();

      const labelName = getContactGroupNameFromBroadcast_(broadcastName);

      let groupId = groupCache[labelName];
      if (!groupId) {
        try {
          const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
          groupId = newGroup.resourceName;
          groupCache[labelName] = groupId;
        } catch (e) {
          procSheet.getRange(i + 1, syncIdx + 1).setValue('Error: ' + e.message);
          continue;
        }
      }

      const contact = {
        names: [{ givenName: broadcastName }],
        phoneNumbers: [{ value: phone, type: 'mobile' }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
      };

      let success = false;
      let lastError = '';

      for (let retry = 0; retry < 3; retry++) {
        try {
          People.People.createContact(contact);
          success = true;
          break;
        } catch (e) {
          lastError = e.message;
          const lowerErr = safeCellString_(lastError).toLowerCase();

          if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
            procSheet.getRange(i + 1, syncIdx + 1).setValue('API Error: ' + lastError);
            break;
          }

          Utilities.sleep(2000 * (retry + 1));
        }
      }

      if (success) {
        procSheet.getRange(i + 1, syncIdx + 1).setValue('Yes');
        Utilities.sleep(300);
      } else if (safeCellString_(procSheet.getRange(i + 1, syncIdx + 1).getValue()).indexOf('API Error') === -1) {
        procSheet.getRange(i + 1, syncIdx + 1).setValue('Pending');
        writeLog_('WARN', 'Background sync hit rate limit/temporary block.');
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function getContactGroupNameFromBroadcast_(broadcastName) {
  const parts = safeCellString_(broadcastName).split('_');
  if (parts.length >= 2) return parts[0] + '_' + parts[1];
  return 'Unknown_Group';
}

function syncProcessedRowsImmediate_(procSheet, startRow, numRows, maxRowsToSync) {
  if (!procSheet || numRows <= 0) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    writeLog_('WARN', 'Immediate sync skipped - lock active. Deferring to background.');
    queueBackgroundSyncSoon_(15000);
    return;
  }

  try {
    const headers = procSheet.getRange(1, 1, 1, procSheet.getLastColumn()).getValues()[0];
    const bNameIdx = headers.indexOf('Broadcast_Name');
    const phoneIdx = headers.indexOf('WhatsApp_Number');
    const syncIdx = headers.indexOf('Synced_to_Contacts');

    if (bNameIdx === -1 || phoneIdx === -1 || syncIdx === -1) return;

    const width = Math.max(bNameIdx, Math.max(phoneIdx, syncIdx)) + 1;
    const rows = procSheet.getRange(startRow, 1, numRows, width).getValues();

    const syncLimit = parseInt(maxRowsToSync, 10) || 0;
    let syncCount = 0;
    let deferredCount = 0;
    
    const updateRange = procSheet.getRange(startRow, syncIdx + 1, numRows, 1);
    const syncStatuses = updateRange.getValues();
    let markedAny = false;
    for (let i = 0; i < syncStatuses.length; i++) {
      if (syncLimit > 0 && syncCount >= syncLimit) {
        deferredCount++;
        continue;
      }
      if (syncStatuses[i][0] === 'Pending') {
        syncStatuses[i][0] = 'Syncing...';
        markedAny = true;
        syncCount++;
      }
    }
    if (markedAny) {
      updateRange.setValues(syncStatuses);
      SpreadsheetApp.flush();
    }

    const groupCache = getContactGroupCache_();
    const statusUpdates = [];
    const startedMs = new Date().getTime();
    const maxRunMs = 120000;

    for (let i = 0; i < rows.length; i++) {
      if (syncStatuses[i][0] !== 'Syncing...') continue;
      
      if (new Date().getTime() - startedMs > maxRunMs) {
        statusUpdates.push({ row: startRow + i, value: 'Pending' }); 
        deferredCount++;
        continue;
      }

      const broadcastName = safeCellString_(rows[i][bNameIdx]);
      const phone = toE164Indian_(rows[i][phoneIdx]);

      if (!broadcastName || !phone) continue;

      const labelName = getContactGroupNameFromBroadcast_(broadcastName);

      let groupId = groupCache[labelName];
      if (!groupId) {
        try {
          const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
          groupId = newGroup.resourceName;
          groupCache[labelName] = groupId;
        } catch (e) {
          statusUpdates.push({ row: startRow + i, value: 'Error: ' + e.message });
          continue;
        }
      }

      const contact = {
        names: [{ givenName: broadcastName }],
        phoneNumbers: [{ value: phone, type: 'mobile' }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
      };

      let success = false;
      let lastError = '';

      for (let retry = 0; retry < 3; retry++) {
        try {
          People.People.createContact(contact);
          success = true;
          break;
        } catch (e) {
          lastError = e.message;
          const lowerErr = safeCellString_(lastError).toLowerCase();
          if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
            statusUpdates.push({ row: startRow + i, value: 'API Error: ' + lastError });
            break;
          }
          Utilities.sleep(1200 * (retry + 1));
        }
      }

      if (success) {
        statusUpdates.push({ row: startRow + i, value: 'Yes' });
        Utilities.sleep(120);
      } else if (!statusUpdates.some(function(u) { return u.row === startRow + i; })) {
        statusUpdates.push({ row: startRow + i, value: 'Pending' }); 
        deferredCount++;
      }
    }

    for (let j = 0; j < statusUpdates.length; j++) {
      procSheet.getRange(statusUpdates[j].row, syncIdx + 1).setValue(statusUpdates[j].value);
    }

    if (deferredCount > 0) {
      writeLog_('INFO', 'Immediate contact sync deferred for ' + deferredCount + ' rows. Background sync will continue.');
      try {
        queueBackgroundSyncSoon_(15000);
      } catch (e) {
        writeLog_('ERROR', 'Failed to queue background sync: ' + e.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function getContactGroupCache_() {
  const cache = {};
  const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
  const existingGroups = groupsResponse.contactGroups || [];

  for (let i = 0; i < existingGroups.length; i++) {
    cache[existingGroups[i].name] = existingGroups[i].resourceName;
  }

  return cache;
}

function autoExportCsvForProcessedRows_(procSheet, startRow, numRows) {
  if (!procSheet || numRows <= 0) return;

  const headers = procSheet.getRange(1, 1, 1, procSheet.getLastColumn()).getValues()[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  if (bNameIdx === -1 || phoneIdx === -1) return;

  const width = Math.max(bNameIdx, phoneIdx) + 1;
  const newRows = procSheet.getRange(startRow, 1, numRows, width).getValues();
  const targetPrefixes = {};

  for (let i = 0; i < newRows.length; i++) {
    const bName = safeCellString_(newRows[i][bNameIdx]);
    if (!bName) continue;
    targetPrefixes[getContactGroupNameFromBroadcast_(bName)] = true;
  }

  const prefixes = Object.keys(targetPrefixes);
  if (!prefixes.length) return;

  const allData = procSheet.getDataRange().getValues();
  const csvByPrefix = {};

  for (let p = 0; p < prefixes.length; p++) {
    csvByPrefix[prefixes[p]] = ['Name,Phone'];
  }

  for (let r = 1; r < allData.length; r++) {
    const bName = safeCellString_(allData[r][bNameIdx]);
    if (!bName) continue;
    const prefix = getContactGroupNameFromBroadcast_(bName);
    if (!targetPrefixes[prefix]) continue;

    const phone = normalizeIndianPhone_(allData[r][phoneIdx]);
    if (!phone) continue;
    csvByPrefix[prefix].push('"' + bName + '","' + phone + '"');
  }

  const folder = getOrCreateFolder_('Whatsapp Category contact');

  for (let k = 0; k < prefixes.length; k++) {
    const prefix = prefixes[k];
    const fileName = prefix + '_WhatsApp.csv';
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
    folder.createFile(Utilities.newBlob(csvByPrefix[prefix].join('\n'), MimeType.CSV, fileName));
  }
}

// ==========================================
// 6. CSV EXPORT LOGIC
// ==========================================

function showCsvPopup() {
  const categories = getAvailableCategories_();
  if (!categories.length) {
    SpreadsheetApp.getUi().alert('No categories found. Add categories in Config sheet or Processed_data first.');
    return;
  }

  const checkboxesHtml = categories.map(function(cat) {
    return '<label class="cb-container"><input type="checkbox" value="' + cat + '" class="cat-checkbox"> ' + cat + '</label>';
  }).join('');

  const html = HtmlService
    .createHtmlOutput(getPopupHtml_('WhatsApp CSV Export', checkboxesHtml, 'runExport()', 'Generate CSV Files'))
    .setWidth(450)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'WhatsApp CSV Exporter');
}

function executeWhatsAppExport(selectedCategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  const catIdx = headers.indexOf('category');

  if (bNameIdx === -1 || phoneIdx === -1 || catIdx === -1) return;

  const selectedMap = toLookupMap_(selectedCategories);
  const csvGroups = {};

  for (let i = 1; i < data.length; i++) {
    const category = safeCellString_(data[i][catIdx]);
    if (!selectedMap[category]) continue;

    const bName = safeCellString_(data[i][bNameIdx]);
    const phone = normalizeIndianPhone_(data[i][phoneIdx]);
    if (!bName || !phone) continue;

    const filePrefix = getContactGroupNameFromBroadcast_(bName);
    if (!csvGroups[filePrefix]) {
      csvGroups[filePrefix] = ['Name,Phone'];
    }

    csvGroups[filePrefix].push('"' + bName + '","' + phone + '"');
  }

  const folder = getOrCreateFolder_('Whatsapp Category contact');
  let fileCount = 0;

  for (const prefix in csvGroups) {
    if (!Object.prototype.hasOwnProperty.call(csvGroups, prefix)) continue;

    const fileName = prefix + '_WhatsApp.csv';
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    folder.createFile(Utilities.newBlob(csvGroups[prefix].join('\n'), MimeType.CSV, fileName));
    fileCount++;
  }
  
  return fileCount;
}

// ==========================================
// 7. MANUAL CONTACT SYNC (THREAD SAFE)
// ==========================================

function showSyncPopup() {
  const categories = getAvailableCategories_();
  if (!categories.length) {
    SpreadsheetApp.getUi().alert('No categories found. Add categories in Config sheet or Processed_data first.');
    return;
  }

  const checkboxesHtml = categories.map(function(cat) {
    return '<label class="cb-container"><input type="checkbox" value="' + cat + '" class="cat-checkbox"> ' + cat + '</label>';
  }).join('');

  const html = HtmlService
    .createHtmlOutput(getPopupHtml_('Manual Contacts Sync', checkboxesHtml, 'startSync()', 'Sync Checked Categories'))
    .setWidth(450)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'Google Contacts Sync');
}

function executeContactSync(selectedCategories, startIndex) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Another sync operation is currently running. Please try again in a moment.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
    if (!sheet) throw new Error("Sheet 'Processed_data' not found.");

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'complete', syncedThisBatch: 0 };

    const headers = data[0];
    const bNameIdx = headers.indexOf('Broadcast_Name');
    const phoneIdx = headers.indexOf('WhatsApp_Number');
    const catIdx = headers.indexOf('category');
    const syncIdx = headers.indexOf('Synced_to_Contacts');

    if (bNameIdx === -1 || phoneIdx === -1 || catIdx === -1 || syncIdx === -1) {
      throw new Error('Processed_data sheet headers are missing required columns.');
    }

    const selectedMap = toLookupMap_(selectedCategories);

    let groupCache = {};
    try {
      const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
      const existingGroups = groupsResponse.contactGroups || [];
      for (let g = 0; g < existingGroups.length; g++) {
        groupCache[existingGroups[g].name] = existingGroups[g].resourceName;
      }
    } catch (e) {
      throw new Error('Could not load Contact Groups. Error: ' + e.message);
    }

    let syncedThisBatch = 0;
    const startTime = new Date().getTime();
    const start = Math.max(1, parseInt(startIndex, 10) || 1);

    for (let i = start; i < data.length; i++) {
      if (new Date().getTime() - startTime > 240000) {
        return { status: 'partial', nextRow: i, syncedThisBatch: syncedThisBatch };
      }

      const row = data[i];
      const category = safeCellString_(row[catIdx]);
      if (!selectedMap[category]) continue;

      const broadcastName = safeCellString_(row[bNameIdx]);
      const phone = toE164Indian_(row[phoneIdx]);
      const isSynced = safeCellString_(row[syncIdx]);

      if (isSynced !== 'Pending' || !broadcastName || !phone) continue;

      sheet.getRange(i + 1, syncIdx + 1).setValue('Syncing...');
      SpreadsheetApp.flush();

      const labelName = getContactGroupNameFromBroadcast_(broadcastName);

      let groupId = groupCache[labelName];
      if (!groupId) {
        const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
        groupId = newGroup.resourceName;
        groupCache[labelName] = groupId;
      }

      const newContact = {
        names: [{ givenName: broadcastName }],
        phoneNumbers: [{ value: phone, type: 'mobile' }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
      };

      let success = false;
      let lastError = '';

      for (let retry = 0; retry < 3; retry++) {
        try {
          People.People.createContact(newContact);
          success = true;
          break;
        } catch (e) {
          lastError = e.message;
          const lowerErr = safeCellString_(lastError).toLowerCase();

          if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
            sheet.getRange(i + 1, syncIdx + 1).setValue('API Error: ' + lastError);
            break;
          }

          Utilities.sleep(2000 * (retry + 1));
        }
      }

      if (success) {
        sheet.getRange(i + 1, syncIdx + 1).setValue('Yes');
        Utilities.sleep(500);
        syncedThisBatch++;
      } else if (safeCellString_(sheet.getRange(i + 1, syncIdx + 1).getValue()).indexOf('API Error') === -1) {
        sheet.getRange(i + 1, syncIdx + 1).setValue('Pending'); 
        throw new Error('Google blocked the connection. Error: ' + lastError);
      }
    }

    return { status: 'complete', syncedThisBatch: syncedThisBatch };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 8. SHARED HELPERS & UI (UPGRADED)
// ==========================================

function showColorfulAlert(title, message) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:'Segoe UI',sans-serif; text-align:center; padding:25px; background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); height: 100%; box-sizing: border-box; border-radius: 12px;">
      <div style="font-size:60px; margin-bottom:15px;">🎉</div>
      <h2 style="color:#166534; margin-top:0; font-size: 22px;">${title}</h2>
      <p style="color:#15803d; font-size:15px; line-height:1.6; margin-bottom: 25px;">${message}</p>
      <button onclick="google.script.host.close()" style="background: linear-gradient(135deg, #16a34a, #15803d); color:white; border:none; padding:12px 24px; border-radius:8px; font-weight:bold; font-size: 15px; cursor:pointer; box-shadow: 0 4px 6px -1px rgba(22, 163, 74, 0.3);">Awesome!</button>
    </div>
  `).setWidth(380).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Success');
}

function getAvailableCategories_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};

  const configSheet = getSheetCaseInsensitive_(ss, S_CONFIG);
  if (configSheet && configSheet.getLastRow() >= 2) {
    const cfg = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < cfg.length; i++) {
      const cat = safeCellString_(cfg[i][0]);
      if (!cat) continue;
      if (cat.toLowerCase() === 'exclude') continue;
      map[cat] = true;
    }
  }

  const processed = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (processed && processed.getLastRow() >= 2) {
    const headers = processed.getRange(1, 1, 1, processed.getLastColumn()).getValues()[0];
    const catIdx = headers.indexOf('category');

    if (catIdx !== -1) {
      const cats = processed.getRange(2, catIdx + 1, processed.getLastRow() - 1, 1).getValues();
      for (let j = 0; j < cats.length; j++) {
        const c = safeCellString_(cats[j][0]);
        if (!c) continue;
        if (c.toLowerCase() === 'excluded') continue;
        map[c] = true;
      }
    }
  }

  const out = Object.keys(map);
  out.sort();
  return out;
}

function toLookupMap_(arr) {
  const map = {};
  if (!arr || !arr.length) return map;

  for (let i = 0; i < arr.length; i++) {
    const value = safeCellString_(arr[i]);
    if (value) map[value] = true;
  }

  return map;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getSheetCaseInsensitive_(ss, name) {
  if (!ss || !name) return null;

  const target = safeCellString_(name).toLowerCase();
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (safeCellString_(sheets[i].getName()).toLowerCase() === target) {
      return sheets[i];
    }
  }

  return null;
}

function getOrCreateSheetCaseInsensitive_(ss, name) {
  let sheet = getSheetCaseInsensitive_(ss, name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  return sheet;
}

// ==========================================
// 9. POPUP HTML GENERATOR
// ==========================================

function getPopupHtml_(title, checkboxes, fnName, btnText) {
  return `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); padding: 20px; color: #1e293b; margin: 0; min-height: 100vh; box-sizing: border-box;}
        .card { background: white; border-radius: 16px; padding: 24px; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1); }
        h3 { margin-top: 0; color: #0f172a; font-size: 22px; border-bottom: 2px solid #f1f5f9; padding-bottom: 12px; margin-bottom: 20px; text-align: center; }
        .scroll-box { max-height: 320px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 10px; padding: 15px; margin-bottom: 20px; background: #f8fafc; }
        .cb-container { display: flex; align-items: center; margin-bottom: 12px; cursor: pointer; font-size: 15px; font-weight: 500; transition: color 0.2s; }
        .cb-container:hover { color: #3b82f6; }
        .cb-container input { margin-right: 12px; width: 18px; height: 18px; accent-color: #3b82f6; cursor: pointer; }
        button#runBtn { width: 100%; padding: 14px; background: linear-gradient(135deg, #3b82f6, #2563eb); color: white; border: none; cursor: pointer; font-weight: 700; border-radius: 10px; font-size: 16px; transition: transform 0.1s, box-shadow 0.2s; box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2); }
        button#runBtn:hover { box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3); }
        button#runBtn:active { transform: scale(0.98); }
        .btn-select { background: #e2e8f0; color: #475569; border: none; padding: 8px 14px; border-radius: 8px; cursor: pointer; font-size: 13px; margin-bottom: 15px; font-weight: 600; transition: background 0.2s; width: 100%; }
        .btn-select:hover { background: #cbd5e1; }
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
      </style>
    </head>
    <body>
      <div class="card">
        <h3>${title}</h3>
        <button class="btn-select" onclick="selectAll()">Toggle Select All</button>
        <div class="scroll-box">${checkboxes}</div>
        <button id="runBtn" onclick="${fnName}">${btnText}</button>
      </div>

      <script>
        let allChecked = false;
        let totalSynced = 0;

        function selectAll() {
          allChecked = !allChecked;
          document.querySelectorAll('.cat-checkbox').forEach(function(cb) { cb.checked = allChecked; });
        }

        function runExport() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(function(cb) { return cb.value; });
          if (cats.length === 0) return alert('Select at least one category.');
          
          const btn = document.getElementById('runBtn');
          btn.innerText = 'Processing...';
          btn.disabled = true;
          
          google.script.run
            .withSuccessHandler(function(count) {
              google.script.run.showColorfulAlert('Export Complete!', 'Generated ' + count + ' CSV file(s) in your Drive.');
              google.script.host.close();
            })
            .executeWhatsAppExport(cats);
        }

        function startSync() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(function(cb) { return cb.value; });
          if (cats.length === 0) { alert('Select at least one category.'); return; }

          const btn = document.getElementById('runBtn');
          btn.disabled = true;
          btn.style.background = 'linear-gradient(135deg, #94a3b8, #64748b)';
          totalSynced = 0;

          runSyncBatch(1, cats);
        }

        function runSyncBatch(startRowIndex, cats) {
          const btn = document.getElementById('runBtn');
          if (startRowIndex === 1) btn.innerText = 'Syncing... DO NOT close window';

          google.script.run
            .withSuccessHandler(function(response) {
              totalSynced += response.syncedThisBatch;

              if (response.status === 'partial') {
                btn.innerText = 'Synced ' + totalSynced + ' so far... Continuing';
                runSyncBatch(response.nextRow, cats);
              } else {
                google.script.run.showColorfulAlert('Sync Complete!', 'Successfully synced ' + totalSynced + ' total contacts.');
                google.script.host.close();
              }
            })
            .withFailureHandler(function(err) {
              alert('Error: ' + err.message);
              btn.innerText = 'Retry Sync';
              btn.disabled = false;
              btn.style.background = 'linear-gradient(135deg, #3b82f6, #2563eb)';
            })
            .executeContactSync(cats, startRowIndex);
        }
      </script>
    </body>
  </html>`;
}
