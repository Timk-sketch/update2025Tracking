// =====================================================
// 02_Sheets.gs — Sheet helpers + logging + dedupe
// =====================================================

function logImportEvent(source, message, count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Log');
  if (!logSheet) {
    logSheet = ss.insertSheet('Log');
    logSheet.appendRow(['Timestamp', 'Source', 'Message', 'Count']);
  }
  logSheet.appendRow([
    new Date().toISOString(),
    source,
    message,
    count !== undefined ? count : ''
  ]);
}

/**
 * Logs progress with both toast notification and log entry.
 * Use this for long-running operations to show user what's happening.
 */
function logProgress(source, message, showToast) {
  showToast = showToast !== false; // Default true

  // Log to sheet
  logImportEvent(source, message);

  // Show toast if enabled
  if (showToast) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(message, source, 2);
  }

  // Also log to console for debugging
  console.log(`[${source}] ${message}`);
}

// IMPORTANT: expands sheet rows/cols BEFORE writing headers (fixes out-of-bounds)
function getOrCreateSheetWithHeaders(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const needCols = headers.length;

  // Expand columns if needed
  const maxCols = sheet.getMaxColumns();
  if (maxCols < needCols) {
    sheet.insertColumnsAfter(maxCols, needCols - maxCols);
  }

  // Ensure at least 1 row for header
  if (sheet.getMaxRows() < 1) {
    sheet.insertRowsAfter(1, 1);
  }

  const lastRow = sheet.getLastRow();

  // If empty, write headers
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, needCols).setValues([headers]);
    return sheet;
  }

  const lastCol = Math.max(sheet.getLastColumn(), needCols);

  // Expand again if a weird case left us short
  if (sheet.getMaxColumns() < lastCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
  }

  // Read existing header row up to lastCol
  const existingRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(x => String(x || '').trim());

  // Determine existing "header length" (last non-empty)
  let existingLen = 0;
  for (let i = existingRow.length - 1; i >= 0; i--) {
    if (existingRow[i]) { existingLen = i + 1; break; }
  }
  if (existingLen === 0) existingLen = Math.min(existingRow.length, needCols);

  const expected = headers.map(x => String(x || '').trim());
  const minLen = Math.min(existingLen, expected.length);

  const existingPrefix = existingRow.slice(0, minLen).join('||');
  const expectedPrefix = expected.slice(0, minLen).join('||');

  if (existingPrefix === expectedPrefix) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return sheet;
  }

  // Out of sync: archive old + create fresh
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const archivedName = `${sheetName}_ARCHIVE_${ts}`;
  sheet.setName(archivedName);

  const fresh = ss.insertSheet(sheetName);

  // expand columns on new sheet too
  if (fresh.getMaxColumns() < needCols) {
    fresh.insertColumnsAfter(fresh.getMaxColumns(), needCols - fresh.getMaxColumns());
  }
  fresh.getRange(1, 1, 1, expected.length).setValues([expected]);

  ss.toast(`⚠️ Headers changed for "${sheetName}". Old sheet archived as "${archivedName}".`, "Headers Updated", 10);
  return fresh;
}

function deduplicateSheet(sheetName, headers) {
  const sheet = getOrCreateSheetWithHeaders(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headerRow = data[0].map(h => String(h || '').trim());
  const idCol = headerRow.indexOf('Order ID');

  let keyCol, keyLabel;
  if (sheetName === 'Shopify Orders') {
    keyCol = headerRow.indexOf('Lineitem ID');
    keyLabel = 'Lineitem ID';
  } else {
    keyCol = headerRow.indexOf('LineItem ID');
    keyLabel = 'LineItem ID';
  }

  if (idCol === -1 || keyCol === -1) {
    throw new Error(`Deduplication failed: Could not find "Order ID" and "${keyLabel}" columns for ${sheetName}.`);
  }

  const seen = new Set();
  const rowsToKeep = [data[0]];
  let duplicatesFound = 0;

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][idCol] || "") + '::' + String(data[i][keyCol] || "");
    if (!seen.has(key)) {
      rowsToKeep.push(data[i]);
      seen.add(key);
    } else {
      duplicatesFound++;
    }
  }

  // Only rewrite if duplicates were found (saves time)
  if (duplicatesFound > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    logProgress(sheetName, `Removed ${duplicatesFound} duplicate rows`, false);
  }
}

function deduplicateAllOrders() {
  deduplicateSheet('Shopify Orders', SHOPIFY_ORDER_HEADERS);
  deduplicateSheet('Squarespace Orders', SQUARESPACE_ORDER_HEADERS);
  logImportEvent('Deduplication', 'All orders deduplicated');
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Deduplication complete', 'Dedup', 5);
  return "Deduplication complete";
}

