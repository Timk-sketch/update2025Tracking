// =====================================================
// 03_BannedList.gs — Banned email list loading + checks
// Uses local "Banned_Emails" tab in the current spreadsheet
// =====================================================

// Cache banned list in memory to avoid redundant loads during same execution
let BANNED_LIST_CACHE_ = null;

const BANNED_EMAILS_SHEET_NAME = 'Banned_Emails';

/**
 * Loads banned emails/domains from the Banned_Emails tab in the current spreadsheet.
 * Always includes dirtlegal.com as a hardcoded banned domain.
 */
function loadBannedList_() {
  // Return cached version if available
  if (BANNED_LIST_CACHE_ !== null) return BANNED_LIST_CACHE_;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bannedSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);

  // If sheet doesn't exist, return just the hardcoded domain
  if (!bannedSheet) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set(["dirtlegal.com"]) };
    return BANNED_LIST_CACHE_;
  }

  const lastRow = bannedSheet.getLastRow();
  if (lastRow < 2) {
    // Only header row or empty
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set(["dirtlegal.com"]) };
    return BANNED_LIST_CACHE_;
  }

  // Read all email/domain entries (skip header row)
  const values = bannedSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  const exact = new Set();
  const domains = new Set();

  values
    .map(v => String(v || "").trim().toLowerCase())
    .filter(Boolean)
    .forEach(cell => {
      // Support comma/space/semicolon separated entries in a single cell
      cell.split(/[,\s;]+/).map(x => x.trim()).filter(Boolean).forEach(entry => {
        if (entry.startsWith("*@")) entry = entry.substring(1);
        if (entry.startsWith("@")) {
          domains.add(entry.substring(1));
          return;
        }
        if (entry.includes("@")) {
          exact.add(normalizeEmailForCompare_(entry));
        } else {
          domains.add(entry);
        }
      });
    });

  // Always exclude @dirtlegal.com emails (internal company domain)
  domains.add("dirtlegal.com");

  BANNED_LIST_CACHE_ = { exact, domains };
  return BANNED_LIST_CACHE_;
}

function isBannedEmail_(emailRaw, banned) {
  const email = normalizeEmailForCompare_(emailRaw);
  if (!email) return false;
  if (banned.exact.has(email)) return true;

  const at = email.lastIndexOf("@");
  if (at === -1) return false;

  const domain = email.substring(at + 1);
  for (const d of banned.domains) {
    if (domain === d || domain.endsWith("." + d)) return true;
  }
  return false;
}

/**
 * Creates the Banned_Emails tab in the current spreadsheet if it doesn't exist.
 * Pre-populates with dirtlegal.com entries.
 */
function setupBannedEmailsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bannedSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);

  if (bannedSheet) {
    return `Banned_Emails tab already exists.`;
  }

  // Create the tab
  bannedSheet = ss.insertSheet(BANNED_EMAILS_SHEET_NAME);

  // Set up headers
  bannedSheet.getRange(1, 1, 1, 2).setValues([["Email/Domain", "Date Added"]]);
  bannedSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#ea4335").setFontColor("#ffffff");
  bannedSheet.setFrozenRows(1);

  // Add dirtlegal.com as the first entries
  const now = new Date();
  bannedSheet.appendRow(["dirtlegal.com", now]);
  bannedSheet.appendRow(["@dirtlegal.com", now]);

  // Auto-resize columns
  bannedSheet.autoResizeColumns(1, 2);

  // Clear cache
  BANNED_LIST_CACHE_ = null;

  const msg = `✅ Created Banned_Emails tab with dirtlegal.com pre-populated.`;
  ss.toast(msg, "Banned Emails Setup", 6);
  logImportEvent("Banned Emails", "Created Banned_Emails tab");

  return msg;
}

/**
 * Adds an email address or domain to the Banned_Emails tab.
 * Supports: full emails (user@domain.com), domains (domain.com or @domain.com)
 */
function addToBannedList(emailOrDomain) {
  const input = String(emailOrDomain || "").trim().toLowerCase();
  if (!input) throw new Error("Please enter an email address or domain to ban.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let bannedSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!bannedSheet) {
    setupBannedEmailsTab();
    bannedSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);
  }

  // Check if already exists
  const lastRow = Math.max(1, bannedSheet.getLastRow());
  if (lastRow > 1) {
    const existing = bannedSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const existingLower = existing.map(e => String(e || "").trim().toLowerCase());

    // Check for exact match or if it's already there with @ prefix
    if (existingLower.includes(input) ||
        existingLower.includes("@" + input) ||
        (input.startsWith("@") && existingLower.includes(input.substring(1)))) {
      return `"${input}" is already in the banned list.`;
    }
  }

  // Add to banned list
  const timestamp = new Date();
  bannedSheet.appendRow([input, timestamp]);

  // Auto-resize columns
  bannedSheet.autoResizeColumns(1, 2);

  // Clear cache so it reloads next time
  BANNED_LIST_CACHE_ = null;

  const msg = `✅ Added "${input}" to Banned_Emails tab. Rebuild All_Order_Clean to apply.`;
  ss.toast(msg, "Banned List", 6);
  logImportEvent("Banned List", `Added: ${input}`);

  // Optionally, clean existing All_Order_Clean data immediately
  // Uncomment the next line if you want automatic cleaning:
  // cleanBannedEmailsFromAllOrdersClean();

  return msg;
}

/**
 * Removes all orders with banned emails from All_Order_Clean.
 * This is faster than rebuilding the entire clean master.
 */
function cleanBannedEmailsFromAllOrdersClean() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanSheet = ss.getSheetByName(CLEAN_OUTPUT_SHEET || 'All_Order_Clean');

  if (!cleanSheet) {
    return 'All_Order_Clean sheet not found.';
  }

  const lastRow = cleanSheet.getLastRow();
  if (lastRow < 2) {
    return 'All_Order_Clean is empty.';
  }

  // Load banned list
  const banned = loadBannedList_();

  // Read all data
  const data = cleanSheet.getDataRange().getValues();
  const headers = data[0];

  // Find email column
  const emailCol = headers.indexOf('customer_email_raw');
  if (emailCol === -1) {
    throw new Error('customer_email_raw column not found in All_Order_Clean');
  }

  // Filter out banned emails
  const headerRow = [headers];
  const cleanRows = [];
  let removedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const emailRaw = data[i][emailCol];

    if (isBannedEmail_(emailRaw, banned)) {
      removedCount++;
      continue;
    }

    cleanRows.push(data[i]);
  }

  // Only rewrite if we removed something
  if (removedCount > 0) {
    cleanSheet.clearContents();
    const allRows = headerRow.concat(cleanRows);
    cleanSheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);

    // Reapply formatting
    cleanSheet.setFrozenRows(1);
    cleanSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    const msg = `✅ Removed ${removedCount} orders with banned emails from All_Order_Clean.`;
    ss.toast(msg, 'Clean Banned Emails', 6);
    logImportEvent('Clean Banned Emails', msg, removedCount);
    return msg;
  }

  return 'No orders with banned emails found in All_Order_Clean.';
}

/**
 * Imports all emails from an external banned list spreadsheet into the local Banned_Emails tab.
 * Useful for migrating from an old external banned list.
 */
function importBannedListFromExternal() {
  const bannedArchiveId = PROPS.getProperty('BANNED_ARCHIVE_ID');

  if (!bannedArchiveId) {
    throw new Error("No external banned list configured (BANNED_ARCHIVE_ID not set).");
  }

  let externalSheet = null;
  try {
    const externalSS = SpreadsheetApp.openById(bannedArchiveId);
    externalSheet = externalSS.getSheetByName(BANNED_SHEET_NAME_PRIMARY) ||
                    externalSS.getSheetByName(BANNED_SHEET_NAME_FALLBACK);
  } catch (e) {
    throw new Error(`Cannot access external banned list: ${e.message}`);
  }

  if (!externalSheet) {
    throw new Error("External banned list sheet not found.");
  }

  // Get current spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let localSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);

  // Create local sheet if it doesn't exist
  if (!localSheet) {
    setupBannedEmailsTab();
    localSheet = ss.getSheetByName(BANNED_EMAILS_SHEET_NAME);
  }

  // Read external data
  const externalLastRow = externalSheet.getLastRow();
  if (externalLastRow < 2) {
    return "External banned list is empty.";
  }

  const externalData = externalSheet.getRange(2, 1, externalLastRow - 1, 1).getValues().flat();

  // Get existing local entries
  const localLastRow = Math.max(1, localSheet.getLastRow());
  const existingLocal = localLastRow > 1
    ? localSheet.getRange(2, 1, localLastRow - 1, 1).getValues().flat().map(e => String(e || "").trim().toLowerCase())
    : [];

  // Add entries that don't already exist
  let added = 0;
  const timestamp = new Date();

  externalData.forEach(entry => {
    const entryLower = String(entry || "").trim().toLowerCase();
    if (entryLower && !existingLocal.includes(entryLower)) {
      localSheet.appendRow([entryLower, timestamp]);
      added++;
    }
  });

  // Auto-resize columns
  localSheet.autoResizeColumns(1, 2);

  // Clear cache
  BANNED_LIST_CACHE_ = null;

  const msg = `✅ Imported ${added} entries from external banned list.`;
  ss.toast(msg, "Import Complete", 6);
  logImportEvent("Banned List", `Imported ${added} entries from external list`);

  return msg;
}
