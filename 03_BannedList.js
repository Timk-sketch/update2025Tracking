// =====================================================
// 03_BannedList.gs — Banned email list loading + checks
// =====================================================

// Cache banned list in memory to avoid redundant loads during same execution
let BANNED_LIST_CACHE_ = null;

function loadBannedList_() {
  // Return cached version if available
  if (BANNED_LIST_CACHE_ !== null) return BANNED_LIST_CACHE_;

  if (!BANNED_ARCHIVE_ID) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set(["dirtlegal.com"]) };
    return BANNED_LIST_CACHE_;
  }

  let bannedSheet = null;
  try {
    const ss = SpreadsheetApp.openById(BANNED_ARCHIVE_ID);
    bannedSheet = ss.getSheetByName(BANNED_SHEET_NAME_PRIMARY) || ss.getSheetByName(BANNED_SHEET_NAME_FALLBACK);
  } catch (e) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set(["dirtlegal.com"]) };
    return BANNED_LIST_CACHE_;
  }

  if (!bannedSheet) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set(["dirtlegal.com"]) };
    return BANNED_LIST_CACHE_;
  }

  const values = bannedSheet
    .getRange(2, 1, Math.max(0, bannedSheet.getLastRow() - 1), 1)
    .getValues()
    .flat();

  const exact = new Set();
  const domains = new Set();

  values
    .map(v => String(v || "").trim().toLowerCase())
    .filter(Boolean)
    .forEach(cell => {
      cell.split(/[,\s;]+/).map(x => x.trim()).filter(Boolean).forEach(entry => {
        if (entry.startsWith("*@")) entry = entry.substring(1);
        if (entry.startsWith("@")) { domains.add(entry.substring(1)); return; }
        if (entry.includes("@")) exact.add(normalizeEmailForCompare_(entry));
        else domains.add(entry);
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
 * Creates or retrieves the banned list spreadsheet.
 * If BANNED_ARCHIVE_ID is not set, creates a new spreadsheet in user's Drive.
 * Returns the spreadsheet ID.
 */
function setupBannedListSpreadsheet() {
  let bannedArchiveId = PROPS.getProperty('BANNED_ARCHIVE_ID');

  if (bannedArchiveId) {
    // Verify it exists
    try {
      SpreadsheetApp.openById(bannedArchiveId);
      return `Banned list already configured: ${bannedArchiveId}`;
    } catch (e) {
      // ID is set but invalid, will create new one below
      PROPS.deleteProperty('BANNED_ARCHIVE_ID');
    }
  }

  // Create new banned list spreadsheet
  const newSS = SpreadsheetApp.create('Banned Email List');
  const sheet = newSS.getActiveSheet();
  sheet.setName(BANNED_SHEET_NAME_PRIMARY);

  // Set up headers
  sheet.getRange(1, 1, 1, 2).setValues([["Email/Domain", "Date Added"]]);
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  // Add dirtlegal.com as the first entry (already hardcoded, but good for reference)
  sheet.appendRow(["dirtlegal.com", new Date()]);
  sheet.appendRow(["@dirtlegal.com", new Date()]);

  // Save the ID
  const newId = newSS.getId();
  PROPS.setProperty('BANNED_ARCHIVE_ID', newId);

  // Clear cache
  BANNED_LIST_CACHE_ = null;

  const url = newSS.getUrl();
  const msg = `✅ Created banned list spreadsheet: ${url}`;
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Banned List Setup", 10);
  logImportEvent("Banned List", `Created new spreadsheet: ${newId}`);

  return msg;
}

/**
 * Adds an email address or domain to the banned list.
 * Supports: full emails (user@domain.com), domains (domain.com or @domain.com)
 */
function addToBannedList(emailOrDomain) {
  const input = String(emailOrDomain || "").trim().toLowerCase();
  if (!input) throw new Error("Please enter an email address or domain to ban.");

  // Check if BANNED_ARCHIVE_ID is set, if not, create it
  let bannedArchiveId = PROPS.getProperty('BANNED_ARCHIVE_ID');
  if (!bannedArchiveId) {
    setupBannedListSpreadsheet();
    bannedArchiveId = PROPS.getProperty('BANNED_ARCHIVE_ID');
  }

  let bannedSheet = null;
  try {
    const ss = SpreadsheetApp.openById(bannedArchiveId);
    bannedSheet = ss.getSheetByName(BANNED_SHEET_NAME_PRIMARY);

    // Create sheet if it doesn't exist
    if (!bannedSheet) {
      bannedSheet = ss.insertSheet(BANNED_SHEET_NAME_PRIMARY);
      bannedSheet.getRange(1, 1, 1, 2).setValues([["Email/Domain", "Date Added"]]);
      bannedSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");
      bannedSheet.setFrozenRows(1);
    }
  } catch (e) {
    throw new Error(`Cannot access banned list spreadsheet: ${e.message}. Try running "Setup Banned List" from Admin menu.`);
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

  // Clear cache so it reloads next time
  BANNED_LIST_CACHE_ = null;

  const msg = `✅ Added "${input}" to banned list. Rebuild All_Order_Clean to apply.`;
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Banned List", 6);
  logImportEvent("Banned List", `Added: ${input}`);

  return msg;
}
