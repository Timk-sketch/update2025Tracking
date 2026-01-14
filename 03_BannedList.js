// =====================================================
// 03_BannedList.gs — Banned email list loading + checks
// =====================================================

// Cache banned list in memory to avoid redundant loads during same execution
let BANNED_LIST_CACHE_ = null;

function loadBannedList_() {
  // Return cached version if available
  if (BANNED_LIST_CACHE_ !== null) return BANNED_LIST_CACHE_;

  if (!BANNED_ARCHIVE_ID) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set() };
    return BANNED_LIST_CACHE_;
  }

  let bannedSheet = null;
  try {
    const ss = SpreadsheetApp.openById(BANNED_ARCHIVE_ID);
    bannedSheet = ss.getSheetByName(BANNED_SHEET_NAME_PRIMARY) || ss.getSheetByName(BANNED_SHEET_NAME_FALLBACK);
  } catch (e) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set() };
    return BANNED_LIST_CACHE_;
  }

  if (!bannedSheet) {
    BANNED_LIST_CACHE_ = { exact: new Set(), domains: new Set() };
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
