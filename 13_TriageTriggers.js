// =====================================================
// 13_TriageTriggers.js — Automated triage trigger setup
// =====================================================

/**
 * Setup automated triage triggers
 * - Every 2 hours: Import 0-30 days + Clean Triage
 * - Daily at midnight: Import 31-60 days + Clean Triage
 * - Weekly on Sunday 6 AM: Import 61-120 days + Clean Triage
 */
function setupTriageTriggers() {
  // First, delete any existing triage triggers to avoid duplicates
  deleteTriageTriggers();

  // Every 2 hours: Recent orders (0-30 days)
  ScriptApp.newTrigger('runTriageEvery2Hours')
    .timeBased()
    .everyHours(2)
    .create();

  // Daily at midnight: Older orders (31-60 days)
  ScriptApp.newTrigger('runTriageDaily')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  // Weekly on Sunday at 6 AM: Full sweep (61-120 days)
  ScriptApp.newTrigger('runTriageWeekly')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(6)
    .create();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msg = 'Triage triggers set up:\n• Every 2 hours (0-30 days)\n• Daily at midnight (31-60 days)\n• Weekly Sunday 6 AM (61-120 days)';
  ss.toast(msg, 'Triggers Active', 10);
  logImportEvent('Triage Triggers', 'Setup complete', 3);
  return msg;
}

/**
 * Delete all triage-related triggers
 */
function deleteTriageTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const triageFunctionNames = [
    'runTriageEvery2Hours',
    'runTriageDaily',
    'runTriageWeekly'
  ];

  let deleted = 0;
  triggers.forEach(trigger => {
    if (triageFunctionNames.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });

  if (deleted > 0) {
    logImportEvent('Triage Triggers', `Deleted ${deleted} old triggers`, deleted);
  }

  return `Deleted ${deleted} triage triggers`;
}

/**
 * Runs every 2 hours: Import recent orders (0-30 days) and clean
 */
function runTriageEvery2Hours() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Import 0-30 days for both platforms
    importShopifyToTriage("0-30");
    importSquarespaceToTriage("0-30");

    // Clean triage (update main sheets)
    cleanTriage();

    logImportEvent('Triage Auto', 'Every 2 hours complete (0-30 days)');
  } catch (error) {
    logImportEvent('Triage Auto', `ERROR in 2-hour run: ${error.message}`);
    throw error;
  }
}

/**
 * Runs daily at midnight: Import older orders (31-60 days) and clean
 */
function runTriageDaily() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Import 31-60 days for both platforms
    importShopifyToTriage("31-60");
    importSquarespaceToTriage("31-60");

    // Clean triage (update main sheets)
    cleanTriage();

    logImportEvent('Triage Auto', 'Daily run complete (31-60 days)');
  } catch (error) {
    logImportEvent('Triage Auto', `ERROR in daily run: ${error.message}`);
    throw error;
  }
}

/**
 * Runs weekly on Sunday 6 AM: Full sweep (61-120 days) and clean
 */
function runTriageWeekly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Import 61-90 and 91-120 days for both platforms
    importShopifyToTriage("61-90");
    importSquarespaceToTriage("61-90");

    importShopifyToTriage("91-120");
    importSquarespaceToTriage("91-120");

    // Clean triage (update main sheets)
    cleanTriage();

    // Run clean again in case there are many rows
    // This ensures triage sheets are completely empty
    const triageShopify = ss.getSheetByName('Shopify Triage');
    const triageSquare = ss.getSheetByName('Squarespace Triage');

    const hasShopifyRows = triageShopify && triageShopify.getLastRow() > 1;
    const hasSquareRows = triageSquare && triageSquare.getLastRow() > 1;

    if (hasShopifyRows || hasSquareRows) {
      // Run clean again to process remaining rows
      cleanTriage();
    }

    logImportEvent('Triage Auto', 'Weekly Sunday run complete (61-120 days)');
  } catch (error) {
    logImportEvent('Triage Auto', `ERROR in weekly run: ${error.message}`);
    throw error;
  }
}

/**
 * View current trigger status
 */
function viewTriageTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const triageTriggers = triggers.filter(t =>
    t.getHandlerFunction().startsWith('runTriage')
  );

  if (triageTriggers.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'No triage triggers found. Run "Setup Triage Triggers" to activate automated updates.',
      'Triggers',
      8
    );
    return 'No triage triggers active';
  }

  const status = triageTriggers.map(t => {
    const func = t.getHandlerFunction();
    const source = t.getTriggerSource();
    return `${func}: ${source}`;
  }).join('\n');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Active triage triggers:\n${status}`,
    'Triggers',
    10
  );

  return status;
}

/**
 * Manual test functions - run these to test without waiting for triggers
 */
function testTriageEvery2Hours() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Testing 2-hour triage run...', 'Test', 5);
  return runTriageEvery2Hours();
}

function testTriageDaily() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Testing daily triage run...', 'Test', 5);
  return runTriageDaily();
}

function testTriageWeekly() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Testing weekly triage run...', 'Test', 5);
  return runTriageWeekly();
}
