// =====================================================
// 10_UI_Menu.gs — Pipeline + Sidebar + Menu + onOpen
// =====================================================

// PIPELINE (sidebar)
function runFullPipelineFromSidebar() {
  // Only refresh refunds/discounts (much faster than full import)
  // Full imports should be run separately when needed
  refreshShopifyAdjustments();
  refreshSquarespaceAdjustments();
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();
  return "Full pipeline complete (Refresh Refunds/Discounts → Dedup → Clean → Summary → Outreach)";
}

// NEW: Pipeline with full imports (slower, use when you need to import new orders)
function runFullPipelineWithImports() {
  importShopifyOrders();
  importSquarespaceOrders();
  refreshShopifyAdjustments();
  refreshSquarespaceAdjustments();
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();
  return "Full pipeline with imports complete (Import → Refresh → Dedup → Clean → Summary → Outreach)";
}

function runFullPipelineTightLast60Days() {
  // Refresh last 60 days with append (faster, no full import)
  refreshShopifyAdjustmentsLast60Days();
  refreshSquarespaceAdjustmentsLast60Days();
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();
  return "TIGHT pipeline complete (Refresh last 60 days → Dedup → Clean → Summary → Outreach)";
}

// MENU
function showSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Order Tools Sidebar')
      .setWidth(380);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    ss.toast(`Sidebar.html missing or error: ${e.message}`, "Sidebar error", 8);
    throw e;
  }
}

function rebuildOrderToolsMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Order Tools')
    .addItem('Show Sidebar', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 Triage System')
      .addItem('🔄 Clean Triage (Update Main Sheets)', 'cleanTriage')
      .addSeparator()
      .addItem('⚙️ Setup Automated Triggers', 'setupTriageTriggers')
      .addItem('📊 View Trigger Status', 'viewTriageTriggerStatus')
      .addItem('🗑️ Delete All Triggers', 'deleteTriageTriggers'))
    .addSeparator()
    .addItem('Build Clean Master (All_Orders_Clean)', 'buildAllOrdersClean')
    .addItem('Build Orders Summary Report', 'buildOrdersSummaryReport')
    .addItem('Build Customer Outreach List', 'buildCustomerOutreachList')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Admin / Advanced')
      .addItem('Import Shopify Orders (NEW only)', 'importShopifyOrders')
      .addItem('Import Squarespace Orders (NEW only)', 'importSquarespaceOrders')
      .addSeparator()
      .addItem('⚠️ Import ALL Shopify History', 'importShopifyOrdersFullHistory')
      .addItem('⚠️ Import ALL Squarespace History', 'importSquarespaceOrdersFullHistory')
      .addSeparator()
      .addItem('🧪 Test 2-Hour Triage Run', 'testTriageEvery2Hours')
      .addItem('🧪 Test Daily Triage Run', 'testTriageDaily')
      .addItem('🧪 Test Weekly Triage Run', 'testTriageWeekly')
      .addSeparator()
      .addSubMenu(ui.createMenu('Manual Triage Imports')
        .addItem('Shopify: 0-30 days', 'importShopifyTriage0to30')
        .addItem('Shopify: 31-60 days', 'importShopifyTriage31to60')
        .addItem('Shopify: 61-90 days', 'importShopifyTriage61to90')
        .addItem('Shopify: 91-120 days', 'importShopifyTriage91to120')
        .addSeparator()
        .addItem('Squarespace: 0-30 days', 'importSquarespaceTriage0to30')
        .addItem('Squarespace: 31-60 days', 'importSquarespaceTriage31to60')
        .addItem('Squarespace: 61-90 days', 'importSquarespaceTriage61to90')
        .addItem('Squarespace: 91-120 days', 'importSquarespaceTriage91to120'))
      .addSeparator()
      .addItem('Refresh Shopify Adjustments (30 days)', 'refreshShopifyAdjustments')
      .addItem('Refresh Shopify Adjustments (60 days)', 'refreshShopifyAdjustmentsLast60Days')
      .addItem('Refresh Squarespace Adjustments (30 days)', 'refreshSquarespaceAdjustments')
      .addItem('Refresh Squarespace Adjustments (60 days)', 'refreshSquarespaceAdjustmentsLast60Days')
      .addSeparator()
      .addItem('Deduplicate All Orders', 'deduplicateAllOrders')
      .addSeparator()
      .addItem('⚡ Pipeline (FAST - refresh only)', 'runFullPipelineFromSidebar')
      .addItem('⚡ Pipeline (TIGHT - last 60 days)', 'runFullPipelineTightLast60Days')
      .addItem('⏱️ Pipeline (with Full Imports)', 'runFullPipelineWithImports'))
    .addToUi();
}

function onOpen() {
  rebuildOrderToolsMenu();
}

function getFullHistoryResumeStatus() {
  const shopifyNext = PROPS.getProperty('SHOPIFY_FULLHISTORY_NEXT_URL');
  const squareCursor = PROPS.getProperty('SQUARESPACE_FULLHISTORY_CURSOR');

  return {
    shopify: shopifyNext ? 'IN PROGRESS (resumable)' : 'Not running / Complete',
    squarespace: squareCursor ? 'IN PROGRESS (resumable)' : 'Not running / Complete'
  };
}

function resetFullHistoryResumePointers() {
  PROPS.deleteProperty('SHOPIFY_FULLHISTORY_NEXT_URL');
  PROPS.deleteProperty('SQUARESPACE_FULLHISTORY_CURSOR');
  SpreadsheetApp.getActiveSpreadsheet().toast('Full-history resume pointers cleared.', 'Reset', 5);
  return 'Resume pointers cleared.';
}
