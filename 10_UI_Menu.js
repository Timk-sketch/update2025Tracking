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
    .addSubMenu(ui.createMenu('📥 Triage System (Recommended)')
      .addItem('🔄 Clean Triage (Update Main Sheets)', 'cleanTriage')
      .addSeparator()
      .addItem('Import Shopify: 0-30 days', 'importShopifyTriage0to30')
      .addItem('Import Shopify: 31-60 days', 'importShopifyTriage31to60')
      .addItem('Import Shopify: 61-90 days', 'importShopifyTriage61to90')
      .addItem('Import Shopify: 91-120 days', 'importShopifyTriage91to120')
      .addSeparator()
      .addItem('Import Squarespace: 0-30 days', 'importSquarespaceTriage0to30')
      .addItem('Import Squarespace: 31-60 days', 'importSquarespaceTriage31to60')
      .addItem('Import Squarespace: 61-90 days', 'importSquarespaceTriage61to90')
      .addItem('Import Squarespace: 91-120 days', 'importSquarespaceTriage91to120'))
    .addSeparator()
    .addItem('Import Shopify Orders', 'importShopifyOrders')
    .addItem('⚠️ Import ALL Shopify History', 'importShopifyOrdersFullHistory')
    .addItem('Import Squarespace Orders', 'importSquarespaceOrders')
    .addItem('⚠️ Import ALL Squarespace History', 'importSquarespaceOrdersFullHistory')
    .addSeparator()
    .addItem('Refresh Shopify Discounts + Refunds (last 30 days)', 'refreshShopifyAdjustments')
    .addItem('Refresh Shopify Discounts + Refunds (last 60 days with append)', 'refreshShopifyAdjustmentsLast60Days')
    .addItem('Refresh Squarespace Refunds (last 30 days)', 'refreshSquarespaceAdjustments')
    .addItem('Refresh Squarespace Refunds (last 60 days with append)', 'refreshSquarespaceAdjustmentsLast60Days')
    .addSeparator()
    .addItem('Deduplicate All Orders', 'deduplicateAllOrders')
    .addSeparator()
    .addItem('Build Clean Master (All_Orders_Clean)', 'buildAllOrdersClean')
    .addItem('Build Orders Summary Report', 'buildOrdersSummaryReport')
    .addItem('Build Customer Outreach List', 'buildCustomerOutreachList')
    .addSeparator()
    .addItem('⚡ Run Pipeline (FAST - refresh only)', 'runFullPipelineFromSidebar')
    .addItem('⚡ Run Pipeline (TIGHT - last 60 days)', 'runFullPipelineTightLast60Days')
    .addItem('⏱️ Run Pipeline with Full Imports (SLOW)', 'runFullPipelineWithImports')
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
