// =====================================================
// 10_UI_Menu.gs — Pipeline + Sidebar + Menu + onOpen
// =====================================================

// PIPELINE (sidebar)
function runFullPipelineFromSidebar() {
  importShopifyOrders();
  importSquarespaceOrders();
  refreshShopifyAdjustments();
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();
  return "Full pipeline complete (Import → Refresh Refunds/Discounts → Dedup → Clean → Summary → Outreach)";
}

function runFullPipelineTightLast60Days() {
  importShopifyOrders();
  importSquarespaceOrders();
  refreshShopifyAdjustmentsLast60Days();
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();
  return "TIGHT pipeline complete (Import → Refresh last 60 days → Dedup → Clean → Summary → Outreach)";
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
    .addItem('Import Shopify Orders', 'importShopifyOrders')
    .addItem('⚠️ Import ALL Shopify History', 'importShopifyOrdersFullHistory')
    .addItem('Import Squarespace Orders', 'importSquarespaceOrders')
    .addItem('⚠️ Import ALL Squarespace History', 'importSquarespaceOrdersFullHistory')
    .addSeparator()
    .addItem('Refresh Shopify Discounts + Refunds (incremental)', 'refreshShopifyAdjustments')
    .addItem('Refresh Shopify Discounts + Refunds (FORCE last 60 days)', 'refreshShopifyAdjustmentsLast60Days')
    .addSeparator()
    .addItem('Deduplicate All Orders', 'deduplicateAllOrders')
    .addSeparator()
    .addItem('Build Clean Master (All_Orders_Clean)', 'buildAllOrdersClean')
    .addItem('Build Orders Summary Report', 'buildOrdersSummaryReport')
    .addItem('Build Customer Outreach List', 'buildCustomerOutreachList')
    .addSeparator()
    .addItem('Run Full Pipeline (fast)', 'runFullPipelineFromSidebar')
    .addItem('Run Full Pipeline (TIGHT last 60 days)', 'runFullPipelineTightLast60Days')
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
