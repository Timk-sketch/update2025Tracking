// =====================================================
// 10_UI_Menu.gs — Pipeline + Sidebar + Menu + onOpen
// =====================================================

// NEW: Combined Import and Update workflow
function importAndUpdateAllOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Importing new orders and checking for refunds...', 'Import & Update', 5);

  // Step 1: Import new orders (last 14 days)
  importShopifyOrders();
  importSquarespaceOrders();

  // Step 2: Update existing orders with refunds (last 90 days)
  const shopifyMsg = updateShopifyOrdersWithRefunds();
  const squarespaceMsg = updateSquarespaceOrdersWithRefunds();

  // Step 3: Rebuild clean master and reports
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();

  const msg = `Import & Update complete!\nNew orders imported, refunds checked, reports rebuilt.`;
  ss.toast('✅ ' + msg, 'Complete', 10);
  return msg;
}

// NEW: Update-only function (for when you just want to check refunds without importing)
function updateAllOrdersWithRefunds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Checking for refunds in both platforms...', 'Update Orders', 5);

  // Update both Shopify and Squarespace orders with refunds
  const shopifyMsg = updateShopifyOrdersWithRefunds();
  const squarespaceMsg = updateSquarespaceOrdersWithRefunds();

  // Then rebuild clean master and reports
  deduplicateAllOrders();
  buildAllOrdersClean();
  buildOrdersSummaryReport();
  buildCustomerOutreachList();

  const msg = `Orders updated with refunds:\n${shopifyMsg}\n${squarespaceMsg}\nReports rebuilt successfully.`;
  ss.toast('✅ ' + msg, 'Update Complete', 10);
  return msg;
}

// PIPELINE (sidebar) - Just calls the new combined function
function runFullPipelineFromSidebar() {
  return importAndUpdateAllOrders();
}

// LEGACY: Pipeline with full imports (now just calls the combined function)
function runFullPipelineWithImports() {
  return importAndUpdateAllOrders();
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
    .addItem('⚡ Import & Update Orders', 'importAndUpdateAllOrders')
    .addItem('🔄 Update Orders (Refunds Only)', 'updateAllOrdersWithRefunds')
    .addSeparator()
    .addItem('Build Clean Master (All_Orders_Clean)', 'buildAllOrdersClean')
    .addItem('Build Orders Summary Report', 'buildOrdersSummaryReport')
    .addItem('Build Customer Outreach List', 'buildCustomerOutreachList')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Admin / Advanced')
      .addItem('📥 Import Shopify Orders Only', 'importShopifyOrders')
      .addItem('📥 Import Squarespace Orders Only', 'importSquarespaceOrders')
      .addSeparator()
      .addItem('⚠️ Import ALL Shopify History', 'importShopifyOrdersFullHistory')
      .addItem('⚠️ Import ALL Squarespace History', 'importSquarespaceOrdersFullHistory')
      .addSeparator()
      .addItem('Refresh Shopify Refunds (30 days)', 'refreshShopifyAdjustments')
      .addItem('Refresh Shopify Refunds (60 days)', 'refreshShopifyAdjustmentsLast60Days')
      .addItem('Refresh Squarespace Refunds (30 days)', 'refreshSquarespaceAdjustments')
      .addItem('Refresh Squarespace Refunds (60 days)', 'refreshSquarespaceAdjustmentsLast60Days')
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
