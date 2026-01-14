// =====================================================
// 10_UI_Menu.gs — Pipeline + Sidebar + Menu + onOpen
// =====================================================

// AUTOMATION PART 1: Import and update orders only (for time-based trigger)
function automatedImportAndUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const steps = [];

  logProgress('Automated Import', '🚀 Starting automated import and update...');

  // Step 1: Import new orders (last 14 days)
  logProgress('Automated Import', '📥 Step 1/6: Importing Shopify orders (last 14 days)...');
  const shopifyImportMsg = importShopifyOrders();
  steps.push('✓ Shopify Import: ' + shopifyImportMsg);

  logProgress('Automated Import', '📥 Step 2/6: Importing Squarespace orders (last 14 days)...');
  const squarespaceImportMsg = importSquarespaceOrders();
  steps.push('✓ Squarespace Import: ' + squarespaceImportMsg);

  // Step 2: Update existing orders with refunds (last 90 days)
  logProgress('Automated Import', '🔄 Step 3/6: Checking Shopify refunds (last 90 days)...');
  const shopifyRefundMsg = updateShopifyOrdersWithRefunds();
  steps.push('✓ Shopify Refunds: ' + shopifyRefundMsg);

  logProgress('Automated Import', '🔄 Step 4/6: Checking Squarespace refunds (last 90 days)...');
  const squarespaceRefundMsg = updateSquarespaceOrdersWithRefunds();
  steps.push('✓ Squarespace Refunds: ' + squarespaceRefundMsg);

  // Step 3: Prepare data
  logProgress('Automated Import', '🧹 Step 5/6: Deduplicating orders...');
  deduplicateAllOrders();
  steps.push('✓ Deduplication complete');

  logProgress('Automated Import', '📊 Step 6/6: Building clean master sheet...');
  buildAllOrdersClean();
  steps.push('✓ Clean master built');

  const msg = '✅ Part 1 Complete (Import & Update)!\n\n' + steps.join('\n');
  logProgress('Automated Import', '✅ All 6 steps complete! Run automatedBuildReports next.');
  logImportEvent('Automated Import', 'Part 1: Import & Update finished', steps.length);
  return msg;
}

// AUTOMATION PART 2: Build all reports (for time-based trigger, runs after Part 1)
function automatedBuildReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const steps = [];

  logProgress('Automated Reports', '📊 Starting automated report building...');

  logProgress('Automated Reports', '📈 Step 1/4: Building summary report...');
  buildOrdersSummaryReport();
  steps.push('✓ Summary report built');

  logProgress('Automated Reports', '💰 Step 2/4: Building refunds report...');
  buildRefundsReport();
  steps.push('✓ Refunds report built');

  logProgress('Automated Reports', '🏷️ Step 3/4: Building discounts report...');
  buildDiscountsReport();
  steps.push('✓ Discounts report built');

  logProgress('Automated Reports', '📧 Step 4/4: Building customer outreach list...');
  buildCustomerOutreachList();
  steps.push('✓ Outreach list built');

  const msg = '✅ Part 2 Complete (All Reports)!\n\n' + steps.join('\n');
  logProgress('Automated Reports', '✅ All 4 reports built!');
  logImportEvent('Automated Reports', 'Part 2: Report building finished', steps.length);
  return msg;
}

// MANUAL: Combined Import and Update workflow (for sidebar button)
function importAndUpdateAllOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const steps = [];

  logProgress('Import & Update', '🚀 Starting full import and update workflow...');

  // Step 1: Import new orders (last 14 days)
  logProgress('Import & Update', '📥 Step 1/8: Importing Shopify orders (last 14 days)...');
  const shopifyImportMsg = importShopifyOrders();
  steps.push('✓ Shopify Import: ' + shopifyImportMsg);

  logProgress('Import & Update', '📥 Step 2/8: Importing Squarespace orders (last 14 days)...');
  const squarespaceImportMsg = importSquarespaceOrders();
  steps.push('✓ Squarespace Import: ' + squarespaceImportMsg);

  // Step 2: Update existing orders with refunds (last 90 days)
  logProgress('Import & Update', '🔄 Step 3/8: Checking Shopify refunds (last 90 days)...');
  const shopifyRefundMsg = updateShopifyOrdersWithRefunds();
  steps.push('✓ Shopify Refunds: ' + shopifyRefundMsg);

  logProgress('Import & Update', '🔄 Step 4/8: Checking Squarespace refunds (last 90 days)...');
  const squarespaceRefundMsg = updateSquarespaceOrdersWithRefunds();
  steps.push('✓ Squarespace Refunds: ' + squarespaceRefundMsg);

  // Step 3: Rebuild clean master and reports
  logProgress('Import & Update', '🧹 Step 5/8: Deduplicating orders...');
  deduplicateAllOrders();
  steps.push('✓ Deduplication complete');

  logProgress('Import & Update', '📊 Step 6/8: Building clean master sheet...');
  buildAllOrdersClean();
  steps.push('✓ Clean master built');

  logProgress('Import & Update', '📈 Step 7/10: Building summary report...');
  buildOrdersSummaryReport();
  steps.push('✓ Summary report built');

  logProgress('Import & Update', '💰 Step 8/10: Building refunds report...');
  buildRefundsReport();
  steps.push('✓ Refunds report built');

  logProgress('Import & Update', '🏷️ Step 9/10: Building discounts report...');
  buildDiscountsReport();
  steps.push('✓ Discounts report built');

  logProgress('Import & Update', '📧 Step 10/10: Building customer outreach list...');
  buildCustomerOutreachList();
  steps.push('✓ Outreach list built');

  const msg = '✅ Import & Update Complete!\n\n' + steps.join('\n');
  logProgress('Import & Update', '✅ All 10 steps complete!');
  logImportEvent('Import & Update', 'Complete workflow finished', steps.length);
  return msg;
}

// NEW: Update-only function (for when you just want to check refunds without importing)
function updateAllOrdersWithRefunds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const steps = [];

  logProgress('Update Orders', '🚀 Starting refund check workflow...');

  // Update both Shopify and Squarespace orders with refunds
  logProgress('Update Orders', '🔄 Step 1/6: Checking Shopify refunds (last 90 days)...');
  const shopifyMsg = updateShopifyOrdersWithRefunds();
  steps.push('✓ Shopify: ' + shopifyMsg);

  logProgress('Update Orders', '🔄 Step 2/6: Checking Squarespace refunds (last 90 days)...');
  const squarespaceMsg = updateSquarespaceOrdersWithRefunds();
  steps.push('✓ Squarespace: ' + squarespaceMsg);

  // Then rebuild clean master and reports
  logProgress('Update Orders', '🧹 Step 3/6: Deduplicating orders...');
  deduplicateAllOrders();
  steps.push('✓ Deduplication complete');

  logProgress('Update Orders', '📊 Step 4/6: Building clean master sheet...');
  buildAllOrdersClean();
  steps.push('✓ Clean master built');

  logProgress('Update Orders', '📈 Step 5/6: Building summary report...');
  buildOrdersSummaryReport();
  steps.push('✓ Summary report built');

  logProgress('Update Orders', '📧 Step 6/6: Building customer outreach list...');
  buildCustomerOutreachList();
  steps.push('✓ Outreach list built');

  const msg = '✅ Refund Check Complete!\n\n' + steps.join('\n');
  logProgress('Update Orders', '✅ All 6 steps complete!');
  logImportEvent('Update Orders', 'Refund check complete', steps.length);
  return msg;
}

// PIPELINE (sidebar) - Runs the complete workflow with detailed progress
function runFullPipelineFromSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Starting full workflow: Import → Update → Build All Reports', 'Full Pipeline', 3);

  // Run the combined import and update
  const importUpdateMsg = importAndUpdateAllOrders();

  // Return detailed message
  return 'Full Pipeline Complete!\n' + importUpdateMsg;
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
    .addItem('📊 Show Sidebar', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Admin / Setup')
      .addItem('📥 Import Shopify Orders Only', 'importShopifyOrders')
      .addItem('📥 Import Squarespace Orders Only', 'importSquarespaceOrders')
      .addSeparator()
      .addItem('Refresh Shopify Refunds (30 days)', 'refreshShopifyAdjustments')
      .addItem('Refresh Shopify Refunds (60 days)', 'refreshShopifyAdjustmentsLast60Days')
      .addItem('Refresh Squarespace Refunds (30 days)', 'refreshSquarespaceAdjustments')
      .addItem('Refresh Squarespace Refunds (60 days)', 'refreshSquarespaceAdjustmentsLast60Days')
      .addSeparator()
      .addItem('Deduplicate All Orders', 'deduplicateAllOrders')
      .addItem('Build Clean Master Only', 'buildAllOrdersClean'))
    .addToUi();
}

function onOpen() {
  rebuildOrderToolsMenu();

  // Note: Simple triggers like onOpen() run with restricted authorization mode
  // and cannot show sidebars automatically. Users should click Order Tools > Show Sidebar
  // from the menu to open the sidebar manually.
}
