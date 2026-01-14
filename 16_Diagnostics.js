// =====================================================
// 16_Diagnostics.js — Diagnostic functions to troubleshoot data issues
// =====================================================

/**
 * Check what date range of data you have imported in Shopify and Squarespace sheets
 * This helps identify if you're missing historical data
 */
function diagnosticCheckDataCoverage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shopifySheet = ss.getSheetByName('Shopify Orders');
  const squarespaceSheet = ss.getSheetByName('Squarespace Orders');
  const cleanSheet = ss.getSheetByName('All_Orders_Clean');

  let report = '=== DATA COVERAGE DIAGNOSTIC ===\n\n';

  // Shopify Orders
  if (shopifySheet) {
    const data = shopifySheet.getDataRange().getValues();
    const headers = data[0];
    const createdCol = headers.indexOf('Created At');

    if (data.length > 1 && createdCol >= 0) {
      const dates = [];
      for (let i = 1; i < data.length; i++) {
        const d = new Date(data[i][createdCol]);
        if (!isNaN(d.getTime())) dates.push(d);
      }
      dates.sort((a, b) => a - b);

      report += `SHOPIFY ORDERS:\n`;
      report += `  Total Rows: ${data.length - 1}\n`;
      report += `  Oldest Order: ${dates[0] ? dates[0].toISOString().split('T')[0] : 'N/A'}\n`;
      report += `  Newest Order: ${dates[dates.length-1] ? dates[dates.length-1].toISOString().split('T')[0] : 'N/A'}\n\n`;
    } else {
      report += `SHOPIFY ORDERS: No data or missing Created At column\n\n`;
    }
  } else {
    report += `SHOPIFY ORDERS: Sheet not found\n\n`;
  }

  // Squarespace Orders
  if (squarespaceSheet) {
    const data = squarespaceSheet.getDataRange().getValues();
    const headers = data[0];
    const createdCol = headers.indexOf('Created On');

    if (data.length > 1 && createdCol >= 0) {
      const dates = [];
      for (let i = 1; i < data.length; i++) {
        const d = new Date(data[i][createdCol]);
        if (!isNaN(d.getTime())) dates.push(d);
      }
      dates.sort((a, b) => a - b);

      report += `SQUARESPACE ORDERS:\n`;
      report += `  Total Rows: ${data.length - 1}\n`;
      report += `  Oldest Order: ${dates[0] ? dates[0].toISOString().split('T')[0] : 'N/A'}\n`;
      report += `  Newest Order: ${dates[dates.length-1] ? dates[dates.length-1].toISOString().split('T')[0] : 'N/A'}\n\n`;
    } else {
      report += `SQUARESPACE ORDERS: No data or missing Created On column\n\n`;
    }
  } else {
    report += `SQUARESPACE ORDERS: Sheet not found\n\n`;
  }

  // All Orders Clean
  if (cleanSheet) {
    const data = cleanSheet.getDataRange().getValues();
    report += `ALL_ORDERS_CLEAN:\n`;
    report += `  Total Rows: ${data.length - 1}\n\n`;
  } else {
    report += `ALL_ORDERS_CLEAN: Sheet not found\n\n`;
  }

  report += '=== RECOMMENDATION ===\n';
  report += 'If oldest order is recent (within last 14 days), you need a\n';
  report += 'one-time historical import to get full year data.\n';
  report += 'Contact your developer to import historical orders.\n';

  Logger.log(report);
  ss.toast(report, 'Data Coverage Diagnostic', 15);
  return report;
}

/**
 * Count excluded orders from All_Orders_Clean to understand what's being filtered out
 */
function diagnosticCheckExcludedOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shopifySheet = ss.getSheetByName('Shopify Orders');
  const cleanSheet = ss.getSheetByName('All_Orders_Clean');

  if (!shopifySheet || !cleanSheet) {
    ss.toast('Missing required sheets', 'Error', 5);
    return;
  }

  const shopifyData = shopifySheet.getDataRange().getValues();
  const cleanData = cleanSheet.getDataRange().getValues();

  let report = '=== EXCLUSION DIAGNOSTIC ===\n\n';
  report += `Shopify Orders (raw): ${shopifyData.length - 1} rows\n`;
  report += `All_Orders_Clean: ${cleanData.length - 1} rows\n`;
  report += `Difference: ${(shopifyData.length - 1) - (cleanData.length - 1)} rows excluded\n\n`;

  report += 'Rows may be excluded due to:\n';
  report += '- Test orders (test = true)\n';
  report += '- Banned customer emails\n';
  report += '- Missing product name\n';
  report += '- Empty rows\n';

  Logger.log(report);
  ss.toast(report, 'Exclusion Diagnostic', 10);
  return report;
}
