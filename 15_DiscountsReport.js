// =====================================================
// 15_DiscountsReport.js — Discounts Report
// Pulls all orders with discounts for the selected date range
// Uses All_Order_Clean sheet for accurate data
// =====================================================

/**
 * Builds a Discounts report for the current date range.
 * Pulls from All_Order_Clean with discounts.
 */
function buildDiscountsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get date range from Orders Summary Report sheet (same as Summary Report)
  const outSheet = ss.getSheetByName('Orders_Summary_Report');
  if (!outSheet) {
    throw new Error('Orders_Summary_Report sheet not found. Please set the date range in the sidebar first.');
  }

  const startRaw = outSheet.getRange('B2').getValue();
  const endRaw = outSheet.getRange('D2').getValue();

  const start = asDate_(startRaw);
  const end = asDate_(endRaw);

  if (!start || !end) {
    throw new Error('Date range not set. Please set the date range in the sidebar first.');
  }

  const startDate = formatDate_(start);
  const endDate = formatDate_(end);

  logProgress('Discounts Report', `Building report for ${startDate} to ${endDate}...`);

  // Get or create the Discounts sheet
  let discountsSheet = ss.getSheetByName('Discounts Report');
  if (discountsSheet) {
    discountsSheet.clear();
  } else {
    discountsSheet = ss.insertSheet('Discounts Report');
  }

  // Define headers
  const headers = [
    'Platform',
    'Order ID',
    'Order Number',
    'Order Date',
    'Customer Email',
    'Customer Name',
    'Product Name',
    'SKU',
    'Quantity',
    'Unit Price',
    'Line Revenue',
    'Order Discount Total',
    'Order Net Revenue',
    'Discount %',
    'Currency',
    'Financial Status',
    'Fulfillment Status'
  ];

  // Collect all discounted orders from All_Order_Clean
  const discountedOrders = [];

  // Get All_Order_Clean sheet
  const cleanSheet = ss.getSheetByName(CLEAN_OUTPUT_SHEET || 'All_Order_Clean');
  if (!cleanSheet) {
    throw new Error('All_Order_Clean sheet not found. Please build the clean master first.');
  }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 1) {
    // No data, just headers
    discountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    discountsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    discountsSheet.getRange(2, 1).setValue('No orders found in All_Order_Clean.');
    return 'No orders found in All_Order_Clean.';
  }

  const cleanHeaders = cleanData[0].map(h => String(h || '').trim());

  // Find column indices
  const colPlatform = cleanHeaders.indexOf('platform');
  const colOrderId = cleanHeaders.indexOf('order_id');
  const colOrderNumber = cleanHeaders.indexOf('order_number');
  const colOrderDate = cleanHeaders.indexOf('order_date');
  const colEmailRaw = cleanHeaders.indexOf('customer_email_raw');
  const colCustomerName = cleanHeaders.indexOf('customer_name');
  const colProductName = cleanHeaders.indexOf('product_name');
  const colSku = cleanHeaders.indexOf('sku');
  const colQuantity = cleanHeaders.indexOf('quantity');
  const colUnitPrice = cleanHeaders.indexOf('unit_price');
  const colLineRevenue = cleanHeaders.indexOf('line_revenue');
  const colOrderDiscountTotal = cleanHeaders.indexOf('order_discount_total');
  const colOrderNetRevenue = cleanHeaders.indexOf('order_net_revenue');
  const colCurrency = cleanHeaders.indexOf('currency');
  const colFinancialStatus = cleanHeaders.indexOf('financial_status');
  const colFulfillmentStatus = cleanHeaders.indexOf('fulfillment_status');

  // Process each row
  for (let r = 1; r < cleanData.length; r++) {
    const row = cleanData[r];

    // Get order date
    const orderDate = asDate_(row[colOrderDate]);
    if (!orderDate || orderDate < start || orderDate > end) continue;

    // Check if there's a discount
    const discountTotal = parseFloat(row[colOrderDiscountTotal]) || 0;
    if (discountTotal <= 0) continue;

    // Extract data
    const lineRevenue = parseFloat(row[colLineRevenue]) || 0;
    const discountPercent = lineRevenue > 0 ? (discountTotal / lineRevenue) * 100 : 0;

    discountedOrders.push([
      row[colPlatform] || '',
      row[colOrderId] || '',
      row[colOrderNumber] || '',
      orderDate,
      row[colEmailRaw] || '',
      row[colCustomerName] || '',
      row[colProductName] || '',
      row[colSku] || '',
      parseFloat(row[colQuantity]) || 0,
      parseFloat(row[colUnitPrice]) || 0,
      lineRevenue,
      discountTotal,
      parseFloat(row[colOrderNetRevenue]) || 0,
      discountPercent,
      row[colCurrency] || '',
      row[colFinancialStatus] || '',
      row[colFulfillmentStatus] || ''
    ]);
  }

  // Write to sheet
  if (discountedOrders.length > 0) {
    // Sort by discount % (highest first) to identify excessive discounting
    discountedOrders.sort((a, b) => b[13] - a[13]); // Column N = index 13 (Discount %)

    // Write headers
    discountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    discountsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

    // Write data
    discountsSheet.getRange(2, 1, discountedOrders.length, headers.length).setValues(discountedOrders);

    // Format columns
    // Date column (D)
    discountsSheet.getRange(2, 4, discountedOrders.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Number columns (I - Quantity)
    discountsSheet.getRange(2, 9, discountedOrders.length, 1).setNumberFormat('#,##0');

    // Currency columns (J, K, L, M - Unit Price, Line Revenue, Discount, Net Revenue)
    const currencyFormat = '"$"#,##0.00';
    discountsSheet.getRange(2, 10, discountedOrders.length, 4).setNumberFormat(currencyFormat);

    // Percent column (N - Discount %)
    discountsSheet.getRange(2, 14, discountedOrders.length, 1).setNumberFormat('0.0"%"');

    // Freeze header row
    discountsSheet.setFrozenRows(1);

    // Auto-resize columns
    for (let col = 1; col <= headers.length; col++) {
      discountsSheet.autoResizeColumn(col);
    }

    // Add summary at the bottom
    const summaryRow = discountedOrders.length + 3;
    discountsSheet.getRange(summaryRow, 1, 1, 2).merge().setValue('TOTAL DISCOUNTS:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow, 3).setFormula(`=SUM(L2:L${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    discountsSheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('TOTAL LINE REVENUE:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 1, 3).setFormula(`=SUM(K2:K${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    discountsSheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('TOTAL NET REVENUE:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 2, 3).setFormula(`=SUM(M2:M${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    discountsSheet.getRange(summaryRow + 3, 1, 1, 2).merge().setValue('AVERAGE DISCOUNT %:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 3, 3).setFormula(`=AVERAGE(N2:N${discountedOrders.length + 1})`).setNumberFormat('0.0"%"').setFontWeight('bold').setBackground('#f1f3f4');

    discountsSheet.getRange(summaryRow + 4, 1, 1, 2).merge().setValue('REVENUE LOST TO DISCOUNTS (%):').setFontWeight('bold').setBackground('#fce8e6');
    discountsSheet.getRange(summaryRow + 4, 3).setFormula(`=C${summaryRow}/C${summaryRow+1}`).setNumberFormat('0.0"%"').setFontWeight('bold').setBackground('#fce8e6');

    discountsSheet.getRange(summaryRow + 5, 1, 1, 2).merge().setValue('NUMBER OF LINE ITEMS:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 5, 3).setValue(discountedOrders.length).setFontWeight('bold').setBackground('#f1f3f4');

    // Add insights section
    const insightsRow = summaryRow + 7;
    discountsSheet.getRange(insightsRow, 1).setValue('💡 INSIGHTS:').setFontWeight('bold').setFontSize(12);
    discountsSheet.getRange(insightsRow + 1, 1, 1, 3).merge().setValue('• Orders sorted by discount % (highest first) - review top rows for excessive discounting');
    discountsSheet.getRange(insightsRow + 2, 1, 1, 3).merge().setValue('• High discount % may indicate: pricing issues, sales training gaps, or competitive pressure');
    discountsSheet.getRange(insightsRow + 3, 1, 1, 3).merge().setValue('• Use this data to evaluate: Are discounts necessary to close sales, or is better training needed?');
  } else {
    // No discounts found
    discountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    discountsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    discountsSheet.getRange(2, 1).setValue('No discounts found for the selected date range.');
  }

  const msg = `Discounts Report built: ${discountedOrders.length} discounted line items found (${startDate} to ${endDate})`;
  logProgress('Discounts Report', msg);
  logImportEvent('Discounts Report', msg, discountedOrders.length);
  return msg;
}
