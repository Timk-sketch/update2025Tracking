// =====================================================
// 14_RefundsReport.js — Refunds Report
// Pulls all orders with refunds for the selected date range
// Uses All_Order_Clean sheet for accurate data
// =====================================================

/**
 * Builds a Refunds report for the current date range.
 * Pulls from All_Order_Clean with refunds.
 */
function buildRefundsReport() {
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

  logProgress('Refunds Report', `Building report for ${startDate} to ${endDate}...`);

  // Get or create the Refunds sheet
  let refundsSheet = ss.getSheetByName('Refunds Report');
  if (refundsSheet) {
    refundsSheet.clear();
  } else {
    refundsSheet = ss.insertSheet('Refunds Report');
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
    'Order Refund Total',
    'Order Net Revenue',
    'Currency',
    'Financial Status',
    'Fulfillment Status'
  ];

  // Collect all refunded orders from All_Order_Clean
  const refundedOrders = [];

  // Get All_Order_Clean sheet
  const cleanSheet = ss.getSheetByName(CLEAN_OUTPUT_SHEET || 'All_Order_Clean');
  if (!cleanSheet) {
    throw new Error('All_Order_Clean sheet not found. Please build the clean master first.');
  }

  const cleanData = cleanSheet.getDataRange().getValues();
  if (cleanData.length <= 1) {
    // No data, just headers
    refundsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    refundsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    refundsSheet.getRange(2, 1).setValue('No orders found in All_Order_Clean.');
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
  const colOrderRefundTotal = cleanHeaders.indexOf('order_refund_total');
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

    // Check if there's a refund
    const refundTotal = parseFloat(row[colOrderRefundTotal]) || 0;
    if (refundTotal <= 0) continue;

    // Extract data
    refundedOrders.push([
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
      parseFloat(row[colLineRevenue]) || 0,
      parseFloat(row[colOrderDiscountTotal]) || 0,
      refundTotal,
      parseFloat(row[colOrderNetRevenue]) || 0,
      row[colCurrency] || '',
      row[colFinancialStatus] || '',
      row[colFulfillmentStatus] || ''
    ]);
  }

  // Write to sheet
  if (refundedOrders.length > 0) {
    // Sort by order date (newest first)
    refundedOrders.sort((a, b) => b[3] - a[3]);

    // Write headers
    refundsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    refundsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

    // Write data
    refundsSheet.getRange(2, 1, refundedOrders.length, headers.length).setValues(refundedOrders);

    // Format columns
    // Date column (D)
    refundsSheet.getRange(2, 4, refundedOrders.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Number columns (I - Quantity)
    refundsSheet.getRange(2, 9, refundedOrders.length, 1).setNumberFormat('#,##0');

    // Currency columns (J, K, L, M, N - Unit Price, Line Revenue, Discount, Refund, Net Revenue)
    const currencyFormat = '"$"#,##0.00';
    refundsSheet.getRange(2, 10, refundedOrders.length, 5).setNumberFormat(currencyFormat);

    // Freeze header row
    refundsSheet.setFrozenRows(1);

    // Auto-resize columns
    for (let col = 1; col <= headers.length; col++) {
      refundsSheet.autoResizeColumn(col);
    }

    // Add summary at the bottom
    const summaryRow = refundedOrders.length + 3;
    refundsSheet.getRange(summaryRow, 1, 1, 2).merge().setValue('TOTAL REFUNDS:').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow, 3).setFormula(`=SUM(M2:M${refundedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    refundsSheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('TOTAL NET REVENUE (AFTER REFUNDS):').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow + 1, 3).setFormula(`=SUM(N2:N${refundedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    refundsSheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('TOTAL LINE REVENUE (BEFORE REFUNDS):').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow + 2, 3).setFormula(`=SUM(K2:K${refundedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    refundsSheet.getRange(summaryRow + 3, 1, 1, 2).merge().setValue('NUMBER OF LINE ITEMS:').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow + 3, 3).setValue(refundedOrders.length).setFontWeight('bold').setBackground('#f1f3f4');
  } else {
    // No refunds found
    refundsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    refundsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    refundsSheet.getRange(2, 1).setValue('No refunds found for the selected date range.');
  }

  const msg = `Refunds Report built: ${refundedOrders.length} refunded line items found (${startDate} to ${endDate})`;
  logProgress('Refunds Report', msg);
  logImportEvent('Refunds Report', msg, refundedOrders.length);
  return msg;
}
