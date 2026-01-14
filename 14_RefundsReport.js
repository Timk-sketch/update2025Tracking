// =====================================================
// 14_RefundsReport.js — Refunds Report
// Pulls all orders with refunds for the selected date range
// =====================================================

/**
 * Builds a Refunds report for the current date range.
 * Pulls from both Shopify and Squarespace orders with refunds.
 */
function buildRefundsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get date range from Marketing Controls (same as Summary Report)
  const startDate = PROPS.getProperty('ORDERS_SUMMARY_START_DATE');
  const endDate = PROPS.getProperty('ORDERS_SUMMARY_END_DATE');

  if (!startDate || !endDate) {
    throw new Error('Date range not set. Please set the date range in the sidebar first.');
  }

  const start = new Date(startDate);
  const end = new Date(endDate);

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
    'Quantity',
    'Original Price',
    'Refund Amount',
    'Discount Amount',
    'Net After Refund',
    'Order Status',
    'Refund Date/Modified'
  ];

  // Collect all refunded orders
  const refundedOrders = [];

  // Process Shopify orders
  const shopifySheet = ss.getSheetByName('Shopify Orders');
  if (shopifySheet) {
    const shopifyData = shopifySheet.getDataRange().getValues();
    if (shopifyData.length > 1) {
      const shopifyHeaders = shopifyData[0].map(h => String(h || '').trim());

      // Find columns
      const sOrderIdCol = shopifyHeaders.indexOf('Order ID');
      const sOrderNumCol = shopifyHeaders.indexOf('Order Number');
      const sCreatedCol = shopifyHeaders.indexOf('Created At');
      const sEmailCol = shopifyHeaders.indexOf('Customer Email');
      const sFirstNameCol = shopifyHeaders.indexOf('Customer First Name');
      const sLastNameCol = shopifyHeaders.indexOf('Customer Last Name');
      const sLineItemNameCol = shopifyHeaders.indexOf('Lineitem Name');
      const sLineItemQtyCol = shopifyHeaders.indexOf('Lineitem Quantity');
      const sLineItemPriceCol = shopifyHeaders.indexOf('Lineitem Price');
      const sRefundsCol = shopifyHeaders.indexOf('Total Refunds');
      const sDiscountsCol = shopifyHeaders.indexOf('Current Total Discounts');
      const sFinancialStatusCol = shopifyHeaders.indexOf('Financial Status');
      const sUpdatedCol = shopifyHeaders.indexOf('Updated At');

      // Process each row
      for (let r = 1; r < shopifyData.length; r++) {
        const row = shopifyData[r];

        // Get order date
        const orderDate = asDate_(row[sCreatedCol]);
        if (!orderDate || orderDate < start || orderDate > end) continue;

        // Check if there's a refund
        const refundAmount = parseFloat(row[sRefundsCol]) || 0;
        if (refundAmount <= 0) continue;

        // Extract data
        const lineItemPrice = parseFloat(row[sLineItemPriceCol]) || 0;
        const quantity = parseFloat(row[sLineItemQtyCol]) || 1;
        const originalPrice = lineItemPrice * quantity;
        const discountAmount = parseFloat(row[sDiscountsCol]) || 0;
        const netAfterRefund = originalPrice - refundAmount - discountAmount;

        const customerName = [
          row[sFirstNameCol] || '',
          row[sLastNameCol] || ''
        ].filter(Boolean).join(' ') || 'N/A';

        refundedOrders.push([
          'Shopify',
          row[sOrderIdCol] || '',
          row[sOrderNumCol] || '',
          orderDate,
          row[sEmailCol] || '',
          customerName,
          row[sLineItemNameCol] || '',
          quantity,
          originalPrice,
          refundAmount,
          discountAmount,
          netAfterRefund,
          row[sFinancialStatusCol] || '',
          row[sUpdatedCol] || ''
        ]);
      }
    }
  }

  // Process Squarespace orders
  const squarespaceSheet = ss.getSheetByName('Squarespace Orders');
  if (squarespaceSheet) {
    const squarespaceData = squarespaceSheet.getDataRange().getValues();
    if (squarespaceData.length > 1) {
      const squarespaceHeaders = squarespaceData[0].map(h => String(h || '').trim());

      // Find columns
      const sqOrderIdCol = squarespaceHeaders.indexOf('Order ID');
      const sqOrderNumCol = squarespaceHeaders.indexOf('Order Number');
      const sqCreatedCol = squarespaceHeaders.indexOf('Created On');
      const sqEmailCol = squarespaceHeaders.indexOf('Customer Email');
      const sqBillFirstCol = squarespaceHeaders.indexOf('Billing First Name');
      const sqBillLastCol = squarespaceHeaders.indexOf('Billing Last Name');
      const sqLineItemNameCol = squarespaceHeaders.indexOf('LineItem Product Name');
      const sqLineItemQtyCol = squarespaceHeaders.indexOf('LineItem Quantity');
      const sqLineItemPriceCol = squarespaceHeaders.indexOf('LineItem Unit Price Value');
      const sqRefundValueCol = squarespaceHeaders.indexOf('Refunded Total Value');
      const sqDiscountValueCol = squarespaceHeaders.indexOf('Discount Total Value');
      const sqFulfillmentCol = squarespaceHeaders.indexOf('Fulfillment Status');
      const sqModifiedCol = squarespaceHeaders.indexOf('Modified On');

      // Process each row
      for (let r = 1; r < squarespaceData.length; r++) {
        const row = squarespaceData[r];

        // Get order date
        const orderDate = asDate_(row[sqCreatedCol]);
        if (!orderDate || orderDate < start || orderDate > end) continue;

        // Check if there's a refund
        const refundAmount = parseFloat(row[sqRefundValueCol]) || 0;
        if (refundAmount <= 0) continue;

        // Extract data
        const lineItemPrice = parseFloat(row[sqLineItemPriceCol]) || 0;
        const quantity = parseFloat(row[sqLineItemQtyCol]) || 1;
        const originalPrice = lineItemPrice * quantity;
        const discountAmount = parseFloat(row[sqDiscountValueCol]) || 0;
        const netAfterRefund = originalPrice - refundAmount - discountAmount;

        const customerName = [
          row[sqBillFirstCol] || '',
          row[sqBillLastCol] || ''
        ].filter(Boolean).join(' ') || 'N/A';

        refundedOrders.push([
          'Squarespace',
          row[sqOrderIdCol] || '',
          row[sqOrderNumCol] || '',
          orderDate,
          row[sqEmailCol] || '',
          customerName,
          row[sqLineItemNameCol] || '',
          quantity,
          originalPrice,
          refundAmount,
          discountAmount,
          netAfterRefund,
          row[sqFulfillmentCol] || '',
          row[sqModifiedCol] || ''
        ]);
      }
    }
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
    // Date columns (D and N)
    refundsSheet.getRange(2, 4, refundedOrders.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    refundsSheet.getRange(2, 14, refundedOrders.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Currency columns (I, J, K, L)
    const currencyFormat = '"$"#,##0.00';
    refundsSheet.getRange(2, 9, refundedOrders.length, 4).setNumberFormat(currencyFormat);

    // Freeze header row
    refundsSheet.setFrozenRows(1);

    // Auto-resize columns
    for (let col = 1; col <= headers.length; col++) {
      refundsSheet.autoResizeColumn(col);
    }

    // Add summary at the bottom
    const summaryRow = refundedOrders.length + 3;
    refundsSheet.getRange(summaryRow, 1, 1, 2).merge().setValue('TOTAL REFUNDS:').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow, 3).setFormula(`=SUM(J2:J${refundedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    refundsSheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('TOTAL NET AFTER REFUND:').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow + 1, 3).setFormula(`=SUM(L2:L${refundedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    refundsSheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('NUMBER OF ORDERS:').setFontWeight('bold').setBackground('#f1f3f4');
    refundsSheet.getRange(summaryRow + 2, 3).setValue(refundedOrders.length).setFontWeight('bold').setBackground('#f1f3f4');
  } else {
    // No refunds found
    refundsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    refundsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    refundsSheet.getRange(2, 1).setValue('No refunds found for the selected date range.');
  }

  const msg = `Refunds Report built: ${refundedOrders.length} refunded orders found (${startDate} to ${endDate})`;
  logProgress('Refunds Report', msg);
  logImportEvent('Refunds Report', msg, refundedOrders.length);
  return msg;
}
