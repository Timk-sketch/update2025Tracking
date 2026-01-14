// =====================================================
// 15_DiscountsReport.js — Discounts Report
// Pulls all orders with discounts for the selected date range
// =====================================================

/**
 * Builds a Discounts report for the current date range.
 * Pulls from both Shopify and Squarespace orders with discounts.
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
    'Quantity',
    'Unit Price',
    'Subtotal',
    'Discount Amount',
    'Discount Code/Type',
    'Net After Discount',
    'Discount %',
    'Order Status'
  ];

  // Collect all discounted orders
  const discountedOrders = [];

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
      const sDiscountsCol = shopifyHeaders.indexOf('Current Total Discounts');
      const sDiscountCodesCol = shopifyHeaders.indexOf('Discount Codes');
      const sFinancialStatusCol = shopifyHeaders.indexOf('Financial Status');

      // Process each row
      for (let r = 1; r < shopifyData.length; r++) {
        const row = shopifyData[r];

        // Get order date
        const orderDate = asDate_(row[sCreatedCol]);
        if (!orderDate || orderDate < start || orderDate > end) continue;

        // Check if there's a discount
        const discountAmount = parseFloat(row[sDiscountsCol]) || 0;
        if (discountAmount <= 0) continue;

        // Extract data
        const lineItemPrice = parseFloat(row[sLineItemPriceCol]) || 0;
        const quantity = parseFloat(row[sLineItemQtyCol]) || 1;
        const subtotal = lineItemPrice * quantity;
        const netAfterDiscount = subtotal - discountAmount;
        const discountPercent = subtotal > 0 ? (discountAmount / subtotal * 100) : 0;

        const customerName = [
          row[sFirstNameCol] || '',
          row[sLastNameCol] || ''
        ].filter(Boolean).join(' ') || 'N/A';

        const discountCode = row[sDiscountCodesCol] || 'Manual Discount';

        discountedOrders.push([
          'Shopify',
          row[sOrderIdCol] || '',
          row[sOrderNumCol] || '',
          orderDate,
          row[sEmailCol] || '',
          customerName,
          row[sLineItemNameCol] || '',
          quantity,
          lineItemPrice,
          subtotal,
          discountAmount,
          discountCode,
          netAfterDiscount,
          discountPercent,
          row[sFinancialStatusCol] || ''
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
      const sqDiscountValueCol = squarespaceHeaders.indexOf('Discount Total Value');
      const sqDiscountLinesCol = squarespaceHeaders.indexOf('Discount Lines');
      const sqFulfillmentCol = squarespaceHeaders.indexOf('Fulfillment Status');

      // Process each row
      for (let r = 1; r < squarespaceData.length; r++) {
        const row = squarespaceData[r];

        // Get order date
        const orderDate = asDate_(row[sqCreatedCol]);
        if (!orderDate || orderDate < start || orderDate > end) continue;

        // Check if there's a discount
        const discountAmount = parseFloat(row[sqDiscountValueCol]) || 0;
        if (discountAmount <= 0) continue;

        // Extract data
        const lineItemPrice = parseFloat(row[sqLineItemPriceCol]) || 0;
        const quantity = parseFloat(row[sqLineItemQtyCol]) || 1;
        const subtotal = lineItemPrice * quantity;
        const netAfterDiscount = subtotal - discountAmount;
        const discountPercent = subtotal > 0 ? (discountAmount / subtotal * 100) : 0;

        const customerName = [
          row[sqBillFirstCol] || '',
          row[sqBillLastCol] || ''
        ].filter(Boolean).join(' ') || 'N/A';

        const discountCode = row[sqDiscountLinesCol] || 'Manual Discount';

        discountedOrders.push([
          'Squarespace',
          row[sqOrderIdCol] || '',
          row[sqOrderNumCol] || '',
          orderDate,
          row[sqEmailCol] || '',
          customerName,
          row[sqLineItemNameCol] || '',
          quantity,
          lineItemPrice,
          subtotal,
          discountAmount,
          discountCode,
          netAfterDiscount,
          discountPercent,
          row[sqFulfillmentCol] || ''
        ]);
      }
    }
  }

  // Write to sheet
  if (discountedOrders.length > 0) {
    // Sort by discount percent (highest first) to identify biggest discounts
    discountedOrders.sort((a, b) => b[13] - a[13]);

    // Write headers
    discountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    discountsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

    // Write data
    discountsSheet.getRange(2, 1, discountedOrders.length, headers.length).setValues(discountedOrders);

    // Format columns
    // Date column (D)
    discountsSheet.getRange(2, 4, discountedOrders.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Currency columns (I, J, K, M)
    const currencyFormat = '"$"#,##0.00';
    discountsSheet.getRange(2, 9, discountedOrders.length, 1).setNumberFormat(currencyFormat); // Unit Price
    discountsSheet.getRange(2, 10, discountedOrders.length, 1).setNumberFormat(currencyFormat); // Subtotal
    discountsSheet.getRange(2, 11, discountedOrders.length, 1).setNumberFormat(currencyFormat); // Discount Amount
    discountsSheet.getRange(2, 13, discountedOrders.length, 1).setNumberFormat(currencyFormat); // Net After Discount

    // Percent column (N)
    discountsSheet.getRange(2, 14, discountedOrders.length, 1).setNumberFormat('0.0"%"');

    // Freeze header row
    discountsSheet.setFrozenRows(1);

    // Auto-resize columns
    for (let col = 1; col <= headers.length; col++) {
      discountsSheet.autoResizeColumn(col);
    }

    // Add summary and analysis at the bottom
    const summaryRow = discountedOrders.length + 3;

    // Total discounts given
    discountsSheet.getRange(summaryRow, 1, 1, 2).merge().setValue('TOTAL DISCOUNTS GIVEN:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow, 3).setFormula(`=SUM(K2:K${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    // Average discount percent
    discountsSheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('AVERAGE DISCOUNT %:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 1, 3).setFormula(`=AVERAGE(N2:N${discountedOrders.length + 1})`).setNumberFormat('0.0"%"').setFontWeight('bold').setBackground('#f1f3f4');

    // Number of discounted orders
    discountsSheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('NUMBER OF ORDERS:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 2, 3).setValue(discountedOrders.length).setFontWeight('bold').setBackground('#f1f3f4');

    // Total potential revenue (before discount)
    discountsSheet.getRange(summaryRow + 3, 1, 1, 2).merge().setValue('TOTAL POTENTIAL REVENUE:').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 3, 3).setFormula(`=SUM(J2:J${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    // Actual revenue (after discount)
    discountsSheet.getRange(summaryRow + 4, 1, 1, 2).merge().setValue('ACTUAL REVENUE (AFTER DISCOUNT):').setFontWeight('bold').setBackground('#f1f3f4');
    discountsSheet.getRange(summaryRow + 4, 3).setFormula(`=SUM(M2:M${discountedOrders.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

    // Revenue lost to discounts (as %)
    discountsSheet.getRange(summaryRow + 5, 1, 1, 2).merge().setValue('REVENUE LOST TO DISCOUNTS (%):').setFontWeight('bold').setBackground('#fce8e6');
    discountsSheet.getRange(summaryRow + 5, 3).setFormula(`=C${summaryRow}/C${summaryRow+3}`).setNumberFormat('0.0"%"').setFontWeight('bold').setBackground('#fce8e6');

    // Add insights section
    const insightsRow = summaryRow + 7;
    discountsSheet.getRange(insightsRow, 1).setValue('💡 INSIGHTS:').setFontWeight('bold').setFontSize(12);
    discountsSheet.getRange(insightsRow + 1, 1, 1, 3).merge().setValue('• Orders are sorted by discount % (highest first) - review top rows for excessive discounting');
    discountsSheet.getRange(insightsRow + 2, 1, 1, 3).merge().setValue('• Check "Discount Code/Type" column to identify which discounts are used most');
    discountsSheet.getRange(insightsRow + 3, 1, 1, 3).merge().setValue('• High discount % may indicate: pricing issues, sales training gaps, or competitive pressure');
    discountsSheet.getRange(insightsRow + 4, 1, 1, 3).merge().setValue('• Compare actual vs potential revenue to quantify impact of discounting strategy');

  } else {
    // No discounts found
    discountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    discountsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    discountsSheet.getRange(2, 1).setValue('No discounts found for the selected date range.');
  }

  const msg = `Discounts Report built: ${discountedOrders.length} discounted orders found (${startDate} to ${endDate})`;
  logProgress('Discounts Report', msg);
  logImportEvent('Discounts Report', msg, discountedOrders.length);
  return msg;
}
