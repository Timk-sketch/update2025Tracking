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

/**
 * Adds a Shopify API comparison sheet to the Refunds Report.
 * Shows refunds directly from Shopify API vs what's in your system.
 * Helps identify discrepancies in refund data.
 */
function addShopifyRefundComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get date range from Orders Summary Report sheet
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

  logProgress('Refund Comparison', `Fetching Shopify API refunds for ${startDate} to ${endDate}...`);

  // Create or clear comparison sheet
  let compSheet = ss.getSheetByName('Refunds_Shopify_API_Comparison');
  if (compSheet) {
    compSheet.clear();
  } else {
    compSheet = ss.insertSheet('Refunds_Shopify_API_Comparison');
  }

  // Fetch refunds directly from Shopify API
  const apiKey = PROPS.getProperty('SHOPIFY_API_KEY');
  const shopDomain = PROPS.getProperty('SHOPIFY_SHOP_DOMAIN');
  const apiVersion = '2023-10';

  if (!apiKey || !shopDomain) {
    throw new Error("Missing SHOPIFY_API_KEY or SHOPIFY_SHOP_DOMAIN in Script Properties.");
  }

  // Calculate days back from start date
  const today = new Date();
  const daysBack = Math.ceil((today - start) / (1000 * 60 * 60 * 24));
  const d = new Date();
  d.setDate(d.getDate() - daysBack);
  const createdAtMin = d.toISOString();

  const apiRefunds = [];
  const financialStatuses = ['refunded', 'partially_refunded'];

  for (const status of financialStatuses) {
    let url = `https://${shopDomain}/admin/api/${apiVersion}/orders.json?status=any&financial_status=${status}&limit=250&created_at_min=${encodeURIComponent(createdAtMin)}`;

    while (url) {
      const resp = fetchWithRetry_(url, {
        method: "get",
        headers: { "X-Shopify-Access-Token": apiKey },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() !== 200) {
        throw new Error(`Shopify API error (${resp.getResponseCode()}): ${resp.getContentText()}`);
      }

      const json = JSON.parse(resp.getContentText());
      const orders = json.orders || [];
      if (!orders.length) break;

      orders.forEach(order => {
        const orderDate = asDate_(order.created_at);
        if (!orderDate || orderDate < start || orderDate > end) return;

        const refundTotal = computeShopifyRefundTotal_(order);
        if (refundTotal <= 0) return;

        apiRefunds.push({
          orderId: order.id || '',
          orderNumber: order.order_number || '',
          orderDate: orderDate,
          email: order.email || '',
          customerName: `${order.customer?.first_name || ''} ${order.customer?.last_name || ''}`.trim(),
          totalPrice: parseFloat(order.total_price) || 0,
          refundTotal: refundTotal,
          financialStatus: order.financial_status || '',
          refundDetails: (order.refunds || []).map(ref => {
            const refDate = ref.created_at || '';
            const refAmount = (ref.transactions || [])
              .filter(t => !t.kind || t.kind === 'refund')
              .reduce((sum, t) => sum + Math.abs(parseFloat(t.amount) || 0), 0);
            return `${refDate}: $${refAmount.toFixed(2)}`;
          }).join(' | ')
        });
      });

      // Handle pagination
      const linkHeader = resp.getHeaders()['Link'] || resp.getHeaders()['link'] || '';
      const nextMatch = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
      url = nextMatch ? nextMatch[1] : null;

      if (url) Utilities.sleep(500); // Rate limit protection
    }
  }

  // Get your system's refund data from All_Order_Clean
  const cleanSheet = ss.getSheetByName(CLEAN_OUTPUT_SHEET || 'All_Order_Clean');
  const systemRefunds = new Map(); // orderId -> {refund, orderNumber, email, etc}

  if (cleanSheet && cleanSheet.getLastRow() > 1) {
    const cleanData = cleanSheet.getDataRange().getValues();
    const cleanHeaders = cleanData[0].map(h => String(h || '').trim());

    const colOrderId = cleanHeaders.indexOf('order_id');
    const colOrderNumber = cleanHeaders.indexOf('order_number');
    const colOrderDate = cleanHeaders.indexOf('order_date');
    const colEmail = cleanHeaders.indexOf('customer_email_raw');
    const colCustomerName = cleanHeaders.indexOf('customer_name');
    const colOrderRefund = cleanHeaders.indexOf('order_refund_total');
    const colPlatform = cleanHeaders.indexOf('platform');

    for (let r = 1; r < cleanData.length; r++) {
      const row = cleanData[r];
      const platform = row[colPlatform];
      if (platform !== 'Shopify') continue;

      const orderDate = asDate_(row[colOrderDate]);
      if (!orderDate || orderDate < start || orderDate > end) continue;

      const orderId = String(row[colOrderId] || '');
      const refund = parseFloat(row[colOrderRefund]) || 0;

      if (refund > 0 && orderId && !systemRefunds.has(orderId)) {
        systemRefunds.set(orderId, {
          orderNumber: row[colOrderNumber] || '',
          orderDate: orderDate,
          email: row[colEmail] || '',
          customerName: row[colCustomerName] || '',
          refund: refund
        });
      }
    }
  }

  // Build comparison data
  const headers = [
    'Order ID',
    'Order #',
    'Order Date',
    'Customer Email',
    'Customer Name',
    'Shopify API Refund',
    'Your System Refund',
    'Difference',
    'Status',
    'Refund Details (API)'
  ];

  const comparisonRows = [];
  const allOrderIds = new Set([...apiRefunds.map(r => String(r.orderId)), ...Array.from(systemRefunds.keys())]);

  allOrderIds.forEach(orderId => {
    const apiRefund = apiRefunds.find(r => String(r.orderId) === orderId);
    const sysRefund = systemRefunds.get(orderId);

    const apiAmount = apiRefund ? apiRefund.refundTotal : 0;
    const sysAmount = sysRefund ? sysRefund.refund : 0;
    const diff = apiAmount - sysAmount;

    let status = '';
    if (Math.abs(diff) < 0.01) {
      status = '✓ Match';
    } else if (!apiRefund) {
      status = '⚠ In System Only';
    } else if (!sysRefund) {
      status = '⚠ In API Only';
    } else {
      status = `✗ Mismatch ($${Math.abs(diff).toFixed(2)})`;
    }

    comparisonRows.push([
      orderId,
      apiRefund ? apiRefund.orderNumber : (sysRefund ? sysRefund.orderNumber : ''),
      apiRefund ? apiRefund.orderDate : (sysRefund ? sysRefund.orderDate : ''),
      apiRefund ? apiRefund.email : (sysRefund ? sysRefund.email : ''),
      apiRefund ? apiRefund.customerName : (sysRefund ? sysRefund.customerName : ''),
      apiAmount,
      sysAmount,
      diff,
      status,
      apiRefund ? apiRefund.refundDetails : ''
    ]);
  });

  // Sort by difference (largest first)
  comparisonRows.sort((a, b) => Math.abs(b[7]) - Math.abs(a[7]));

  // Currency format (needed for summary section too)
  const currencyFormat = '"$"#,##0.00';

  // Write to sheet
  compSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  compSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff');

  if (comparisonRows.length > 0) {
    compSheet.getRange(2, 1, comparisonRows.length, headers.length).setValues(comparisonRows);

    // Format columns
    compSheet.getRange(2, 3, comparisonRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Date
    compSheet.getRange(2, 6, comparisonRows.length, 3).setNumberFormat(currencyFormat); // Refund columns

    // Color code status column
    for (let r = 0; r < comparisonRows.length; r++) {
      const statusCell = compSheet.getRange(r + 2, 9);
      const status = comparisonRows[r][8];
      if (status.startsWith('✓')) {
        statusCell.setBackground('#d9ead3'); // Light green
      } else if (status.startsWith('✗')) {
        statusCell.setBackground('#fce8e6'); // Light red
      } else {
        statusCell.setBackground('#fff3cd'); // Light yellow
      }
    }
  }

  compSheet.setFrozenRows(1);
  for (let col = 1; col <= headers.length; col++) {
    compSheet.autoResizeColumn(col);
  }

  // Add summary
  const summaryRow = comparisonRows.length + 3;
  compSheet.getRange(summaryRow, 1, 1, 2).merge().setValue('SHOPIFY API TOTAL:').setFontWeight('bold').setBackground('#f1f3f4');
  compSheet.getRange(summaryRow, 3).setFormula(`=SUM(F2:F${comparisonRows.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

  compSheet.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('YOUR SYSTEM TOTAL:').setFontWeight('bold').setBackground('#f1f3f4');
  compSheet.getRange(summaryRow + 1, 3).setFormula(`=SUM(G2:G${comparisonRows.length + 1})`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#f1f3f4');

  compSheet.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('DIFFERENCE:').setFontWeight('bold').setBackground('#fce8e6');
  compSheet.getRange(summaryRow + 2, 3).setFormula(`=C${summaryRow}-C${summaryRow+1}`).setNumberFormat(currencyFormat).setFontWeight('bold').setBackground('#fce8e6');

  const matchCount = comparisonRows.filter(r => r[8].startsWith('✓')).length;
  const mismatchCount = comparisonRows.filter(r => r[8].startsWith('✗')).length;

  compSheet.getRange(summaryRow + 4, 1, 1, 2).merge().setValue('MATCHES:').setFontWeight('bold');
  compSheet.getRange(summaryRow + 4, 3).setValue(matchCount);

  compSheet.getRange(summaryRow + 5, 1, 1, 2).merge().setValue('MISMATCHES:').setFontWeight('bold');
  compSheet.getRange(summaryRow + 5, 3).setValue(mismatchCount).setFontColor(mismatchCount > 0 ? '#a50e0e' : '#000000');

  const msg = `Shopify API comparison added: ${comparisonRows.length} orders checked, ${mismatchCount} mismatches found`;
  logProgress('Refund Comparison', msg);
  logImportEvent('Refund Comparison', msg, comparisonRows.length);

  return msg;
}
