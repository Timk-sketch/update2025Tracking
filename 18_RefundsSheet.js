// =====================================================
// 18_RefundsSheet.js — Separate Refunds Sheet Management
// Maintains a complete log of all refunds with issue dates
// =====================================================

const SHOPIFY_REFUNDS_SHEET_NAME = 'Shopify_Refunds';
const SQUARESPACE_REFUNDS_SHEET_NAME = 'Squarespace_Refunds';

const SHOPIFY_REFUNDS_HEADERS = [
  'Refund ID',
  'Order ID',
  'Order Number',
  'Order Date',
  'Refund Date',
  'Refund Amount',
  'Customer Email',
  'Customer Name',
  'Note',
  'Created At'
];

const SQUARESPACE_REFUNDS_HEADERS = [
  'Order ID',
  'Order Number',
  'Order Date',
  'Refund Date',
  'Refund Amount',
  'Customer Email',
  'Customer Name',
  'Note',
  'Created At'
];

/**
 * Imports all Shopify refunds from the last N days (default 30).
 * Only adds NEW refunds (checks for existing Refund IDs to skip duplicates).
 * The sheet is permanent history (never cleared) - reports query by date range.
 * We only query recent refunds to catch new ones added since last import.
 * @param {number} days - Number of days to look back (default 30, use 180 for full historical import)
 */
function importShopifyRefunds(days) {
  days = days || 30; // Default to 30 days if not specified
  const apiKey = PROPS.getProperty('SHOPIFY_API_KEY');
  const shopDomain = PROPS.getProperty('SHOPIFY_SHOP_DOMAIN');
  const apiVersion = '2023-10';

  if (!apiKey || !shopDomain) {
    throw new Error('Missing SHOPIFY_API_KEY or SHOPIFY_SHOP_DOMAIN in Script Properties.');
  }

  logProgress('Shopify Refunds', 'Fetching refunds from API...');

  // Get or create refunds sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refundsSheet = getOrCreateSheetWithHeaders(SHOPIFY_REFUNDS_SHEET_NAME, SHOPIFY_REFUNDS_HEADERS);

  // Load existing refund IDs to avoid duplicates
  const existingRefundIds = new Set();
  const lastRow = refundsSheet.getLastRow();
  if (lastRow > 1) {
    const existingData = refundsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(row => {
      if (row[0]) existingRefundIds.add(String(row[0]));
    });
  }

  // Query orders with refunds from last N days
  // (We only need recent refunds since we run this regularly and the sheet keeps full history)
  const d = new Date();
  d.setDate(d.getDate() - days);
  const updatedAtMin = d.toISOString();

  const newRefunds = [];
  const financialStatuses = ['refunded', 'partially_refunded'];

  for (const status of financialStatuses) {
    let url = `https://${shopDomain}/admin/api/${apiVersion}/orders.json?status=any&financial_status=${status}&limit=250&updated_at_min=${encodeURIComponent(updatedAtMin)}`;

    while (url) {
      const resp = fetchWithRetry_(url, {
        method: 'get',
        headers: { 'X-Shopify-Access-Token': apiKey },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() !== 200) {
        throw new Error(`Shopify API error (${resp.getResponseCode()}): ${resp.getContentText()}`);
      }

      const json = JSON.parse(resp.getContentText());
      const orders = json.orders || [];
      if (!orders.length) break;

      orders.forEach(order => {
        if (!order.refunds || !order.refunds.length) return;

        const orderId = String(order.id);
        const orderNumber = order.order_number || '';
        const orderDate = asDate_(order.created_at);
        const customerEmail = order.email || '';
        const customerName = `${order.customer?.first_name || ''} ${order.customer?.last_name || ''}`.trim();

        order.refunds.forEach(refund => {
          const refundId = String(refund.id);

          // Skip if we already have this refund
          if (existingRefundIds.has(refundId)) return;

          const refundDate = asDate_(refund.created_at);
          const note = refund.note || '';

          // Calculate refund amount from transactions
          const refundAmount = (refund.transactions || [])
            .filter(t => !t.kind || t.kind === 'refund')
            .reduce((sum, t) => sum + Math.abs(parseFloat(t.amount) || 0), 0);

          if (refundAmount > 0) {
            newRefunds.push([
              refundId,
              orderId,
              orderNumber,
              orderDate,
              refundDate,
              refundAmount,
              customerEmail,
              customerName,
              note,
              new Date()
            ]);
          }
        });
      });

      // Handle pagination
      const linkHeader = resp.getHeaders()['Link'] || resp.getHeaders()['link'] || '';
      const nextMatch = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
      url = nextMatch ? nextMatch[1] : null;
      if (url) Utilities.sleep(500);
    }
  }

  // Append new refunds to sheet
  if (newRefunds.length > 0) {
    refundsSheet.getRange(lastRow + 1, 1, newRefunds.length, SHOPIFY_REFUNDS_HEADERS.length).setValues(newRefunds);

    // Format columns
    const newStartRow = lastRow + 1;
    refundsSheet.getRange(newStartRow, 4, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Order Date
    refundsSheet.getRange(newStartRow, 5, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Refund Date
    refundsSheet.getRange(newStartRow, 6, newRefunds.length, 1).setNumberFormat('"$"#,##0.00'); // Refund Amount
    refundsSheet.getRange(newStartRow, 10, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Created At

    refundsSheet.setFrozenRows(1);
    refundsSheet.getRange(1, 1, 1, SHOPIFY_REFUNDS_HEADERS.length).setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
  }

  const msg = `✅ Imported ${newRefunds.length} new Shopify refunds (${existingRefundIds.size} existing)`;
  logProgress('Shopify Refunds', msg);
  logImportEvent('Shopify Refunds', msg, newRefunds.length);

  return msg;
}

/**
 * Imports all Squarespace refunds from the last N days (default 30).
 * Only adds NEW refunds (checks for existing Order IDs to skip duplicates).
 * The sheet is permanent history (never cleared) - reports query by date range.
 * We only query recent refunds to catch new ones added since last import.
 * @param {number} days - Number of days to look back (default 30, use 180 for full historical import)
 */
function importSquarespaceRefunds(days) {
  days = days || 30; // Default to 30 days if not specified
  const apiKey = PROPS.getProperty('SQUARESPACE_API_KEY');

  if (!apiKey) {
    throw new Error('Missing SQUARESPACE_API_KEY in Script Properties.');
  }

  logProgress('Squarespace Refunds', 'Fetching refunds from API...');

  // Get or create refunds sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refundsSheet = getOrCreateSheetWithHeaders(SQUARESPACE_REFUNDS_SHEET_NAME, SQUARESPACE_REFUNDS_HEADERS);

  // Load existing refund order IDs to avoid duplicates
  const existingRefundOrders = new Set();
  const lastRow = refundsSheet.getLastRow();
  if (lastRow > 1) {
    const existingData = refundsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(row => {
      if (row[0]) existingRefundOrders.add(String(row[0]));
    });
  }

  // Query orders from last N days
  // (We only need recent refunds since we run this regularly and the sheet keeps full history)
  const modAfter = new Date();
  modAfter.setDate(modAfter.getDate() - days);
  const modBefore = new Date(); // Now

  const modAfterStr = modAfter.toISOString();
  const modBeforeStr = modBefore.toISOString();

  const newRefunds = [];
  let cursor = null;

  do {
    let url;
    if (cursor) {
      // When using cursor, don't include date parameters
      url = `https://api.squarespace.com/1.0/commerce/orders?cursor=${encodeURIComponent(cursor)}`;
    } else {
      // Only use date parameters on first request
      url = `https://api.squarespace.com/1.0/commerce/orders?modifiedAfter=${encodeURIComponent(modAfterStr)}&modifiedBefore=${encodeURIComponent(modBeforeStr)}`;
    }

    const resp = fetchWithRetry_(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error(`Squarespace API error (${resp.getResponseCode()}): ${resp.getContentText()}`);
    }

    const json = JSON.parse(resp.getContentText());
    const orders = json.result || [];

    orders.forEach(order => {
      const orderId = order.id || '';
      const orderNumber = order.orderNumber || '';
      const orderDate = asDate_(order.createdOn);
      const refundedTotal = parseFloat(order.refundedTotal?.value) || 0;

      if (refundedTotal <= 0) return;
      if (existingRefundOrders.has(orderId)) return;

      const customerEmail = order.customerEmail || '';
      const billingAddress = order.billingAddress || {};
      const customerName = `${billingAddress.firstName || ''} ${billingAddress.lastName || ''}`.trim();

      // Squarespace doesn't provide refund issue date, use order date
      newRefunds.push([
        orderId,
        orderNumber,
        orderDate,
        orderDate, // Refund date = order date (best we can do)
        refundedTotal,
        customerEmail,
        customerName,
        'Squarespace refund (date approximated)',
        new Date()
      ]);
    });

    cursor = json.pagination?.nextPageCursor || null;
    if (cursor) Utilities.sleep(500);
  } while (cursor);

  // Append new refunds to sheet
  if (newRefunds.length > 0) {
    refundsSheet.getRange(lastRow + 1, 1, newRefunds.length, SQUARESPACE_REFUNDS_HEADERS.length).setValues(newRefunds);

    // Format columns
    const newStartRow = lastRow + 1;
    refundsSheet.getRange(newStartRow, 3, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Order Date
    refundsSheet.getRange(newStartRow, 4, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Refund Date
    refundsSheet.getRange(newStartRow, 5, newRefunds.length, 1).setNumberFormat('"$"#,##0.00'); // Refund Amount
    refundsSheet.getRange(newStartRow, 9, newRefunds.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Created At

    refundsSheet.setFrozenRows(1);
    refundsSheet.getRange(1, 1, 1, SQUARESPACE_REFUNDS_HEADERS.length).setFontWeight('bold').setBackground('#0072ce').setFontColor('#ffffff');
  }

  const msg = `✅ Imported ${newRefunds.length} new Squarespace refunds (${existingRefundOrders.size} existing)`;
  logProgress('Squarespace Refunds', msg);
  logImportEvent('Squarespace Refunds', msg, newRefunds.length);

  return msg;
}

/**
 * Wrapper function to import refunds from both platforms.
 * @param {number} days - Number of days to look back (default 30, use 180 for full historical import)
 */
function importAllRefunds(days) {
  const shopifyMsg = importShopifyRefunds(days);
  const squarespaceMsg = importSquarespaceRefunds(days);

  const msg = `${shopifyMsg}\n${squarespaceMsg}`;
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Refunds Import', 8);

  return msg;
}

/**
 * Returns a Map of Shopify refunds for a given date range.
 * Filters by REFUND ISSUE DATE (not order date).
 * Returns Map: orderId -> refundAmount
 */
function getShopifyRefundsForPeriod_(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refundsSheet = ss.getSheetByName(SHOPIFY_REFUNDS_SHEET_NAME);

  const refundsMap = new Map(); // orderId -> refundAmount

  if (!refundsSheet || refundsSheet.getLastRow() < 2) {
    return refundsMap;
  }

  const data = refundsSheet.getDataRange().getValues();
  const headers = data[0];

  const colOrderId = headers.indexOf('Order ID');
  const colRefundDate = headers.indexOf('Refund Date');
  const colRefundAmount = headers.indexOf('Refund Amount');

  if (colOrderId === -1 || colRefundDate === -1 || colRefundAmount === -1) {
    throw new Error('Missing required columns in Shopify_Refunds sheet');
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const orderId = String(row[colOrderId] || '');
    const refundDate = asDate_(row[colRefundDate]);
    const refundAmount = parseFloat(row[colRefundAmount]) || 0;

    if (!orderId || !refundDate || refundAmount <= 0) continue;

    // Filter by refund issue date
    if (refundDate >= startDate && refundDate <= endDate) {
      const existing = refundsMap.get(orderId) || 0;
      refundsMap.set(orderId, existing + refundAmount);
    }
  }

  return refundsMap;
}

/**
 * Returns a Map of Squarespace refunds for a given date range.
 * Filters by REFUND ISSUE DATE (approximated to order date since Squarespace doesn't provide it).
 * Returns Map: orderId -> refundAmount
 */
function getSquarespaceRefundsForPeriod_(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refundsSheet = ss.getSheetByName(SQUARESPACE_REFUNDS_SHEET_NAME);

  const refundsMap = new Map(); // orderId -> refundAmount

  if (!refundsSheet || refundsSheet.getLastRow() < 2) {
    return refundsMap;
  }

  const data = refundsSheet.getDataRange().getValues();
  const headers = data[0];

  const colOrderId = headers.indexOf('Order ID');
  const colRefundDate = headers.indexOf('Refund Date');
  const colRefundAmount = headers.indexOf('Refund Amount');

  if (colOrderId === -1 || colRefundDate === -1 || colRefundAmount === -1) {
    throw new Error('Missing required columns in Squarespace_Refunds sheet');
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const orderId = String(row[colOrderId] || '');
    const refundDate = asDate_(row[colRefundDate]);
    const refundAmount = parseFloat(row[colRefundAmount]) || 0;

    if (!orderId || !refundDate || refundAmount <= 0) continue;

    // Filter by refund issue date
    if (refundDate >= startDate && refundDate <= endDate) {
      const existing = refundsMap.get(orderId) || 0;
      refundsMap.set(orderId, existing + refundAmount);
    }
  }

  return refundsMap;
}
