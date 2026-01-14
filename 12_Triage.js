// =====================================================
// 12_Triage.js — Triage system for incremental order updates
// Handles refunds/discounts that happen after initial import
// =====================================================
//
// DEPRECATED: This triage system has been replaced with a simpler
// "Update Orders (Check Refunds)" function that uses:
// - Shopify: financial_status filters (refunded/partially_refunded)
// - Squarespace: modifiedAfter filter
//
// These functions are kept for backward compatibility but are no longer
// used in the main menu or recommended workflow.
// =====================================================

/**
 * Import Shopify orders to triage based on date range
 * @param {string} dateRange - "0-30", "31-60", "61-90", or "91-120"
 */
function importShopifyToTriage(dateRange) {
  const ranges = {
    "0-30": { start: 0, end: 30 },
    "31-60": { start: 31, end: 60 },
    "61-90": { start: 61, end: 90 },
    "91-120": { start: 91, end: 120 }
  };

  const range = ranges[dateRange];
  if (!range) throw new Error(`Invalid date range: ${dateRange}. Use "0-30", "31-60", "61-90", or "91-120"`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triageSheet = getOrCreateSheetWithHeaders('Shopify Triage', SHOPIFY_ORDER_HEADERS);

  const apiKey = PROPS.getProperty('SHOPIFY_API_KEY');
  const shopDomain = PROPS.getProperty('SHOPIFY_SHOP_DOMAIN');
  const apiVersion = '2023-10';

  if (!apiKey || !shopDomain) throw new Error("Missing SHOPIFY_API_KEY or SHOPIFY_SHOP_DOMAIN in Script Properties.");

  // Calculate date range
  const endDate = new Date();
  endDate.setDate(endDate.getDate() - range.start);
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - range.end);

  const updatedAtMin = startDate.toISOString();
  const updatedAtMax = endDate.toISOString();

  let url = `https://${shopDomain}/admin/api/${apiVersion}/orders.json?status=any&limit=250&updated_at_min=${encodeURIComponent(updatedAtMin)}&updated_at_max=${encodeURIComponent(updatedAtMax)}`;

  let rowsImported = 0;
  const triageRows = [];

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
      const refundsTotal = computeShopifyRefundTotal_(order);
      const isTest = order && order.test === true;
      const discountCodes = (order.discount_codes || []).map(d => d.code).join(", ");
      const shippingMethod = (order.shipping_lines || []).map(s => s.title || s.code).join(", ");

      (order.line_items || []).forEach(lineItem => {
        const outRow = [
          order.id || "",
          order.order_number || "",
          order.created_at || "",
          order.processed_at || "",
          order.updated_at || "",
          order.financial_status || "",
          order.fulfillment_status || "",
          order.currency || "",
          order.total_price || "",
          order.subtotal_price || "",
          order.total_tax || "",
          order.total_discounts || "",
          order.current_total_price || "",
          order.current_total_discounts || "",
          refundsTotal,
          isTest ? "TRUE" : "",
          order.email || "",
          order.customer?.first_name || "",
          order.customer?.last_name || "",
          order.billing_address?.name || "",
          order.billing_address?.address1 || "",
          order.billing_address?.address2 || "",
          order.billing_address?.city || "",
          order.billing_address?.province || "",
          order.billing_address?.country || "",
          order.billing_address?.zip || "",
          order.billing_address?.phone || "",
          order.shipping_address?.name || "",
          order.shipping_address?.address1 || "",
          order.shipping_address?.address2 || "",
          order.shipping_address?.city || "",
          order.shipping_address?.province || "",
          order.shipping_address?.country || "",
          order.shipping_address?.zip || "",
          order.shipping_address?.phone || "",
          lineItem.id || "",
          lineItem.name || "",
          lineItem.quantity || "",
          lineItem.price || "",
          lineItem.sku || "",
          lineItem.requires_shipping != null ? String(lineItem.requires_shipping) : "",
          lineItem.taxable != null ? String(lineItem.taxable) : "",
          lineItem.fulfillment_status || "",
          (order.tags || ""),
          (order.note || ""),
          order.gateway || "",
          order.total_weight || "",
          discountCodes,
          shippingMethod,
          toLocalString_(order.created_at),
          toLocalString_(order.processed_at)
        ];

        // Ensure row length matches
        while (outRow.length < SHOPIFY_ORDER_HEADERS.length) outRow.push("");
        if (outRow.length > SHOPIFY_ORDER_HEADERS.length) outRow.length = SHOPIFY_ORDER_HEADERS.length;

        triageRows.push(outRow);
        rowsImported++;
      });
    });

    const headersObj = resp.getHeaders ? resp.getHeaders() : {};
    const linkHeader = headersObj['Link'] || headersObj['link'] || headersObj['LINK'];
    const nextUrl = parseLinkHeader_(linkHeader);
    url = nextUrl || null;
  }

  // Write all triage rows at once
  if (triageRows.length > 0) {
    const startRow = triageSheet.getLastRow() + 1;
    triageSheet.getRange(startRow, 1, triageRows.length, SHOPIFY_ORDER_HEADERS.length).setValues(triageRows);
  }

  const msg = `Imported ${rowsImported} Shopify orders (days ${range.start}-${range.end}) to triage`;
  logImportEvent('Shopify Triage', msg, rowsImported);
  ss.toast(`✅ ${msg}`, "Shopify Triage", 6);
  return msg;
}

/**
 * Import Squarespace orders to triage based on date range
 * @param {string} dateRange - "0-30", "31-60", "61-90", or "91-120"
 */
function importSquarespaceToTriage(dateRange) {
  const ranges = {
    "0-30": { start: 0, end: 30 },
    "31-60": { start: 31, end: 60 },
    "61-90": { start: 61, end: 90 },
    "91-120": { start: 91, end: 120 }
  };

  const range = ranges[dateRange];
  if (!range) throw new Error(`Invalid date range: ${dateRange}. Use "0-30", "31-60", "61-90", or "91-120"`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triageSheet = getOrCreateSheetWithHeaders('Squarespace Triage', SQUARESPACE_ORDER_HEADERS);

  const apiKey = PROPS.getProperty('SQUARESPACE_API_KEY');
  if (!apiKey) throw new Error("Missing SQUARESPACE_API_KEY in Script Properties.");

  const endpoint = "https://api.squarespace.com/1.0/commerce/orders";

  // Calculate date range
  const endDate = new Date();
  endDate.setDate(endDate.getDate() - range.start);
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - range.end);

  const modifiedAfter = startDate.toISOString();
  const modifiedBefore = endDate.toISOString();

  let cursor = null;
  let rowsImported = 0;
  let page = 0;
  const triageRows = [];

  do {
    let finalUrl;
    if (page === 0) {
      finalUrl = endpoint + '?' +
        `modifiedAfter=${encodeURIComponent(modifiedAfter)}` +
        `&modifiedBefore=${encodeURIComponent(modifiedBefore)}`;
    } else {
      finalUrl = endpoint + `?cursor=${encodeURIComponent(cursor)}`;
    }

    const resp = UrlFetchApp.fetch(finalUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + apiKey, "accept": "application/json" },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error(`Squarespace API error (${resp.getResponseCode()}): ${resp.getContentText()}`);
    }

    const json = JSON.parse(resp.getContentText());
    const orders = json.result || [];
    cursor = json.pagination && json.pagination.nextPageCursor ? json.pagination.nextPageCursor : null;

    orders.forEach(order => {
      (order.lineItems || []).forEach(lineItem => {
        triageRows.push([
          order.id || "",
          order.orderNumber || "",
          order.createdOn || "",
          order.modifiedOn || "",
          order.channel || "",
          order.testMode || "",
          order.customerEmail || "",
          order.billingAddress?.firstName || "",
          order.billingAddress?.lastName || "",
          order.billingAddress?.address1 || "",
          order.billingAddress?.address2 || "",
          order.billingAddress?.city || "",
          order.billingAddress?.state || "",
          order.billingAddress?.countryCode || "",
          order.billingAddress?.postalCode || "",
          order.billingAddress?.phone || "",
          order.shippingAddress?.firstName || "",
          order.shippingAddress?.lastName || "",
          order.shippingAddress?.address1 || "",
          order.shippingAddress?.address2 || "",
          order.shippingAddress?.city || "",
          order.shippingAddress?.state || "",
          order.shippingAddress?.countryCode || "",
          order.shippingAddress?.postalCode || "",
          order.shippingAddress?.phone || "",
          order.fulfillmentStatus || "",
          order.internalNotes || "",
          order.subtotal?.currency || "",
          order.subtotal?.value || "",
          order.shippingTotal?.currency || "",
          order.shippingTotal?.value || "",
          order.discountTotal?.currency || "",
          order.discountTotal?.value || "",
          order.taxTotal?.currency || "",
          order.taxTotal?.value || "",
          order.refundedTotal?.currency || "",
          order.refundedTotal?.value || "",
          order.grandTotal?.currency || "",
          order.grandTotal?.value || "",
          lineItem.id || "",
          lineItem.sku || "",
          lineItem.weight || "",
          lineItem.width || "",
          lineItem.length || "",
          lineItem.height || "",
          lineItem.productId || "",
          lineItem.productName || "",
          lineItem.quantity || "",
          lineItem.unitPrice?.currency || "",
          lineItem.unitPrice?.value || "",
          (lineItem.customizations || []).map(c => c.value || "").join(", "),
          lineItem.type || "",
          (order.shippingLines || []).map(s => s.title || s.type || "").join(", "),
          (order.discountLines || []).map(d => d.name || "").join(", ")
        ]);

        rowsImported++;
      });
    });

    page++;
  } while (cursor);

  // Write all triage rows at once
  if (triageRows.length > 0) {
    triageSheet.getRange(triageSheet.getLastRow() + 1, 1, triageRows.length, SQUARESPACE_ORDER_HEADERS.length).setValues(triageRows);
  }

  const msg = `Imported ${rowsImported} Squarespace orders (days ${range.start}-${range.end}) to triage`;
  logImportEvent('Squarespace Triage', msg, rowsImported);
  ss.toast(`✅ ${msg}`, "Squarespace Triage", 6);
  return msg;
}

/**
 * Clean triage: reconcile triage orders against main sheets and update
 * Deletes processed rows from triage
 */
function cleanTriage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let totalUpdated = 0;
  let totalProcessed = 0;

  // Process Shopify Triage
  const shopifyResult = cleanShopifyTriage_();
  totalUpdated += shopifyResult.updated;
  totalProcessed += shopifyResult.processed;

  // Process Squarespace Triage
  const squarespaceResult = cleanSquarespaceTriage_();
  totalUpdated += squarespaceResult.updated;
  totalProcessed += squarespaceResult.processed;

  // Check if there are still rows remaining
  const shopifyTriage = ss.getSheetByName('Shopify Triage');
  const squarespaceTriage = ss.getSheetByName('Squarespace Triage');

  const shopifyRemaining = shopifyTriage ? Math.max(0, shopifyTriage.getLastRow() - 1) : 0;
  const squarespaceRemaining = squarespaceTriage ? Math.max(0, squarespaceTriage.getLastRow() - 1) : 0;
  const totalRemaining = shopifyRemaining + squarespaceRemaining;

  let msg = `Triage cleaned: ${totalProcessed} orders processed, ${Math.round(totalUpdated)} main sheet rows updated`;

  if (totalRemaining > 0) {
    msg += `\n⚠️ ${totalRemaining} rows remain (${shopifyRemaining} Shopify, ${squarespaceRemaining} Squarespace). Run Clean Triage again to continue.`;
    logImportEvent('Triage', `Clean triage partial (${totalRemaining} rows remain)`, totalProcessed);
  } else {
    msg += '\n✅ All triage sheets are now empty!';
    logImportEvent('Triage', 'Clean triage complete (all empty)', totalProcessed);
  }

  ss.toast(msg, "Triage Clean", 10);
  return msg;
}

/**
 * Internal: Clean Shopify triage
 */
function cleanShopifyTriage_() {
  const MAX_ROWS_PER_RUN = 500; // Process max 500 rows per run to avoid timeout

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triageSheet = ss.getSheetByName('Shopify Triage');
  const mainSheet = ss.getSheetByName('Shopify Orders');

  if (!triageSheet || triageSheet.getLastRow() < 2) {
    return { updated: 0, processed: 0 };
  }
  if (!mainSheet) {
    throw new Error('Shopify Orders sheet not found');
  }

  const triageData = triageSheet.getDataRange().getValues();
  const mainData = mainSheet.getDataRange().getValues();

  const triageHeaders = triageData[0].map(h => String(h || '').trim());
  const mainHeaders = mainData[0].map(h => String(h || '').trim());

  // Find columns
  const tIdCol = triageHeaders.indexOf('Order ID');
  const tLineIdCol = triageHeaders.indexOf('Lineitem ID');
  const tRefundsCol = triageHeaders.indexOf('Total Refunds');
  const tDiscountsCol = triageHeaders.indexOf('Total Discounts');
  const tUpdatedCol = triageHeaders.indexOf('Updated At');

  const mIdCol = mainHeaders.indexOf('Order ID');
  const mLineIdCol = mainHeaders.indexOf('Lineitem ID');
  const mRefundsCol = mainHeaders.indexOf('Total Refunds');
  const mDiscountsCol = mainHeaders.indexOf('Total Discounts');
  const mUpdatedCol = mainHeaders.indexOf('Updated At');

  // Build main sheet index
  const mainIndex = {};
  for (let r = 1; r < mainData.length; r++) {
    const oid = mainData[r][mIdCol];
    const lid = mainData[r][mLineIdCol];
    if (oid && lid) {
      const key = String(oid) + '_' + String(lid);
      mainIndex[key] = r;
    }
  }

  // Collect updates in memory for batch write
  const updates = [];
  const rowsToDelete = [];

  // Limit rows processed per run
  const rowsToProcess = Math.min(triageData.length - 1, MAX_ROWS_PER_RUN);

  // Process triage rows
  for (let r = 1; r <= rowsToProcess; r++) {
    const oid = triageData[r][tIdCol];
    const lid = triageData[r][tLineIdCol];
    if (!oid || !lid) {
      rowsToDelete.push(r + 1); // Delete invalid rows
      continue;
    }

    const key = String(oid) + '_' + String(lid);
    const mainRow = mainIndex[key];

    if (mainRow !== undefined) {
      // Collect updates for batch write
      const mainRowNum = mainRow + 1; // 1-based

      if (mRefundsCol !== -1 && tRefundsCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mRefundsCol + 1,
          value: triageData[r][tRefundsCol]
        });
      }
      if (mDiscountsCol !== -1 && tDiscountsCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mDiscountsCol + 1,
          value: triageData[r][tDiscountsCol]
        });
      }
      if (mUpdatedCol !== -1 && tUpdatedCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mUpdatedCol + 1,
          value: triageData[r][tUpdatedCol]
        });
      }
    }

    // Mark for deletion (processed)
    rowsToDelete.push(r + 1); // 1-based
  }

  // Batch write all updates
  if (updates.length > 0) {
    updates.forEach(upd => {
      mainSheet.getRange(upd.row, upd.col).setValue(upd.value);
    });
  }

  // Delete processed rows from triage (bottom to top to preserve indices)
  if (rowsToDelete.length > 0) {
    rowsToDelete.reverse().forEach(rowNum => {
      triageSheet.deleteRow(rowNum);
    });
  }

  return { updated: updates.length / 3, processed: rowsToDelete.length }; // Divide by 3 since we update 3 cols per order
}

/**
 * Internal: Clean Squarespace triage
 */
function cleanSquarespaceTriage_() {
  const MAX_ROWS_PER_RUN = 500; // Process max 500 rows per run to avoid timeout

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triageSheet = ss.getSheetByName('Squarespace Triage');
  const mainSheet = ss.getSheetByName('Squarespace Orders');

  if (!triageSheet || triageSheet.getLastRow() < 2) {
    return { updated: 0, processed: 0 };
  }
  if (!mainSheet) {
    throw new Error('Squarespace Orders sheet not found');
  }

  const triageData = triageSheet.getDataRange().getValues();
  const mainData = mainSheet.getDataRange().getValues();

  const triageHeaders = triageData[0].map(h => String(h || '').trim());
  const mainHeaders = mainData[0].map(h => String(h || '').trim());

  // Find columns
  const tIdCol = triageHeaders.indexOf('Order ID');
  const tLineIdCol = triageHeaders.indexOf('LineItem ID');
  const tRefundValueCol = triageHeaders.indexOf('Refunded Total Value');
  const tDiscountValueCol = triageHeaders.indexOf('Discount Total Value');
  const tModifiedCol = triageHeaders.indexOf('Modified On');

  const mIdCol = mainHeaders.indexOf('Order ID');
  const mLineIdCol = mainHeaders.indexOf('LineItem ID');
  const mRefundValueCol = mainHeaders.indexOf('Refunded Total Value');
  const mDiscountValueCol = mainHeaders.indexOf('Discount Total Value');
  const mModifiedCol = mainHeaders.indexOf('Modified On');

  // Build main sheet index
  const mainIndex = {};
  for (let r = 1; r < mainData.length; r++) {
    const oid = mainData[r][mIdCol];
    const lid = mainData[r][mLineIdCol];
    if (oid && lid) {
      const key = String(oid) + '_' + String(lid);
      mainIndex[key] = r;
    }
  }

  // Collect updates in memory for batch write
  const updates = [];
  const rowsToDelete = [];

  // Limit rows processed per run
  const rowsToProcess = Math.min(triageData.length - 1, MAX_ROWS_PER_RUN);

  // Process triage rows
  for (let r = 1; r <= rowsToProcess; r++) {
    const oid = triageData[r][tIdCol];
    const lid = triageData[r][tLineIdCol];
    if (!oid || !lid) {
      rowsToDelete.push(r + 1); // Delete invalid rows
      continue;
    }

    const key = String(oid) + '_' + String(lid);
    const mainRow = mainIndex[key];

    if (mainRow !== undefined) {
      // Collect updates for batch write
      const mainRowNum = mainRow + 1; // 1-based

      if (mRefundValueCol !== -1 && tRefundValueCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mRefundValueCol + 1,
          value: triageData[r][tRefundValueCol]
        });
      }
      if (mDiscountValueCol !== -1 && tDiscountValueCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mDiscountValueCol + 1,
          value: triageData[r][tDiscountValueCol]
        });
      }
      if (mModifiedCol !== -1 && tModifiedCol !== -1) {
        updates.push({
          row: mainRowNum,
          col: mModifiedCol + 1,
          value: triageData[r][tModifiedCol]
        });
      }
    }

    // Mark for deletion (processed)
    rowsToDelete.push(r + 1); // 1-based
  }

  // Batch write all updates
  if (updates.length > 0) {
    updates.forEach(upd => {
      mainSheet.getRange(upd.row, upd.col).setValue(upd.value);
    });
  }

  // Delete processed rows from triage (bottom to top to preserve indices)
  if (rowsToDelete.length > 0) {
    rowsToDelete.reverse().forEach(rowNum => {
      triageSheet.deleteRow(rowNum);
    });
  }

  return { updated: updates.length / 3, processed: rowsToDelete.length }; // Divide by 3 since we update 3 cols per order
}

// Convenience wrappers for menu
function importShopifyTriage0to30() { return importShopifyToTriage("0-30"); }
function importShopifyTriage31to60() { return importShopifyToTriage("31-60"); }
function importShopifyTriage61to90() { return importShopifyToTriage("61-90"); }
function importShopifyTriage91to120() { return importShopifyToTriage("91-120"); }

function importSquarespaceTriage0to30() { return importSquarespaceToTriage("0-30"); }
function importSquarespaceTriage31to60() { return importSquarespaceToTriage("31-60"); }
function importSquarespaceTriage61to90() { return importSquarespaceToTriage("61-90"); }
function importSquarespaceTriage91to120() { return importSquarespaceToTriage("91-120"); }
