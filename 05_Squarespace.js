// =====================================================
// 05_Squarespace.gs — Squarespace imports
// Now: each import prunes last 120 days (by Modified On) and reimports them.
// =====    ================================================

/**     
 * Deletes ALL rows whose "Modified On" date is within the last N days.
 * This is the clean "overwrite last 120 days" approach.
 */
function pruneSquarespaceLastNDays_(daysBack) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Squarespace Orders');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h || '').trim());

  // Find Modified On column robustly
  let colModified = headers.indexOf("Modified On");
  if (colModified === -1) colModified = headers.indexOf("modifiedOn");
  if (colModified === -1) {
    colModified = headers.findIndex(h => String(h).toLowerCase().includes("modified"));
  }
  if (colModified === -1) {
    throw new Error('Squarespace Orders sheet missing "Modified On" column (needed to prune last 120 days).');
  }

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysBack);

  const kept = [data[0]];
  let removed = 0;

  for (let r = 1; r < data.length; r++) {
    const d = asDate_(data[r][colModified]);
    if (d && d >= cutoff) {
      removed++;
      continue;
    }
    kept.push(data[r]);
  }

  if (removed > 0) {
    sheet.clearContents();
    const width = headers.length;
    const normalized = kept.map(row => {
      const out = row.slice(0, width);
      while (out.length < width) out.push("");
      return out;
    });
    sheet.getRange(1, 1, normalized.length, width).setValues(normalized);
  }

  return removed;
}

function importSquarespaceOrders() {
  const DAYS_BACK = 14; // Only fetch recent orders for new imports (triage handles updates)

  logImportEvent('Squarespace', `Import started (new orders only, last ${DAYS_BACK} days)`);

  const sheet = getOrCreateSheetWithHeaders('Squarespace Orders', SQUARESPACE_ORDER_HEADERS);

  // NO PRUNING - triage system handles all updates to existing orders
  // Build dedupe set from ALL existing rows
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const idCol = headers.indexOf('Order ID');
  const lineCol = headers.indexOf('LineItem ID');
  if (idCol === -1 || lineCol === -1) {
    throw new Error('Squarespace Orders sheet missing required columns: "Order ID" and/or "LineItem ID"');
  }

  // Dedupe against ALL existing rows (no prune, so we check everything)
  const existing = new Set();
  for (let i = 1; i < data.length; i++) {
    const oid = data[i][idCol];
    const lid = data[i][lineCol];
    if (oid && lid) existing.add(String(oid) + '_' + String(lid));
  }

  const apiKey = PROPS.getProperty('SQUARESPACE_API_KEY');
  if (!apiKey) throw new Error("Missing SQUARESPACE_API_KEY in Script Properties.");

  const endpoint = "https://api.squarespace.com/1.0/commerce/orders";

  // Pull all orders MODIFIED in the last 14 days (only new orders - triage handles updates)
  const d = new Date();
  d.setDate(d.getDate() - DAYS_BACK);
  const modifiedAfter = d.toISOString();
  const modifiedBefore = new Date().toISOString();

  let cursor = null;
  let rowsImported = 0;
  let page = 0;

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

    const rows = [];

    orders.forEach(order => {
      (order.lineItems || []).forEach(lineItem => {
        const uniqueKey = String(order.id || "") + '_' + String(lineItem.id || "");
        if (!uniqueKey || uniqueKey === "_") return;

        // Should not exist because we pruned the last 120 days,
        // but keep this guard for older rows.
        if (existing.has(uniqueKey)) return;

        rows.push([
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

        existing.add(uniqueKey);
        rowsImported++;
      });
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SQUARESPACE_ORDER_HEADERS.length).setValues(rows);
    }

    page++;
  } while (cursor);

  const msg = `Imported ${rowsImported} NEW Squarespace line items (last ${DAYS_BACK} days)`;
  logImportEvent('Squarespace', 'Import success (new orders only)', rowsImported);
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${msg}`, "Squarespace", 8);
  return msg;
}

/**
 * Refresh refunds for orders MODIFIED in the last daysBack days.
 * - daysBack: integer days to look back (default 120)
 * - appendMissing: boolean: if true, append line items not present in sheet (default false)
 *
 * OPTIMIZED: Batches all updates together to avoid timeout
 */
function refreshSquarespaceRefundsLastNDays_(daysBack, appendMissing) {
  daysBack = Math.max(1, parseInt(daysBack || 120, 10));
  appendMissing = !!appendMissing;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Squarespace Orders');
  if (!sheet) throw new Error('Squarespace Orders sheet not found');

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) throw new Error('Squarespace Orders sheet appears empty');

  const headers = data[0].map(h => String(h || '').trim());

  // Find key columns
  const idCol = findHeaderIndex_(headers, ['Order ID', 'order id', 'order_id']);
  const lineIdCol = findHeaderIndex_(headers, ['LineItem ID', 'LineitemID', 'lineitem id', 'line_item_id']);
  const refundCurrencyCol = findHeaderIndex_(headers, ['Refunded Total Currency', 'refunded currency']);
  const refundValueCol = findHeaderIndex_(headers, ['Refunded Total Value', 'refunded value', 'refunds']);
  const modifiedOnCol = findHeaderIndex_(headers, ['Modified On', 'modified_on', 'modifiedOn']);

  if (idCol === -1 || lineIdCol === -1) {
    throw new Error('Squarespace Orders sheet missing required columns: "Order ID" and/or "LineItem ID"');
  }

  // Build existing map: key => {row, currentRefund}
  const existing = {};
  for (let r = 1; r < data.length; r++) {
    const oid = data[r][idCol];
    const lid = data[r][lineIdCol];
    if (oid != null && lid != null && oid !== '' && lid !== '') {
      const key = String(oid) + '_' + String(lid);
      const currentRefund = refundValueCol !== -1 ? data[r][refundValueCol] : 0;
      existing[key] = { row: r + 1, currentRefund };
    }
  }

  const apiKey = PROPS.getProperty('SQUARESPACE_API_KEY');
  if (!apiKey) throw new Error("Missing SQUARESPACE_API_KEY in Script Properties.");

  const endpoint = "https://api.squarespace.com/1.0/commerce/orders";

  const d = new Date();
  d.setDate(d.getDate() - daysBack);
  const modifiedAfter = d.toISOString();
  const modifiedBefore = new Date().toISOString();

  let cursor = null;
  let page = 0;

  // Collect all updates in memory first
  const updates = [];
  const appendRows = [];

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
      const refundCurrency = order.refundedTotal?.currency || "";
      const refundValue = order.refundedTotal?.value || "";
      const modifiedOn = order.modifiedOn || "";

      (order.lineItems || []).forEach(lineItem => {
        const uniqueKey = String(order.id || "") + '_' + String(lineItem.id || "");
        if (!uniqueKey || uniqueKey === "_") return;

        const existingRow = existing[uniqueKey];
        if (existingRow) {
          // Check if refund changed
          const currentNum = typeof existingRow.currentRefund === 'number'
            ? existingRow.currentRefund
            : (parseFloat(String(existingRow.currentRefund).replace(/[^0-9.-]+/g, '')) || 0);
          const newNum = Number(refundValue) || 0;

          if (Math.abs(currentNum - newNum) > 0.0001) {
            updates.push({
              row: existingRow.row,
              refund: newNum,
              currency: refundCurrency,
              modifiedOn: modifiedOn
            });
          }
        } else if (appendMissing) {
          const width = headers.length;
          const out = Array(width).fill("");

          if (idCol !== -1) out[idCol] = order.id || "";
          if (lineIdCol !== -1) out[lineIdCol] = lineItem.id || "";
          if (refundValueCol !== -1) out[refundValueCol] = refundValue;
          if (refundCurrencyCol !== -1) out[refundCurrencyCol] = refundCurrency;
          if (modifiedOnCol !== -1) out[modifiedOnCol] = modifiedOn;

          // Populate other key fields
          const emailIdx = findHeaderIndex_(headers, ['Customer Email', 'email']);
          const liNameIdx = findHeaderIndex_(headers, ['LineItem Product Name', 'product name']);
          const liQtyIdx = findHeaderIndex_(headers, ['LineItem Quantity', 'quantity']);
          if (emailIdx !== -1) out[emailIdx] = order.customerEmail || "";
          if (liNameIdx !== -1) out[liNameIdx] = lineItem.productName || "";
          if (liQtyIdx !== -1) out[liQtyIdx] = lineItem.quantity || "";

          appendRows.push(out);
          existing[uniqueKey] = { row: sheet.getLastRow() + appendRows.length, currentRefund: refundValue };
        }
      });
    });

    page++;
  } while (cursor);

  // Batch write all updates
  let rowsUpdated = 0;
  if (updates.length > 0 && refundValueCol !== -1) {
    updates.forEach(upd => {
      sheet.getRange(upd.row, refundValueCol + 1).setValue(upd.refund);
      if (refundCurrencyCol !== -1 && upd.currency) {
        sheet.getRange(upd.row, refundCurrencyCol + 1).setValue(upd.currency);
      }
      if (modifiedOnCol !== -1 && upd.modifiedOn) {
        sheet.getRange(upd.row, modifiedOnCol + 1).setValue(upd.modifiedOn);
      }
    });
    rowsUpdated = updates.length;
  }

  // Batch write all appends
  let rowsAppended = 0;
  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
    rowsAppended = appendRows.length;
  }

  const msg = `Refreshed refunds for last ${daysBack} days: updated ${rowsUpdated} rows${appendMissing ? ', appended ' + rowsAppended + ' missing rows' : ''}.`;
  logImportEvent('Squarespace', 'Refund refresh success', rowsUpdated + rowsAppended);
  SpreadsheetApp.getActiveSpreadsheet().toast(`🔁 ${msg}`, "Squarespace Refund Refresh", 8);
  return msg;
}

/**
 * Convenience wrapper (run from Script Editor).
 */
function refreshSquarespaceRefunds() {
  return refreshSquarespaceRefundsLastNDays_(30, false);
}

/**
 * Compatibility wrappers expected by the UI/menu.
 */
function refreshSquarespaceAdjustments() {
  return refreshSquarespaceRefundsLastNDays_(30, false);
}

function refreshSquarespaceAdjustmentsLast60Days() {
  return refreshSquarespaceRefundsLastNDays_(60, true);
}
