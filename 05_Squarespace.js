// =====================================================
// 05_Squarespace.gs — Squarespace imports
// Now: each import prunes last 120 days (by Modified On) and reimports them.
// =====================================================

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
  const DAYS_BACK = 120;

  logImportEvent('Squarespace', `Import started (overwrite last ${DAYS_BACK} days)`);

  const sheet = getOrCreateSheetWithHeaders('Squarespace Orders', SQUARESPACE_ORDER_HEADERS);

  // ✅ Clean overwrite behavior: delete anything modified in last 120 days
  const removed = pruneSquarespaceLastNDays_(DAYS_BACK);

  // Re-read after prune
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const idCol = headers.indexOf('Order ID');
  const lineCol = headers.indexOf('LineItem ID');
  if (idCol === -1 || lineCol === -1) {
    throw new Error('Squarespace Orders sheet missing required columns: "Order ID" and/or "LineItem ID"');
  }

  // Dedupe against remaining historical rows only (older than 120-day cutoff)
  const existing = new Set();
  for (let i = 1; i < data.length; i++) {
    const oid = data[i][idCol];
    const lid = data[i][lineCol];
    if (oid && lid) existing.add(String(oid) + '_' + String(lid));
  }

  const apiKey = PROPS.getProperty('SQUARESPACE_API_KEY');
  if (!apiKey) throw new Error("Missing SQUARESPACE_API_KEY in Script Properties.");

  const endpoint = "https://api.squarespace.com/1.0/commerce/orders";

  // Pull all orders MODIFIED in the last 120 days (captures refunds that happen later)
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

  const msg = `Overwrite last ${DAYS_BACK} days: removed ${removed} rows, imported ${rowsImported} Squarespace line items`;
  logImportEvent('Squarespace', 'Import success', rowsImported);
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${msg}`, "Squarespace", 8);
  return msg;
}
