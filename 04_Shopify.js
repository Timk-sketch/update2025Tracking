// =====================================================
// 04_Shopify.gs — Shopify imports + refunds/discounts refresh
// Reworked to include:
// - fetchWithRetry_ (retries on 429/5xx with backoff)
// - robust Link header parsing
// - refreshShopifyRefundsLastNDays_ to update refund totals in-place
// - defensive row width handling when writing rows
// =====================================================

function computeShopifyRefundTotal_(order) {
  let total = 0;

  if (order && order.refunds && Array.isArray(order.refunds) && order.refunds.length) {
    order.refunds.forEach(ref => {
      if (ref && ref.transactions && Array.isArray(ref.transactions)) {
        ref.transactions.forEach(t => {
          if (t && (t.kind === undefined || t.kind === null || t.kind === "refund")) {
            total += Math.abs(parseMoney_(t.amount));
          }
        });
      }
    });
  }

  if ((!total || total === 0) && order && order.total_refunds != null) {
    total = Math.abs(parseMoney_(order.total_refunds));
  }

  return total || 0;
}

function toLocalString_(isoOrDate) {
  if (!isoOrDate) return "";
  let d = isoOrDate;
  if (typeof isoOrDate === 'string') {
    d = new Date(isoOrDate);
  }
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  try {
    return d.toLocaleString();
  } catch (e) {
    return d.toString();
  }
}

/**
 * Fetch wrapper with retries for transient errors (429 and 5xx).
 */
function fetchWithRetry_(url, options) {
  const MAX_RETRIES = 4;
  const RETRY_BASE_MS = 1000;
  let attempt = 0;
  let lastErr = null;

  options = options || {};

  while (attempt <= MAX_RETRIES) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      if (code === 200) {
        return resp;
      }

      // Retry on rate limit or server errors
      if (code === 429 || (code >= 500 && code < 600)) {
        lastErr = new Error('HTTP ' + code + ': ' + resp.getContentText());
        // fall through to retry
      } else {
        // Non-retryable - throw immediately with response content
        throw new Error(`Shopify API error (${code}): ${resp.getContentText()}`);
      }
    } catch (e) {
      lastErr = e;
    }

    attempt++;
    if (attempt > MAX_RETRIES) break;

    const backoff = RETRY_BASE_MS * Math.pow(2, attempt - 1);
    const jitter = Math.floor(Math.random() * 300);
    Utilities.sleep(backoff + jitter);
  }

  throw lastErr || new Error('Unknown fetch error for ' + url);
}

/**
 * Parses Link header (RFC5988 style) and returns the URL for rel="next" if present.
 */
function parseLinkHeader_(linkHeader) {
  if (!linkHeader) return null;
  let link = linkHeader;
  if (Array.isArray(linkHeader)) {
    link = linkHeader.join(',');
  }
  const parts = link.split(',');
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i].trim();
    const match = p.match(/<([^>]+)>;\s*rel\s*=\s*"?([^"]+)"?/);
    if (match) {
      const url = match[1];
      const rel = match[2];
      if (rel === 'next') return url;
    }
  }
  return null;
}

/**
 * Flexible header index finder.
 * headers: array of header strings
 * candidates: array of candidate header names (case-insensitive)
 */
function findHeaderIndex_(headers, candidates) {
  if (!headers || !headers.length) return -1;
  const normalized = headers.map(h => String(h || '').trim().toLowerCase());
  for (let j = 0; j < candidates.length; j++) {
    const cand = String(candidates[j] || '').trim().toLowerCase();
    for (let i = 0; i < normalized.length; i++) {
      if (normalized[i] === cand) return i;
    }
  }
  // fuzzy contains
  for (let j = 0; j < candidates.length; j++) {
    const cand = String(candidates[j] || '').trim().toLowerCase();
    for (let i = 0; i < normalized.length; i++) {
      if (normalized[i].includes(cand)) return i;
    }
  }
  return -1;
}

/**
 * Deletes ALL rows whose "Updated At" date is within the last N days.
 */
function pruneShopifyLastNDays_(daysBack) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Shopify Orders');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h || '').trim());

  // Find Updated At column robustly
  let colUpdated = headers.indexOf("Updated At");
  if (colUpdated === -1) colUpdated = headers.indexOf("updated_at");
  if (colUpdated === -1) {
    colUpdated = headers.findIndex(h => String(h || '').toLowerCase().includes("updated"));
  }
  if (colUpdated === -1) {
    throw new Error('Shopify Orders sheet missing "Updated At" column (needed to prune last ' + daysBack + ' days).');
  }

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysBack);

  const kept = [data[0]];
  let removed = 0;

  for (let r = 1; r < data.length; r++) {
    const d = asDate_(data[r][colUpdated]);
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

/**
 * Main import function (keeps overwrite-last-N-days behavior).
 * Remains compatible with SHOPIFY_ORDER_HEADERS defined in 00_Config.gs.
 */
function importShopifyOrders() {
  const DAYS_BACK = 14; // Only fetch recent orders for new imports (triage handles updates)

  logImportEvent('Shopify', `Import started (new orders only, last ${DAYS_BACK} days)`);

  const sheet = getOrCreateSheetWithHeaders('Shopify Orders', SHOPIFY_ORDER_HEADERS);

  // NO PRUNING - triage system handles all updates to existing orders
  // Build dedupe set from ALL existing rows
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());

  const idCol = headers.indexOf('Order ID');
  const lineIdCol = headers.indexOf('Lineitem ID');

  if (idCol === -1 || lineIdCol === -1) {
    throw new Error('Shopify Orders sheet missing required columns: "Order ID" and/or "Lineitem ID"');
  }

  // Dedupe against ALL existing rows (no prune, so we check everything)
  const existing = new Set();
  for (let i = 1; i < data.length; i++) {
    const oid = data[i][idCol];
    const lid = data[i][lineIdCol];
    if (oid && lid) existing.add(String(oid) + '_' + String(lid));
  }

  const apiKey = PROPS.getProperty('SHOPIFY_API_KEY');
  const shopDomain = PROPS.getProperty('SHOPIFY_SHOP_DOMAIN');
  const apiVersion = '2023-10';

  if (!apiKey || !shopDomain) throw new Error("Missing SHOPIFY_API_KEY or SHOPIFY_SHOP_DOMAIN in Script Properties.");

  const d = new Date();
  d.setDate(d.getDate() - DAYS_BACK);
  const updatedAtMin = d.toISOString();

  let url = `https://${shopDomain}/admin/api/${apiVersion}/orders.json?status=any&limit=250&updated_at_min=${encodeURIComponent(updatedAtMin)}`;

  let rowsImported = 0;

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

    const rows = [];

    orders.forEach(order => {
      const refundsTotal = computeShopifyRefundTotal_(order);
      const isTest = order && order.test === true;

      const discountCodes = (order.discount_codes || []).map(d => d.code).join(", ");
      const shippingMethod = (order.shipping_lines || []).map(s => s.title || s.code).join(", ");

      (order.line_items || []).forEach(lineItem => {
        const uniqueKey = String(order.id || "") + '_' + String(lineItem.id || "");
        if (!uniqueKey || uniqueKey === "_") return;

        // Should not exist because we pruned the last DAYS_BACK days,
        // but keep this guard for older rows.
        if (existing.has(uniqueKey)) return;

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

        // Ensure row length matches header length (pad/truncate)
        if (outRow.length < SHOPIFY_ORDER_HEADERS.length) {
          while (outRow.length < SHOPIFY_ORDER_HEADERS.length) outRow.push("");
        } else if (outRow.length > SHOPIFY_ORDER_HEADERS.length) {
          outRow.length = SHOPIFY_ORDER_HEADERS.length;
        }

        rows.push(outRow);
        existing.add(uniqueKey);
        rowsImported++;
      });
    });

    if (rows.length) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, SHOPIFY_ORDER_HEADERS.length).setValues(rows);
    }

    const headersObj = resp.getHeaders ? resp.getHeaders() : {};
    const linkHeader = headersObj['Link'] || headersObj['link'] || headersObj['LINK'];
    const nextUrl = parseLinkHeader_(linkHeader);
    url = nextUrl || null;
  }

  const msg = `Imported ${rowsImported} NEW Shopify line items (last ${DAYS_BACK} days)`;
  logImportEvent('Shopify', 'Import success (new orders only)', rowsImported);
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${msg}`, "Shopify", 8);
  return msg;
}

/**
 * Refresh refunds (and optional small set of columns) for orders UPDATED in the last daysBack days.
 * - daysBack: integer days to look back (default 120)
 * - appendMissing: boolean: if true, append line items not present in sheet (default false)
 *
 * OPTIMIZED: Batches all updates together to avoid timeout
 */
function refreshShopifyRefundsLastNDays_(daysBack, appendMissing) {
  daysBack = Math.max(1, parseInt(daysBack || 120, 10));
  appendMissing = !!appendMissing;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Shopify Orders');
  if (!sheet) throw new Error('Shopify Orders sheet not found');

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) throw new Error('Shopify Orders sheet appears empty');

  const headers = data[0].map(h => String(h || '').trim());

  // Find key columns
  const idCol = findHeaderIndex_(headers, ['Order ID', 'order id', 'order_id']);
  const lineIdCol = findHeaderIndex_(headers, ['Lineitem ID', 'LineitemID', 'lineitem id', 'line_item_id']);
  const refundsCol = findHeaderIndex_(headers, ['Total Refunds', 'Refunds', 'refunds', 'total_refunds', 'refund_total']);
  const updatedAtCol = findHeaderIndex_(headers, ['Updated At', 'updated_at', 'updated at']);

  if (idCol === -1 || lineIdCol === -1) {
    throw new Error('Shopify Orders sheet missing required columns: "Order ID" and/or "Lineitem ID"');
  }

  // Build existing map: key => {row, currentRefund}
  const existing = {};
  for (let r = 1; r < data.length; r++) {
    const oid = data[r][idCol];
    const lid = data[r][lineIdCol];
    if (oid != null && lid != null && oid !== '' && lid !== '') {
      const key = String(oid) + '_' + String(lid);
      const currentRefund = refundsCol !== -1 ? data[r][refundsCol] : 0;
      existing[key] = { row: r + 1, currentRefund };
    }
  }

  const apiKey = PROPS.getProperty('SHOPIFY_API_KEY');
  const shopDomain = PROPS.getProperty('SHOPIFY_SHOP_DOMAIN');
  const apiVersion = '2023-10';

  if (!apiKey || !shopDomain) throw new Error("Missing SHOPIFY_API_KEY or SHOPIFY_SHOP_DOMAIN in Script Properties.");

  const d = new Date();
  d.setDate(d.getDate() - daysBack);
  const updatedAtMin = d.toISOString();

  let url = `https://${shopDomain}/admin/api/${apiVersion}/orders.json?status=any&limit=250&updated_at_min=${encodeURIComponent(updatedAtMin)}`;

  // Collect all updates in memory first
  const updates = [];
  const appendRows = [];

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
      const updatedAt = order.updated_at || "";

      (order.line_items || []).forEach(lineItem => {
        const uniqueKey = String(order.id || "") + '_' + String(lineItem.id || "");
        if (!uniqueKey || uniqueKey === "_") return;

        const existingRow = existing[uniqueKey];
        if (existingRow) {
          // Check if refund changed
          const currentNum = typeof existingRow.currentRefund === 'number'
            ? existingRow.currentRefund
            : (parseFloat(String(existingRow.currentRefund).replace(/[^0-9.-]+/g, '')) || 0);
          const newNum = Number(refundsTotal) || 0;

          if (Math.abs(currentNum - newNum) > 0.0001) {
            updates.push({
              row: existingRow.row,
              refund: newNum,
              updatedAt: updatedAt
            });
          }
        } else if (appendMissing) {
          const width = headers.length;
          const out = Array(width).fill("");

          if (idCol !== -1) out[idCol] = order.id || "";
          if (lineIdCol !== -1) out[lineIdCol] = lineItem.id || "";
          if (refundsCol !== -1) out[refundsCol] = refundsTotal;
          if (updatedAtCol !== -1) out[updatedAtCol] = updatedAt;

          // Populate other key fields
          const liNameIdx = findHeaderIndex_(headers, ['Lineitem Name', 'lineitem name']);
          const liQtyIdx = findHeaderIndex_(headers, ['Lineitem Quantity', 'quantity']);
          const emailIdx = findHeaderIndex_(headers, ['Customer Email', 'email']);
          if (liNameIdx !== -1) out[liNameIdx] = lineItem.name || "";
          if (liQtyIdx !== -1) out[liQtyIdx] = lineItem.quantity || "";
          if (emailIdx !== -1) out[emailIdx] = order.email || "";

          appendRows.push(out);
          existing[uniqueKey] = { row: sheet.getLastRow() + appendRows.length, currentRefund: refundsTotal };
        }
      });
    });

    const headersObj = resp.getHeaders ? resp.getHeaders() : {};
    const linkHeader = headersObj['Link'] || headersObj['link'] || headersObj['LINK'];
    const nextUrl = parseLinkHeader_(linkHeader);
    url = nextUrl || null;
  }

  // Batch write all updates using setValues for maximum speed
  let rowsUpdated = 0;
  if (updates.length > 0 && refundsCol !== -1) {
    // Group consecutive rows for batch writes
    updates.sort((a, b) => a.row - b.row);

    for (let i = 0; i < updates.length; i++) {
      sheet.getRange(updates[i].row, refundsCol + 1).setValue(updates[i].refund);
      if (updatedAtCol !== -1 && updates[i].updatedAt) {
        sheet.getRange(updates[i].row, updatedAtCol + 1).setValue(updates[i].updatedAt);
      }
    }
    rowsUpdated = updates.length;
  }

  // Batch write all appends
  let rowsAppended = 0;
  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
    rowsAppended = appendRows.length;
  }

  const msg = `Refreshed refunds for last ${daysBack} days: updated ${rowsUpdated} rows${appendMissing ? ', appended ' + rowsAppended + ' missing rows' : ''}.`;
  logImportEvent('Shopify', 'Refund refresh success', rowsUpdated + rowsAppended);
  SpreadsheetApp.getActiveSpreadsheet().toast(`🔁 ${msg}`, "Shopify Refund Refresh", 8);
  return msg;
}

/**
 * Convenience wrapper (run from Script Editor).
 */
function refreshShopifyRefunds() {
  return refreshShopifyRefundsLastNDays_(30, false);
}

/**
 * Compatibility wrappers expected by the UI/menu.
 */
function refreshShopifyAdjustments() {
  return refreshShopifyRefundsLastNDays_(30, false);
}

function refreshShopifyAdjustmentsLast60Days() {
  return refreshShopifyRefundsLastNDays_(60, true);
}