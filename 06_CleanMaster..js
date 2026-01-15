/**
 * 06_CleanMaster.gs
 * Fixes:
 * - Robust date parsing so All_Orders_Clean.order_date is populated
 * - Adds reset + backfill tools so you can repair existing clean data
 * - FIX Squarespace: if unit_price/line_revenue missing, allocate order net across lines
 *
 * ADD (2026-01): Squarespace duplicate prevention:
 * - Exclude Squarespace orders in year 2026 whose product name indicates a Montana LLC renewal
 *   (contains: (montana OR mt) AND llc AND renew*)
 *
 * NOTE:
 * This file assumes these already exist in your other files:
 * - PROPS (ScriptProperties), CLEAN_OUTPUT_SHEET, CLEAN_HEADERS
 * - loadBannedList_(), isBannedEmail_(), getOrCreateSheetWithHeaders()
 * - truthy_(), s_(), n_(), parseMoney_(), parseQty_(), normEmail_()
 * - logImportEvent()
 */

// ---------------------------
// BANNED PRODUCT KEYWORDS
// Orders containing these product names will be excluded
// ---------------------------
const BANNED_PRODUCT_KEYWORDS = [
  "Roxo",
  "Rough Country",
  "Rigid",
  "Rugged Ridge",
  "ScanGauge",
  "Shorty Stunt",
  "Smittybilt",
  "Spoke",
  "Squadron",
  "Honda Talon",
  "Stainless Steel",
  "Standard Side",
  "Stealth",
  "Sticker Bomb",
  "SUZUKI DRZ400SM",
  "Subaru Crosstrek",
  "Tactical",
  "Speedometer",
  "Trail Tech",
  "Trailmax",
  "Skid Plate",
  "Tusk",
  "Universal",
  "Signal",
  "UTV Conversion",
  "UTV Plug",
  "Legal Conversion",
  "Vehicle Sales Tax",
  "Vintage Air",
  "Winch",
  "Wheelie",
  "Windshield",
  "WR250 R/X",
  "OEM",
  "ZETA"
];

/**
 * Checks if a product name contains any banned keywords.
 * Case-insensitive partial matching.
 */
function isBannedProduct_(productName) {
  if (!productName) return false;
  const productLower = String(productName).toLowerCase();

  for (const keyword of BANNED_PRODUCT_KEYWORDS) {
    if (productLower.includes(keyword.toLowerCase())) {
      return true;
    }
  }
  return false;
}

// ---------------------------
// PUBLIC: reset clean build state + clear output (optional)
// ---------------------------
function resetCleanMasterBuildState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PROPS.deleteProperty("CLEAN_MASTER_BUILD_STATE_V2");

  const out = ss.getSheetByName(CLEAN_OUTPUT_SHEET);
  if (out && out.getLastRow() > 1) {
    out.getRange(2, 1, out.getLastRow() - 1, out.getMaxColumns()).clearContent();
  }

  ss.toast("✅ Clean Master state reset. You can rebuild Clean Master from scratch now.", "Clean Master", 6);
  return "Clean Master state reset";
}

// ---------------------------
// PUBLIC: backfill order_date on All_Orders_Clean (repairs existing rows)
// ---------------------------
function backfillOrderDatesInCleanMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clean = ss.getSheetByName(CLEAN_OUTPUT_SHEET);
  if (!clean) throw new Error(`Missing "${CLEAN_OUTPUT_SHEET}" sheet.`);

  const shop = ss.getSheetByName("Shopify Orders");
  const sq = ss.getSheetByName("Squarespace Orders");
  if (!shop || !sq) throw new Error('Missing "Shopify Orders" and/or "Squarespace Orders" tabs.');

  const cleanData = clean.getDataRange().getValues();
  if (cleanData.length < 2) return "No clean rows to backfill.";

  const cleanHeaders = cleanData[0].map(h => String(h || "").trim());
  const colPlatform = cleanHeaders.indexOf("platform");
  const colOrderId  = cleanHeaders.indexOf("order_id");
  const colOrderDt  = cleanHeaders.indexOf("order_date");

  if (colPlatform === -1 || colOrderId === -1 || colOrderDt === -1) {
    throw new Error(`All_Orders_Clean missing required headers: platform / order_id / order_date`);
  }

  // Build date maps from raw sheets
  const shopMap = buildShopifyOrderDateMap_();
  const sqMap = buildSquarespaceOrderDateMap_();

  const rowsToUpdate = [];
  for (let r = 1; r < cleanData.length; r++) {
    const platform = String(cleanData[r][colPlatform] || "").trim();
    const orderId = String(cleanData[r][colOrderId] || "").trim();
    const existing = cleanData[r][colOrderDt];

    if (!platform || !orderId) continue;

    const hasDate = parseAnyDate_(existing);
    if (hasDate) continue;

    let d = null;
    if (platform === "Shopify") d = shopMap.get(orderId) || null;
    if (platform === "Squarespace") d = sqMap.get(orderId) || null;

    if (d) rowsToUpdate.push({ rowNum: r + 1, date: d });
  }

  if (!rowsToUpdate.length) {
    ss.toast("No blank order_date values found to backfill.", "Backfill", 6);
    return "No blank order_date values to backfill.";
  }

  // Write in contiguous batches (fast + safe)
  rowsToUpdate.sort((a, b) => a.rowNum - b.rowNum);

  let start = rowsToUpdate[0].rowNum;
  let prev = start;
  let buf = [rowsToUpdate[0].date];
  let writes = 0;

  function flush_() {
    clean.getRange(start, colOrderDt + 1, buf.length, 1).setValues(buf.map(d => [d]));
    writes++;
  }

  for (let i = 1; i < rowsToUpdate.length; i++) {
    const cur = rowsToUpdate[i].rowNum;
    if (cur === prev + 1) {
      buf.push(rowsToUpdate[i].date);
      prev = cur;
    } else {
      flush_();
      start = cur;
      prev = cur;
      buf = [rowsToUpdate[i].date];
    }
  }
  flush_();

  ss.toast(`✅ Backfilled order_date for ${rowsToUpdate.length} rows (${writes} write batches).`, "Backfill", 8);
  return `Backfilled order_date for ${rowsToUpdate.length} rows`;
}

// ---------------------------
// MAIN: Build Clean Master
// ---------------------------
function buildAllOrdersClean() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('Lock timeout: another process is running. Wait ~10 seconds and try again.');
  }

  const STATE_KEY = "CLEAN_MASTER_BUILD_STATE_V2";
  const SOFT_LIMIT_MS = 5.3 * 60 * 1000;
  const CHUNK_ROWS = 1500;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const shopSheet = ss.getSheetByName("Shopify Orders");
    const sqSheet = ss.getSheetByName("Squarespace Orders");
    if (!shopSheet || !sqSheet) throw new Error('Missing "Shopify Orders" and/or "Squarespace Orders" tabs.');

    const banned = loadBannedList_();
    const outSheet = getOrCreateSheetWithHeaders(CLEAN_OUTPUT_SHEET, CLEAN_HEADERS);

    let state = null;
    const rawState = PROPS.getProperty(STATE_KEY);
    if (rawState) {
      try { state = JSON.parse(rawState); } catch (e) { state = null; }
    }

    // start fresh
    if (!state || !state.started) {
      outSheet.getRange(1, 1, 1, CLEAN_HEADERS.length).setValues([CLEAN_HEADERS]);
      const lr = outSheet.getLastRow();
      if (lr > 1) outSheet.getRange(2, 1, lr - 1, outSheet.getMaxColumns()).clearContent();

      state = {
        started: true,
        phase: "shopify",
        rowCursor: 2,
        outRow: 2,
        excluded: 0,
        written: 0,
        lastOrderKey: ""
      };
      PROPS.setProperty(STATE_KEY, JSON.stringify(state));
    }

    const started = Date.now();
    const timeUp_ = () => (Date.now() - started) > SOFT_LIMIT_MS;

    function saveState_(msg) {
      PROPS.setProperty(STATE_KEY, JSON.stringify(state));
      ss.toast(msg, "Clean Master", 8);
      return msg;
    }

    function finish_() {
      PROPS.deleteProperty(STATE_KEY);

      outSheet.setFrozenRows(1);
      outSheet.getRange(1, 1, 1, CLEAN_HEADERS.length).setFontWeight("bold");

      const lr = outSheet.getLastRow();
      if (lr >= 2) {
        outSheet.getRange(2, 4, lr - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
        outSheet.getRange(2, 10, lr - 1, 1).setNumberFormat("0.00");
        outSheet.getRange(2, 11, lr - 1, 1).setNumberFormat("0.00");
        outSheet.getRange(2, 12, lr - 1, 1).setNumberFormat("0.00");
        outSheet.getRange(2, 13, lr - 1, 3).setNumberFormat("0.00");
      }

      logImportEvent("CleanMaster", `Built All_Orders_Clean (excluded:${state.excluded})`, state.written);
      ss.toast(`✅ Clean Master complete. Rows: ${state.written}. Excluded: ${state.excluded}.`, "Clean Master", 8);
      return `Built All_Orders_Clean (${state.written} rows), excluded ${state.excluded}`;
    }

    function processShopify_() {
      const lastRow = shopSheet.getLastRow();
      const lastCol = shopSheet.getLastColumn();
      if (lastRow < 2) return null;

      const headers = shopSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

      const c = {
        orderId: headers.indexOf("Order ID"),
        orderNumber: headers.indexOf("Order Number"),
        processedLocal: headers.indexOf("Processed At (Local)"),
        processedAt: headers.indexOf("Processed At"),
        createdLocal: headers.indexOf("Created At (Local)"),
        createdAt: headers.indexOf("Created At"),
        updatedAt: headers.indexOf("Updated At"),
        financial: headers.indexOf("Financial Status"),
        fulfill: headers.indexOf("Fulfillment Status"),
        currency: headers.indexOf("Currency"),
        totalDiscounts: headers.indexOf("Total Discounts"),
        totalPrice: headers.indexOf("Total Price"),
        currentTotalPrice: headers.indexOf("Current Total Price"),
        totalRefunds: headers.indexOf("Total Refunds"),
        testOrder: headers.indexOf("Test Order"),
        email: headers.indexOf("Customer Email"),
        first: headers.indexOf("Customer First Name"),
        last: headers.indexOf("Customer Last Name"),
        lineName: headers.indexOf("Lineitem Name"),
        lineQty: headers.indexOf("Lineitem Quantity"),
        linePrice: headers.indexOf("Lineitem Price"),
        lineSku: headers.indexOf("Lineitem SKU"),
        tags: headers.indexOf("Tags")
      };

      ["orderId","orderNumber","email","lineName","lineQty","linePrice"].forEach(k => {
        if (c[k] === -1) throw new Error(`Shopify Orders missing required column: ${k}`);
      });

      // If NONE of these exist, you will get blank order_date — fail loudly.
      const hasAnyDateCol =
        c.processedLocal >= 0 || c.processedAt >= 0 || c.createdLocal >= 0 || c.createdAt >= 0 || c.updatedAt >= 0;
      if (!hasAnyDateCol) {
        throw new Error(
          'Shopify Orders is missing date columns. Expected one of: "Processed At (Local)", "Processed At", "Created At (Local)", "Created At", "Updated At".'
        );
      }

      let r = state.rowCursor;
      while (r <= lastRow) {
        if (timeUp_()) return saveState_(`⏸️ Paused (timeout protection). Re-run “Build Clean Master” to continue. (${state.written} rows so far)`);

        const take = Math.min(CHUNK_ROWS, lastRow - r + 1);
        const values = shopSheet.getRange(r, 1, take, lastCol).getValues();

        const buffer = [];

        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          if (row.every(v => v === "" || v === null)) continue;

          if (c.testOrder >= 0 && truthy_(row[c.testOrder])) { state.excluded++; continue; }

          const emailRaw = s_(row[c.email]);
          if (emailRaw && isBannedEmail_(emailRaw, banned)) { state.excluded++; continue; }

          const orderId = s_(row[c.orderId]);
          const orderNumber = s_(row[c.orderNumber]);

          // ✅ robust date parse
          const orderDate =
            parseAnyDate_(c.processedLocal >= 0 ? row[c.processedLocal] : null) ||
            parseAnyDate_(c.processedAt >= 0 ? row[c.processedAt] : null) ||
            parseAnyDate_(c.createdLocal >= 0 ? row[c.createdLocal] : null) ||
            parseAnyDate_(c.createdAt >= 0 ? row[c.createdAt] : null) ||
            parseAnyDate_(c.updatedAt >= 0 ? row[c.updatedAt] : null) ||
            null;

          const first = c.first >= 0 ? s_(row[c.first]) : "";
          const last = c.last >= 0 ? s_(row[c.last]) : "";
          const customerName = (first || last) ? (first + " " + last).trim() : "";

          const productName = s_(row[c.lineName]);
          if (!productName) continue;

          // Exclude banned products
          if (isBannedProduct_(productName)) {
            state.excluded++;
            continue;
          }

          const sku = c.lineSku >= 0 ? s_(row[c.lineSku]) : "";
          const qty = parseQty_(row[c.lineQty]);
          const unitPrice = parseMoney_(row[c.linePrice]);
          const lineRevenue = (qty || 0) * (unitPrice || 0);

          const currency = c.currency >= 0 ? s_(row[c.currency]) : "";
          const financialStatus = c.financial >= 0 ? s_(row[c.financial]) : "";
          const fulfillmentStatus = c.fulfill >= 0 ? s_(row[c.fulfill]) : "";
          const tags = c.tags >= 0 ? s_(row[c.tags]) : "";

          const orderKey = "Shopify||" + orderId;
          const writeTotals = orderKey !== state.lastOrderKey;
          if (writeTotals) state.lastOrderKey = orderKey;

          const grossTotal = (c.totalPrice >= 0) ? Math.abs(parseMoney_(row[c.totalPrice])) : 0;
          const discountTotal = (c.totalDiscounts >= 0) ? Math.abs(parseMoney_(row[c.totalDiscounts])) : 0;
          const currentTotal = (c.currentTotalPrice >= 0) ? parseMoney_(row[c.currentTotalPrice]) : 0;
          const netRevenue = (currentTotal !== 0 || s_(row[c.currentTotalPrice]) !== "") ? Math.abs(currentTotal) : grossTotal;

          let refundTotal = 0;
          if (c.totalRefunds >= 0 && s_(row[c.totalRefunds]) !== "") {
            refundTotal = Math.abs(parseMoney_(row[c.totalRefunds]));
          } else {
            refundTotal = Math.max(0, grossTotal - netRevenue);
          }

          buffer.push([
            "Shopify",
            orderId,
            orderNumber,
            orderDate || "",
            emailRaw,
            normEmail_(emailRaw),
            customerName,
            productName,
            sku,
            qty,
            unitPrice,
            lineRevenue,
            writeTotals ? discountTotal : 0,
            writeTotals ? refundTotal : 0,
            writeTotals ? netRevenue : 0,
            currency,
            financialStatus,
            fulfillmentStatus,
            tags,
            "Shopify Orders"
          ]);
        }

        if (buffer.length) {
          outSheet.getRange(state.outRow, 1, buffer.length, CLEAN_HEADERS.length).setValues(buffer);
          state.outRow += buffer.length;
          state.written += buffer.length;
        }

        r += take;
        state.rowCursor = r;
        PROPS.setProperty(STATE_KEY, JSON.stringify(state));
      }

      state.phase = "squarespace";
      state.rowCursor = 2;
      state.lastOrderKey = "";
      PROPS.setProperty(STATE_KEY, JSON.stringify(state));
      return null;
    }

    // ✅ UPDATED: Squarespace pricing fix + EXCLUDE duplicate Montana/MT LLC renew* (2026) orders
    function processSquarespace_() {
      const lastRow = sqSheet.getLastRow();
      const lastCol = sqSheet.getLastColumn();
      if (lastRow < 2) return null;

      const headers = sqSheet.getRange(1, 1, 1, lastCol).getValues()[0]
        .map(h => String(h || "").trim());

      // case-insensitive header lookup
      const idx_ = (name) => headers.findIndex(h => String(h || "").trim().toLowerCase() === String(name).toLowerCase());

      const c = {
        orderId: idx_("Order ID"),
        orderNumber: idx_("Order Number"),
        createdOn: idx_("Created On") >= 0 ? idx_("Created On") : idx_("Created"),
        modifiedOn: idx_("Modified On") >= 0 ? idx_("Modified On") : idx_("Modified"),
        testMode: idx_("Test Mode"),
        email: idx_("Customer Email"),
        billFirst: idx_("Billing First Name"),
        billLast: idx_("Billing Last Name"),
        product: idx_("LineItem Product Name"),
        sku: idx_("LineItem SKU"),
        qty: idx_("LineItem Quantity"),
        unitPriceValue: idx_("LineItem Unit Price Value"),
        subtotalCurrency: idx_("Subtotal Currency"),
        discountValue: idx_("Discount Total Value"),
        refundedValue: idx_("Refunded Total Value"),
        grandTotalValue: idx_("Grand Total Value"),
        grandTotalCurrency: idx_("Grand Total Currency")
      };

      ["orderId","orderNumber","email","product","qty","grandTotalValue"].forEach(k => {
        if (c[k] === -1) throw new Error(`Squarespace Orders missing required column: ${k}`);
      });

      const hasAnyDateCol = (c.createdOn >= 0 || c.modifiedOn >= 0);
      if (!hasAnyDateCol) {
        throw new Error(
          'Squarespace Orders is missing date columns. Expected "Created On" and/or "Modified On" (case-insensitive).'
        );
      }

      // We write in chunks, but we MUST keep order lines together to allocate pricing correctly.
      let r = state.rowCursor;

      // Track Squarespace orders we want to exclude entirely
      const excludedOrderIds = new Set();

      // Hold current order lines across rows (and even chunk boundaries)
      let curOrderId = null;
      let curOrderLines = [];
      let curOrderMeta = null;

      const buffer = [];

      function normalizeProductText_(productName) {
        return String(productName || "")
          .toLowerCase()
          .replace(/[^a-z0-9\s]/g, " ")  // punctuation -> spaces
          .replace(/\s+/g, " ")
          .trim();
      }

      function shouldExcludeSquarespaceOrder_(orderDate, productName) {
        if (!orderDate || Object.prototype.toString.call(orderDate) !== "[object Date]" || isNaN(orderDate.getTime())) return false;

        const yr = orderDate.getFullYear();
        if (yr !== 2026) return false;

        const pn = normalizeProductText_(productName);

        // Keyword groups (AND across groups, OR inside group)
        // Must have (montana OR mt) AND llc AND renew*
        const groups = [
          ["montana", "mt"],
          ["llc"],
          ["renew"] // matches renewal/renewing/renewal etc.
        ];

        for (const group of groups) {
          const hit = group.some(kw => pn.includes(kw));
          if (!hit) return false;
        }
        return true;
      }

      function dropCurrentOrderAsExcluded_() {
        // Count any already-collected lines as excluded, then reset the order group so nothing writes.
        if (curOrderLines && curOrderLines.length) state.excluded += curOrderLines.length;
        curOrderId = null;
        curOrderLines = [];
        curOrderMeta = null;
      }

      function flushCurrentOrder_() {
        if (!curOrderId || !curOrderLines.length || !curOrderMeta) return;

        // If this order is marked excluded, do not write it
        if (excludedOrderIds.has(curOrderId)) {
          state.excluded += curOrderLines.length;
          curOrderId = null;
          curOrderLines = [];
          curOrderMeta = null;
          return;
        }

        // Fix missing Squarespace line pricing if needed
        fixSquarespaceLinePricing_(curOrderLines);

        // Write rows (totals once per order)
        for (let j = 0; j < curOrderLines.length; j++) {
          const lo = curOrderLines[j];
          const writeTotals = (j === 0);

          buffer.push([
            "Squarespace",
            curOrderMeta.orderId,
            curOrderMeta.orderNumber,
            curOrderMeta.orderDate || "",
            curOrderMeta.emailRaw,
            normEmail_(curOrderMeta.emailRaw),
            curOrderMeta.customerName,
            curOrderMeta.productNames[j],
            curOrderMeta.skus[j],
            lo.quantity,
            lo.unitPrice,
            lo.lineRevenue,
            writeTotals ? curOrderMeta.discountTotal : 0,
            writeTotals ? curOrderMeta.refundTotal : 0,
            writeTotals ? curOrderMeta.netRevenue : 0,
            curOrderMeta.currency,
            "",
            "",
            "",
            "Squarespace Orders"
          ]);
        }

        // Reset
        curOrderId = null;
        curOrderLines = [];
        curOrderMeta = null;
      }

      while (r <= lastRow) {
        if (timeUp_()) {
          // Flush what we have so far before saving state
          flushCurrentOrder_();

          if (buffer.length) {
            outSheet.getRange(state.outRow, 1, buffer.length, CLEAN_HEADERS.length).setValues(buffer);
            state.outRow += buffer.length;
            state.written += buffer.length;
          }

          state.rowCursor = r; // continue from here
          PROPS.setProperty("CLEAN_MASTER_BUILD_STATE_V2", JSON.stringify(state));
          return saveState_(`⏸️ Paused (timeout protection). Re-run “Build Clean Master” to continue. (${state.written} rows so far)`);
        }

        const take = Math.min(CHUNK_ROWS, lastRow - r + 1);
        const values = sqSheet.getRange(r, 1, take, lastCol).getValues();

        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          if (row.every(v => v === "" || v === null)) continue;

          // exclude Squarespace test mode orders
          if (c.testMode >= 0 && truthy_(row[c.testMode])) { state.excluded++; continue; }

          const emailRaw = s_(row[c.email]);
          if (emailRaw && isBannedEmail_(emailRaw, banned)) { state.excluded++; continue; }

          const orderId = s_(row[c.orderId]);
          const orderNumber = s_(row[c.orderNumber]);

          if (!orderId) continue;

          // If this orderId was already flagged to exclude, skip all remaining lines
          if (excludedOrderIds.has(orderId)) { state.excluded++; continue; }

          // robust date parse
          const orderDate =
            parseAnyDate_(c.createdOn >= 0 ? row[c.createdOn] : null) ||
            parseAnyDate_(c.modifiedOn >= 0 ? row[c.modifiedOn] : null) ||
            null;

          const first = c.billFirst >= 0 ? s_(row[c.billFirst]) : "";
          const last = c.billLast >= 0 ? s_(row[c.billLast]) : "";
          const customerName = (first || last) ? (first + " " + last).trim() : "";

          const productName = s_(row[c.product]);
          if (!productName) continue;

          // Exclude banned products
          if (isBannedProduct_(productName)) {
            state.excluded++;
            continue;
          }

          // ✅ NEW: Exclude duplicate Squarespace renewal orders in 2026
          if (shouldExcludeSquarespaceOrder_(orderDate, productName)) {
            excludedOrderIds.add(orderId);

            // If we're currently building this same order, drop already-collected lines
            if (curOrderId && curOrderId === orderId) {
              dropCurrentOrderAsExcluded_();
            }

            // Exclude this line too (and all subsequent ones)
            state.excluded++;
            continue;
          }

          const sku = c.sku >= 0 ? s_(row[c.sku]) : "";
          const qty = parseQty_(row[c.qty]);

          const unitPrice = (c.unitPriceValue >= 0 && s_(row[c.unitPriceValue]) !== "")
            ? parseMoney_(row[c.unitPriceValue])
            : 0;

          const lineRevenue = (qty || 0) * (unitPrice || 0);

          const discountTotal = (c.discountValue >= 0) ? Math.abs(parseMoney_(row[c.discountValue])) : 0;
          const refundTotal = (c.refundedValue >= 0) ? Math.abs(parseMoney_(row[c.refundedValue])) : 0;
          const grandTotal = Math.abs(parseMoney_(row[c.grandTotalValue]));
          const netRevenue = Math.max(0, grandTotal - refundTotal);

          const currency =
            (c.subtotalCurrency >= 0 ? s_(row[c.subtotalCurrency]) : "") ||
            (c.grandTotalCurrency >= 0 ? s_(row[c.grandTotalCurrency]) : "");

          // If order changes, flush previous order before starting new one
          if (curOrderId && orderId !== curOrderId) {
            flushCurrentOrder_();
          }

          // Start new order group if needed
          if (!curOrderId) {
            curOrderId = orderId;
            curOrderMeta = {
              orderId,
              orderNumber,
              orderDate,
              emailRaw,
              customerName,
              discountTotal,
              refundTotal,
              netRevenue,
              currency,
              productNames: [],
              skus: []
            };
          }

          // Append this line to current order group
          curOrderMeta.productNames.push(productName);
          curOrderMeta.skus.push(sku);

          curOrderLines.push({
            platform: "Squarespace",
            orderId: orderId,
            quantity: qty,
            unitPrice: unitPrice,
            lineRevenue: lineRevenue,
            orderNet: netRevenue
          });
        }

        r += take;
        state.rowCursor = r;

        // Write buffer periodically to avoid huge memory usage
        if (buffer.length >= 2000) {
          // Flush current order first so lines are not split
          flushCurrentOrder_();

          if (buffer.length) {
            outSheet.getRange(state.outRow, 1, buffer.length, CLEAN_HEADERS.length).setValues(buffer);
            state.outRow += buffer.length;
            state.written += buffer.length;
            buffer.length = 0;
          }

          PROPS.setProperty("CLEAN_MASTER_BUILD_STATE_V2", JSON.stringify(state));
        }
      }

      // End of sheet: flush any remaining order
      flushCurrentOrder_();

      if (buffer.length) {
        outSheet.getRange(state.outRow, 1, buffer.length, CLEAN_HEADERS.length).setValues(buffer);
        state.outRow += buffer.length;
        state.written += buffer.length;
      }

      return null;
    }

    if (state.phase === "shopify") {
      const msg = processShopify_();
      if (msg) return msg;
    }
    if (state.phase === "squarespace") {
      const msg = processSquarespace_();
      if (msg) return msg;
    }

    return finish_();

  } finally {
    lock.releaseLock();
  }
}

// ---------------------------
// INTERNAL: Build order date maps from raw sheets
// ---------------------------
function buildShopifyOrderDateMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Shopify Orders");
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

  const colOrderId = headers.indexOf("Order ID");
  const colProcessedLocal = headers.indexOf("Processed At (Local)");
  const colProcessedAt = headers.indexOf("Processed At");
  const colCreatedLocal = headers.indexOf("Created At (Local)");
  const colCreatedAt = headers.indexOf("Created At");
  const colUpdatedAt = headers.indexOf("Updated At");

  if (colOrderId === -1) return map;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const oid = String(row[colOrderId] || "").trim();
    if (!oid) continue;

    const d =
      parseAnyDate_(colProcessedLocal >= 0 ? row[colProcessedLocal] : null) ||
      parseAnyDate_(colProcessedAt >= 0 ? row[colProcessedAt] : null) ||
      parseAnyDate_(colCreatedLocal >= 0 ? row[colCreatedLocal] : null) ||
      parseAnyDate_(colCreatedAt >= 0 ? row[colCreatedAt] : null) ||
      parseAnyDate_(colUpdatedAt >= 0 ? row[colUpdatedAt] : null) ||
      null;

    if (d && !map.has(oid)) map.set(oid, d);
  }
  return map;
}

function buildSquarespaceOrderDateMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Squarespace Orders");
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim());

  const idx_ = (name) => headers.findIndex(h => String(h || "").trim().toLowerCase() === String(name).toLowerCase());

  const colOrderId = idx_("Order ID");
  const colCreatedOn = idx_("Created On") >= 0 ? idx_("Created On") : idx_("Created");
  const colModifiedOn = idx_("Modified On") >= 0 ? idx_("Modified On") : idx_("Modified");

  if (colOrderId === -1) return map;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const oid = String(row[colOrderId] || "").trim();
    if (!oid) continue;

    const d =
      parseAnyDate_(colCreatedOn >= 0 ? row[colCreatedOn] : null) ||
      parseAnyDate_(colModifiedOn >= 0 ? row[colModifiedOn] : null) ||
      null;

    if (d && !map.has(oid)) map.set(oid, d);
  }
  return map;
}

// ---------------------------
// INTERNAL: robust date parsing
// ---------------------------
function parseAnyDate_(v) {
  if (v === null || v === undefined || v === "") return null;

  // already a Date from Sheets
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;

  // numeric serial fallback
  if (typeof v === "number" && isFinite(v)) {
    // Google Sheets serial days since 1899-12-30
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }

  const s = String(v).trim().replace(/^'+/, "");
  if (!s) return null;

  // ISO or parseable string
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // "YYYY-MM-DD HH:MM:SS" -> make ISO-ish
  if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}/.test(s)) {
    d = new Date(s.replace(" ", "T"));
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

/**
 * Fixes missing Squarespace line pricing:
 * - If unit_price/line_revenue are 0 but order_net_revenue > 0,
 *   allocate order_net across the order's lines (weighted by quantity).
 *
 * rows: array of line objects for a single order:
 *   { platform, orderId, quantity, unitPrice, lineRevenue, orderNet }
 *
 * Mutates rows in place.
 */
function fixSquarespaceLinePricing_(rows) {
  if (!rows || !rows.length) return;

  // Only apply to Squarespace
  const platform = String(rows[0].platform || "").trim();
  if (platform !== "Squarespace") return;

  // If any line already has revenue, don't overwrite—only fill missing lines
  let hasAnyLineRevenue = false;
  for (const r of rows) {
    if ((Number(r.lineRevenue) || 0) > 0 || (Number(r.unitPrice) || 0) > 0) {
      hasAnyLineRevenue = true;
      break;
    }
  }

  const orderNet = Number(rows[0].orderNet) || 0;
  if (orderNet <= 0) return;

  // If we have some pricing already, only fill blanks using unitPrice*qty if possible
  if (hasAnyLineRevenue) {
    for (const r of rows) {
      const qty = Math.max(1, Number(r.quantity) || 1);
      const up = Number(r.unitPrice) || 0;
      const lr = Number(r.lineRevenue) || 0;

      if (lr <= 0 && up > 0) r.lineRevenue = up * qty;
      if (up <= 0 && lr > 0) r.unitPrice = lr / qty;
    }
    return;
  }

  // Otherwise: allocate the full orderNet across lines weighted by quantity
  let qtySum = 0;
  for (const r of rows) qtySum += Math.max(1, Number(r.quantity) || 1);
  if (qtySum <= 0) qtySum = rows.length;

  for (const r of rows) {
    const qty = Math.max(1, Number(r.quantity) || 1);
    const share = orderNet * (qty / qtySum);
    r.lineRevenue = share;
    r.unitPrice = share / qty;
  }
}
