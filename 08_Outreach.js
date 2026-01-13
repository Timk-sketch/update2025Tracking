// =====================================================
// 08_Outreach.gs — Customer Outreach List
// =====================================================

function ensureOutreachControlsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OUTREACH_CONTROLS_SHEET);
  if (!sh) sh = ss.insertSheet(OUTREACH_CONTROLS_SHEET);

  const defaults = [
    ["Customer Outreach List Controls", ""],
    ["LTV Threshold (Good Customer)", 5000],
    ["Aggressiveness Multiplier", 1.50],
    ["Max Days Since Last Order", 365],
    ["Min Total Orders", 2],
    ["Min LTV (Optional override)", ""]
  ];

  sh.getRange(1, 1, defaults.length, 2).setValues(defaults);
  sh.getRange(1, 1).setFontWeight("bold");
  sh.autoResizeColumns(1, 2);

  return sh;
}

function getOutreachControls_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OUTREACH_CONTROLS_SHEET) || ensureOutreachControlsSheet_();
  const values = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 2).getValues();

  const map = {};
  for (let i = 0; i < values.length; i++) {
    const k = s_(values[i][0]);
    const v = values[i][1];
    if (!k) continue;
    map[k] = v;
  }

  const ltvThreshold = n_(map["LTV Threshold (Good Customer)"] ?? 5000);
  const aggressiveness = n_(map["Aggressiveness Multiplier"] ?? 1.5) || 1.5;
  const maxDays = n_(map["Max Days Since Last Order"] ?? 365) || 365;
  const minOrders = n_(map["Min Total Orders"] ?? 2) || 2;
  const minLtvOverride = map["Min LTV (Optional override)"] === "" ? null : n_(map["Min LTV (Optional override)"]);

  return { ltvThreshold, aggressiveness, maxDays, minOrders, minLtvOverride };
}

function setOutreachControlsFromSidebar(ltvThreshold, aggressiveness, maxDays, minOrders, minLtvOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OUTREACH_CONTROLS_SHEET) || ensureOutreachControlsSheet_();

  const rows = [
    ["Customer Outreach List Controls", ""],
    ["LTV Threshold (Good Customer)", n_(ltvThreshold)],
    ["Aggressiveness Multiplier", n_(aggressiveness)],
    ["Max Days Since Last Order", n_(maxDays)],
    ["Min Total Orders", n_(minOrders)],
    ["Min LTV (Optional override)", (minLtvOverride === "" || minLtvOverride == null) ? "" : n_(minLtvOverride)]
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  ss.toast("Outreach controls saved.", "Outreach", 4);
  return "Outreach controls saved.";
}

function getOutreachControlsForSidebar() {
  const c = getOutreachControls_();
  return {
    ltvThreshold: c.ltvThreshold,
    aggressiveness: c.aggressiveness,
    maxDays: c.maxDays,
    minOrders: c.minOrders,
    minLtvOverride: (c.minLtvOverride == null ? "" : c.minLtvOverride)
  };
}

function buildCustomerOutreachList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clean = ss.getSheetByName(CLEAN_OUTPUT_SHEET);
  if (!clean) throw new Error(`Missing "${CLEAN_OUTPUT_SHEET}" tab.`);

  const controls = getOutreachControls_();
  const ltvFloor = (controls.minLtvOverride != null && controls.minLtvOverride > 0)
    ? controls.minLtvOverride
    : controls.ltvThreshold;

  const data = clean.getDataRange().getValues();
  if (data.length < 2) throw new Error(`No data found in "${CLEAN_OUTPUT_SHEET}".`);

  const headers = data[0].map(h => String(h || "").trim());
  const hm = {};
  headers.forEach((h, i) => hm[h] = i);

  const COL = {
    email: mustColStrict_(hm, "customer_email_norm"),
    orderId: mustColStrict_(hm, "order_id"),
    orderDate: mustColStrict_(hm, "order_date"),
    product: mustColStrict_(hm, "product_name"),
    qty: mustColStrict_(hm, "quantity"),
    rev: mustColStrict_(hm, "line_revenue")
  };

  const byCustomer = new Map();

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const email = s_(row[COL.email]).toLowerCase();
    if (!email) continue;

    const dt = asDate_(row[COL.orderDate]);
    if (!dt) continue;

    const orderId = s_(row[COL.orderId]);
    const product = s_(row[COL.product]);
    const qty = n_(row[COL.qty]);
    const rev = n_(row[COL.rev]);

    let c = byCustomer.get(email);
    if (!c) {
      c = {
        orders: new Set(),
        dates: [],
        ltv: 0,
        prodRev: new Map(),
        prodUnits: new Map(),
        prodOrders: new Map()
      };
      byCustomer.set(email, c);
    }

    if (orderId) c.orders.add(orderId);
    c.dates.push(dt);
    c.ltv += rev;

    if (product) {
      c.prodRev.set(product, (c.prodRev.get(product) || 0) + rev);
      c.prodUnits.set(product, (c.prodUnits.get(product) || 0) + qty);

      if (!c.prodOrders.has(product)) c.prodOrders.set(product, new Set());
      if (orderId) c.prodOrders.get(product).add(orderId);
    }
  }

  const today = new Date();
  const rowsOut = [];

  byCustomer.forEach((c, email) => {
    const totalOrders = c.orders.size;
    if (totalOrders < controls.minOrders) return;
    if (c.ltv < ltvFloor) return;

    c.dates.sort((a, b) => a.getTime() - b.getTime());
    const first = c.dates[0];
    const last = c.dates[c.dates.length - 1];

    const daysAsCustomer = Math.max(0, Math.round(daysBetween_(first, today)));
    const daysSinceLast = Math.max(0, Math.round(daysBetween_(last, today)));

    // ✅ HARD CUTOFF: respect "Max Days Since Last Order"
    // If the customer is older than maxDays since last purchase, exclude them entirely.
    if (daysSinceLast > controls.maxDays) return;

    let normalGap = null;
    if (c.dates.length >= 2) {
      let sum = 0;
      let count = 0;
      for (let i = 1; i < c.dates.length; i++) {
        const d = daysBetween_(c.dates[i - 1], c.dates[i]);
        if (isFinite(d) && d >= 0) { sum += d; count++; }
      }
      if (count > 0) normalGap = sum / count;
    }

    const targetGap = (normalGap != null ? (controls.aggressiveness * normalGap) : controls.maxDays);
    const triggerGap = Math.min(controls.maxDays, Math.max(1, targetGap));
    if (daysSinceLast < triggerGap) return;

    const avgOrder = totalOrders > 0 ? (c.ltv / totalOrders) : 0;

    const topByRev = topKeyValue_(c.prodRev);
    const topByUnits = topKeyValue_(c.prodUnits);
    const topByOrders = topKeySetSize_(c.prodOrders);

    const gapScore = Math.min(1, daysSinceLast / triggerGap);
    const ltvScore = Math.min(1, c.ltv / Math.max(1, controls.ltvThreshold * 2));
    const orderScore = Math.min(1, totalOrders / 10);
    const priority = Math.max(0, Math.min(100, Math.round(100 * (0.60 * gapScore + 0.30 * ltvScore + 0.10 * orderScore))));

    rowsOut.push([
      priority,
      email,
      c.ltv,
      totalOrders,
      avgOrder,
      formatDate_(first),
      formatDate_(last),
      daysAsCustomer,
      daysSinceLast,
      (normalGap == null ? "" : Math.round(normalGap)),
      Math.round(triggerGap),
      topByRev.key || "",
      topByRev.val || 0,
      topByUnits.key || "",
      topByUnits.val || 0,
      topByOrders.key || "",
      topByOrders.val || 0
    ]);
  });

  rowsOut.sort((a, b) => (b[0] || 0) - (a[0] || 0));

  const outHeaders = [
    "Contact Priority (0-100)",
    "Customer Email",
    "Lifetime Value (All-time)",
    "Total Orders (All-time)",
    "Avg Order Amount (All-time)",
    "First Purchase Date",
    "Most Recent Purchase Date",
    "Days as Customer",
    "Days Since Last Order",
    "Normal Gap (avg days)",
    "Trigger Gap (days)",
    "Top Product by Revenue",
    "Top Product Revenue",
    "Top Product by Units",
    "Top Product Units",
    "Top Product by Orders",
    "Top Product Orders"
  ];

  const ss2 = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss2.getSheetByName(OUTREACH_OUTPUT_SHEET) || ss2.insertSheet(OUTREACH_OUTPUT_SHEET);
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]).setFontWeight("bold");

  if (rowsOut.length) {
    outSheet.getRange(2, 1, rowsOut.length, outHeaders.length).setValues(rowsOut);
    outSheet.getRange(2, 3, rowsOut.length, 1).setNumberFormat("$#,##0.00");
    outSheet.getRange(2, 5, rowsOut.length, 1).setNumberFormat("$#,##0.00");
    outSheet.getRange(2, 13, rowsOut.length, 1).setNumberFormat("$#,##0.00");
  }

  outSheet.setFrozenRows(1);
  outSheet.autoResizeColumns(1, outHeaders.length);

  ss2.toast(`Customer Outreach List built: ${rowsOut.length} customers`, "Outreach", 6);
  logImportEvent("Outreach", "Built Customer_Outreach_List", rowsOut.length);

  return `Customer Outreach List built (${rowsOut.length} customers)`;
}

function topKeyValue_(map) {
  let bestK = "";
  let bestV = -Infinity;
  map.forEach((v, k) => {
    if ((v || 0) > bestV) { bestV = (v || 0); bestK = k; }
  });
  if (bestV === -Infinity) bestV = 0;
  return { key: bestK, val: bestV };
}

function topKeySetSize_(mapOfSets) {
  let bestK = "";
  let bestV = -Infinity;
  mapOfSets.forEach((set, k) => {
    const v = set ? set.size : 0;
    if (v > bestV) { bestV = v; bestK = k; }
  });
  if (bestV === -Infinity) bestV = 0;
  return { key: bestK, val: bestV };
}
