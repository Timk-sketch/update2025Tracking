// =====================================================
// 07_SummaryReport.gs — Orders Summary Report
// =====================================================

const OSR_CFG = {
  CLEAN_SHEET: "All_Orders_Clean",
  OUTPUT_SHEET: "Orders_Summary_Report",
  START_CELL: "B2", // must remain a real Date
  END_CELL: "D2",   // must remain a real Date
  TOP_PRODUCTS_N: 15,
  RETURNING_CUSTOMERS_N: 50
};

function buildOrdersSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  logProgress('Summary Report', 'Reading clean orders data...');

  const clean = ss.getSheetByName(OSR_CFG.CLEAN_SHEET);
  if (!clean) throw new Error(`Missing "${OSR_CFG.CLEAN_SHEET}" tab.`);

  const outSheet = ss.getSheetByName(OSR_CFG.OUTPUT_SHEET) || ss.insertSheet(OSR_CFG.OUTPUT_SHEET);

  // Read date range FIRST (before clearing)
  const { startDate, endDate } = getDateRange_(outSheet);
  if (!startDate || !endDate) {
    throw new Error(`Date range not found. Enter start date in ${OSR_CFG.OUTPUT_SHEET}!${OSR_CFG.START_CELL} and end date in ${OSR_CFG.OUTPUT_SHEET}!${OSR_CFG.END_CELL}.`);
  }

  logProgress('Summary Report', `Processing orders from ${formatDate_(startDate)} to ${formatDate_(endDate)}...`);

  const data = clean.getDataRange().getValues();
  if (data.length < 2) throw new Error(`No data found in "${OSR_CFG.CLEAN_SHEET}".`);

  const headers = data[0].map(h => String(h || "").trim());
  const hm = {};
  headers.forEach((h, i) => (hm[h] = i));

  const COL = {
    platform: mustColStrict_(hm, "platform"),
    order_id: mustColStrict_(hm, "order_id"),
    order_date: mustColStrict_(hm, "order_date"),
    email: mustColStrict_(hm, "customer_email_norm"),
    product: mustColStrict_(hm, "product_name"),
    qty: mustColStrict_(hm, "quantity"),
    line_rev: mustColStrict_(hm, "line_revenue"),
    order_discount: mustColStrict_(hm, "order_discount_total"),
    order_refund: mustColStrict_(hm, "order_refund_total"),
    order_net: mustColStrict_(hm, "order_net_revenue")
  };

  const orderAgg = new Map();   // platform||orderId -> {grossLines, discount, refund, net, units}
  const byProduct = new Map();  // platform||product -> {units, revenue}
  const byCustomer = new Map(); // email -> { first, last, lifetimeRev, periodRev, hadBefore, hadInPeriod }

  // Diagnostics (helps verify Squarespace inclusion)
  let linesInPeriodShopify = 0;
  let linesInPeriodSquarespace = 0;
  let squarespaceBlankProductLines = 0;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    const platformRaw = s_(row[COL.platform]);
    const platform = normalizePlatform_(platformRaw); // ✅ normalized
    const orderId = s_(row[COL.order_id]);
    const email = s_(row[COL.email]).toLowerCase();

    // ✅ product fallback: don't drop Squarespace just because product_name is blank
    const productRaw = s_(row[COL.product]);
    const product = productRaw || "(Unknown Product)";

    const dt = asDate_(row[COL.order_date]);
    if (!dt) continue;

    const qty = n_(row[COL.qty]);
    const rev = n_(row[COL.line_rev]);

    // CUSTOMER: all-time + cohort flags (date-range sensitive)
    if (email) {
      let c = byCustomer.get(email);
      if (!c) {
        c = { first: dt, last: dt, lifetimeRev: 0, periodRev: 0, hadBefore: false, hadInPeriod: false };
        byCustomer.set(email, c);
      } else {
        if (dt < c.first) c.first = dt;
        if (dt > c.last) c.last = dt;
      }

      c.lifetimeRev += rev;

      if (dt < startDate) c.hadBefore = true;
      if (isWithinRangeInclusive_(dt, startDate, endDate)) c.hadInPeriod = true;
    }

    // Stop if not in period
    if (!isWithinRangeInclusive_(dt, startDate, endDate)) continue;

    if (platform === "Shopify") linesInPeriodShopify++;
    if (platform === "Squarespace") linesInPeriodSquarespace++;
    if (platform === "Squarespace" && !productRaw) squarespaceBlankProductLines++;

    // Period revenue per customer (only in range)
    if (email) {
      const c = byCustomer.get(email);
      if (c) c.periodRev += rev;
    }

    // PRODUCT rollup (✅ now includes Squarespace even if product_name missing)
    if (platform) {
      const keyP = platform + "||" + product;
      let p = byProduct.get(keyP);
      if (!p) {
        p = { platform, product, units: 0, revenue: 0 };
        byProduct.set(keyP, p);
      }
      p.units += qty;
      p.revenue += rev;
    }

    // ORDER rollup
    if (platform && orderId) {
      const keyO = platform + "||" + orderId;
      let o = orderAgg.get(keyO);
      if (!o) {
        o = { platform, orderId, grossLines: 0, discount: 0, refund: 0, net: 0, units: 0 };
        orderAgg.set(keyO, o);

        // ✅ FIX: Only add order-level totals ONCE when we first see the order
        // These are order-level values that repeat for every line item in All_Orders_Clean
        o.discount = n_(row[COL.order_discount]);
        o.refund = n_(row[COL.order_refund]);
        o.net = n_(row[COL.order_net]);
      }

      // Always accumulate line-level values
      o.grossLines += rev;
      o.units += qty;
    }
  }

  let totalOrdersAllSources = 0;
  let grossRevenuePeriod = 0;
  let discountTotalPeriod = 0;
  let refundTotalPeriod = 0;
  let netRevenuePeriod = 0;
  let periodUnits = 0;

  const bySource = new Map(); // platform -> {orders, gross, discount, refund, net, units}

  orderAgg.forEach(o => {
    totalOrdersAllSources += 1;

    const net = (o.net && o.net > 0) ? o.net : Math.max(0, o.grossLines - o.discount - o.refund);
    const gross = (o.net && o.net > 0) ? (net + o.discount + o.refund) : o.grossLines;

    grossRevenuePeriod += gross;
    discountTotalPeriod += o.discount;
    refundTotalPeriod += o.refund;
    netRevenuePeriod += net;
    periodUnits += o.units;

    let s = bySource.get(o.platform);
    if (!s) {
      s = { orders: 0, gross: 0, discount: 0, refund: 0, net: 0, units: 0 };
      bySource.set(o.platform, s);
    }
    s.orders += 1;
    s.gross += gross;
    s.discount += o.discount;
    s.refund += o.refund;
    s.net += net;
    s.units += o.units;
  });

  const aovGross = totalOrdersAllSources > 0 ? (grossRevenuePeriod / totalOrdersAllSources) : 0;
  const aovNet = totalOrdersAllSources > 0 ? (netRevenuePeriod / totalOrdersAllSources) : 0;
  const unitsPerOrder = totalOrdersAllSources > 0 ? (periodUnits / totalOrdersAllSources) : 0;

  // cohorts
  let newCustomers = 0;
  let returningCustomers = 0;
  let returningLtvAllTime = 0;
  const returningRows = [];

  byCustomer.forEach((c, email) => {
    if (!c.hadInPeriod || c.periodRev <= 0) return;

    if (c.hadBefore) {
      returningCustomers++;
      returningLtvAllTime += c.lifetimeRev;
      returningRows.push([email, c.periodRev, c.lifetimeRev, formatDate_(c.first), formatDate_(c.last)]);
    } else {
      newCustomers++;
    }
  });

  const uniqueCustomersPeriod = newCustomers + returningCustomers;
  const avgRevenuePerCustomerPeriod = uniqueCustomersPeriod > 0 ? (grossRevenuePeriod / uniqueCustomersPeriod) : 0;

  // top products
  const products = Array.from(byProduct.values()).sort((a, b) => (b.revenue || 0) - (a.revenue || 0));
  const topProducts = products
    .slice(0, OSR_CFG.TOP_PRODUCTS_N)
    .map((p, i) => ([i + 1, p.platform, p.product, p.units, p.revenue]));

  // returning top list
  returningRows.sort((a, b) => (b[2] || 0) - (a[2] || 0));
  const returningTop = returningRows.slice(0, OSR_CFG.RETURNING_CUSTOMERS_N);

  // marketing
  const m = getMarketingControls_();
  const marketingSpend = m.marketingSpend || 0;
  const contributionMarginPct = (m.contributionMarginPct || 0);
  const targetSpendPct = (m.targetSpendPct || 0);

  const spendPctOfRevenue = netRevenuePeriod > 0 ? (marketingSpend / netRevenuePeriod) : 0;
  const spendPctVariance = spendPctOfRevenue - targetSpendPct;
  const mer = marketingSpend > 0 ? (netRevenuePeriod / marketingSpend) : 0;
  const cac = newCustomers > 0 ? (marketingSpend / newCustomers) : 0;
  const breakEvenMargin = (aovNet > 0 ? (cac / aovNet) : 0);
  const paybackOrders = (aovNet > 0 && contributionMarginPct > 0) ? (cac / (aovNet * contributionMarginPct)) : 0;

  // render
  outSheet.clearContents();
  outSheet.clearFormats();

  // Title
  outSheet.getRange(1, 1).setValue("Orders Summary Report").setFontWeight("bold").setFontSize(14);

  // Keep B2 and D2 as TRUE date inputs
  outSheet.getRange("A2").setValue("Start Date").setFontWeight("bold");
  outSheet.getRange("B2").setValue(startDate).setNumberFormat("yyyy-mm-dd");
  outSheet.getRange("C2").setValue("End Date").setFontWeight("bold");
  outSheet.getRange("D2").setValue(endDate).setNumberFormat("yyyy-mm-dd");

  // Display period string somewhere else
  outSheet.getRange("A3").setValue("Analysis Period:").setFontWeight("bold");
  outSheet.getRange("B3").setValue(`${formatDate_(startDate)} to ${formatDate_(endDate)}`);
  outSheet.getRange("B3:D3").merge();

  const kpis = [
    ["Total Orders (All Sources)", totalOrdersAllSources],
    ["Gross Revenue (Period)", grossRevenuePeriod],
    ["Discounts (Period)", discountTotalPeriod],
    ["Refunds (Period)", refundTotalPeriod],
    ["Net Revenue (Period)", netRevenuePeriod],
    ["Total Units (Period)", periodUnits],
    ["AOV (Gross)", aovGross],
    ["AOV (Net)", aovNet],
    ["Units per Order (Period)", unitsPerOrder],
    ["New Customers (Period)", newCustomers],
    ["Returning Customers (Purchased in Period)", returningCustomers],
    ["Unique Customers (Period)", uniqueCustomersPeriod],
    ["Avg Revenue per Customer (Period)", avgRevenuePerCustomerPeriod],
    ["Returning Customers: Total LTV (All-time)", returningLtvAllTime],

    // ✅ Diagnostics to prove Squarespace is included in Top Products
    ["Lines In Period — Shopify", linesInPeriodShopify],
    ["Lines In Period — Squarespace", linesInPeriodSquarespace],
    ["Squarespace Lines With Blank Product", squarespaceBlankProductLines]
  ];

  outSheet.getRange(4, 1).setValue("Revenue / Customer KPIs").setFontWeight("bold");
  outSheet.getRange(4, 1, 1, 2).merge();

  const kpiStartRow = 5;
  outSheet.getRange(kpiStartRow, 1, kpis.length, 2).setValues(kpis);
  outSheet.getRange(kpiStartRow, 1, kpis.length, 1).setFontWeight("bold");

  const currencyLabels = new Set([
    "Gross Revenue (Period)",
    "Discounts (Period)",
    "Refunds (Period)",
    "Net Revenue (Period)",
    "AOV (Gross)",
    "AOV (Net)",
    "Avg Revenue per Customer (Period)",
    "Returning Customers: Total LTV (All-time)"
  ]);

  for (let i = 0; i < kpis.length; i++) {
    const label = kpis[i][0];
    const cell = outSheet.getRange(kpiStartRow + i, 2);
    if (currencyLabels.has(label)) cell.setNumberFormat("$#,##0.00");
    else if (label === "Units per Order (Period)") cell.setNumberFormat("0.00");
    else cell.setNumberFormat("0");
  }

  outSheet.getRange(4, 4).setValue("Marketing Efficiency (Period)").setFontWeight("bold");
  outSheet.getRange(4, 4, 1, 2).merge();

  const marketingRows = [
    ["Marketing Spend ($)", marketingSpend],
    ["Marketing Spend % of Net Revenue", spendPctOfRevenue],
    ["Target Spend % of Revenue", targetSpendPct],
    ["Spend % Variance vs Target", spendPctVariance],
    ["MER (Net Revenue / Spend)", mer],
    ["CAC (Spend / New Customers)", cac],
    ["Break-even Margin on 1st Order (CAC / Net AOV)", breakEvenMargin],
    ["Contribution Margin (%) — provided", contributionMarginPct],
    ["Payback (orders) @ provided margin", paybackOrders]
  ];

  const marketingStartRow = 5;
  outSheet.getRange(marketingStartRow, 4, marketingRows.length, 2).setValues(marketingRows);
  outSheet.getRange(marketingStartRow, 4, marketingRows.length, 1).setFontWeight("bold");

  for (let i = 0; i < marketingRows.length; i++) {
    const label = marketingRows[i][0];
    const cell = outSheet.getRange(marketingStartRow + i, 5);
    if (label === "Marketing Spend ($)" || label === "CAC (Spend / New Customers)") cell.setNumberFormat("$#,##0.00");
    else if (label.includes("%") || label.includes("Margin")) cell.setNumberFormat("0.00%");
    else cell.setNumberFormat("0.00");
  }

  let row = Math.max(kpiStartRow + kpis.length, marketingStartRow + marketingRows.length) + 2;

  outSheet.getRange(row, 1).setValue("Orders by Source").setFontWeight("bold");
  row++;

  outSheet.getRange(row, 1, 1, 6)
    .setValues([["Source", "Orders", "Gross Revenue", "Discounts", "Refunds", "Net Revenue"]])
    .setFontWeight("bold");
  row++;

  const platforms = ["Shopify", "Squarespace"];
  const sourceRows = [];
  let totOrders = 0, totGross = 0, totDisc = 0, totRef = 0, totNet = 0;

  platforms.forEach(p => {
    const s = bySource.get(p) || { orders: 0, gross: 0, discount: 0, refund: 0, net: 0 };
    sourceRows.push([p, s.orders, s.gross, s.discount, s.refund, s.net]);
    totOrders += s.orders; totGross += s.gross; totDisc += s.discount; totRef += s.refund; totNet += s.net;
  });

  sourceRows.push(["TOTAL", totOrders, totGross, totDisc, totRef, totNet]);

  outSheet.getRange(row, 1, sourceRows.length, 6).setValues(sourceRows);
  outSheet.getRange(row, 2, sourceRows.length, 1).setNumberFormat("0");
  outSheet.getRange(row, 3, sourceRows.length, 4).setNumberFormat("$#,##0.00");

  row += sourceRows.length + 2;

  outSheet.getRange(row, 1).setValue("Top 15 Products (Revenue + Source)").setFontWeight("bold");
  row++;
  outSheet.getRange(row, 1, 1, 5).setValues([["Rank", "Source", "Product", "Units", "Revenue"]]).setFontWeight("bold");
  row++;

  if (topProducts.length) {
    outSheet.getRange(row, 1, topProducts.length, 5).setValues(topProducts);
    outSheet.getRange(row, 4, topProducts.length, 1).setNumberFormat("0.00");
    outSheet.getRange(row, 5, topProducts.length, 1).setNumberFormat("$#,##0.00");
    row += topProducts.length + 2;
  } else {
    row += 2;
  }

  outSheet.getRange(row, 1).setValue("Returning Customers (Top 50 by LTV)").setFontWeight("bold");
  row++;

  outSheet.getRange(row, 1, 1, 5).setValues([[
    "Customer Email",
    "Period Revenue",
    "Lifetime Value (All-time)",
    "First Purchase Date",
    "Most Recent Purchase Date"
  ]]).setFontWeight("bold");
  row++;

  if (returningTop.length) {
    outSheet.getRange(row, 1, returningTop.length, 5).setValues(returningTop);
    outSheet.getRange(row, 2, returningTop.length, 1).setNumberFormat("$#,##0.00");
    outSheet.getRange(row, 3, returningTop.length, 1).setNumberFormat("$#,##0.00");
  }

  outSheet.setFrozenRows(3);
  outSheet.autoResizeColumns(1, 10);

  const summaryMsg = `Orders Summary built: ${totalOrdersAllSources} orders (${formatDate_(startDate)} to ${formatDate_(endDate)})`;
  logProgress('Summary Report', summaryMsg);
  logImportEvent("Summary", `Built Orders_Summary_Report (${formatDate_(startDate)} to ${formatDate_(endDate)})`, totalOrdersAllSources);
  return summaryMsg;
}

function mustColStrict_(hm, name) {
  const idx = hm[name];
  if (idx === undefined || idx === null || idx < 0) throw new Error(`Missing column "${name}" in All_Orders_Clean.`);
  return idx;
}

function getDateRange_(outputSheet) {
  const startRaw = outputSheet.getRange(OSR_CFG.START_CELL).getValue();
  const endRaw = outputSheet.getRange(OSR_CFG.END_CELL).getValue();

  let startDate = asDate_(startRaw);
  let endDate = asDate_(endRaw);

  if (startDate && endDate) {
    return { startDate: startOfDay_(startDate), endDate: endOfDay_(endDate) };
  }

  // fallback parse (shouldn't happen anymore)
  const b2 = String(startRaw || "").trim();
  const m = b2.match(/(\d{4}-\d{2}-\d{2})\s*(?:to|\-)\s*(\d{4}-\d{2}-\d{2})/i);
  if (m) {
    startDate = asDate_(m[1]);
    endDate = asDate_(m[2]);
    if (startDate && endDate) {
      return { startDate: startOfDay_(startDate), endDate: endOfDay_(endDate) };
    }
  }

  return { startDate: null, endDate: null };
}

// SIDEBAR DATE RANGE SETTER (presets)
function setOrdersSummaryDateRangeFromSidebar(preset, startStr, endStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OSR_CFG.OUTPUT_SHEET) || ss.insertSheet(OSR_CFG.OUTPUT_SHEET);

  const today = new Date();
  const t0 = startOfDay_(today);

  let start, end;

  const parseYMD = (s) => {
    if (!s) return null;
    const d = new Date(s + "T00:00:00");
    return isNaN(d.getTime()) ? null : d;
  };

  switch (preset) {
    case "today":
      start = t0; end = endOfDay_(today); break;
    case "yesterday": {
      const y = new Date(t0); y.setDate(y.getDate() - 1);
      start = startOfDay_(y); end = endOfDay_(y);
      break;
    }
    case "last7": {
      const s = new Date(t0); s.setDate(s.getDate() - 6);
      start = startOfDay_(s); end = endOfDay_(today);
      break;
    }
    case "last30": {
      const s = new Date(t0); s.setDate(s.getDate() - 29);
      start = startOfDay_(s); end = endOfDay_(today);
      break;
    }
    case "monthToDate": {
      const s = new Date(today.getFullYear(), today.getMonth(), 1);
      start = startOfDay_(s); end = endOfDay_(today);
      break;
    }
    case "yearToDate": {
      const s = new Date(today.getFullYear(), 0, 1);
      start = startOfDay_(s); end = endOfDay_(today);
      break;
    }
    case "custom":
    default:
      start = parseYMD(startStr);
      end = parseYMD(endStr);
      if (start) start = startOfDay_(start);
      if (end) end = endOfDay_(end);
      break;
  }

  if (!start || !end) throw new Error("Invalid date range. Pick a preset or provide start + end dates.");

  // Write REAL dates
  sh.getRange(OSR_CFG.START_CELL).setValue(start).setNumberFormat("yyyy-mm-dd");
  sh.getRange(OSR_CFG.END_CELL).setValue(end).setNumberFormat("yyyy-mm-dd");

  ss.toast(`Date range set: ${formatDate_(start)} to ${formatDate_(end)}`, "Date Range", 5);
  return `Date range set to ${formatDate_(start)} to ${formatDate_(end)}`;
}

/**
 * ✅ Normalizes platform names so Squarespace lines don't get dropped/mismatched.
 * This is the main reason Top Products often shows only Shopify.
 */
function normalizePlatform_(p) {
  const s = String(p || "").trim().toLowerCase();
  if (!s) return "";
  if (s.includes("shopify")) return "Shopify";
  if (s.includes("squarespace") || s.includes("square space") || s.includes("sqsp")) return "Squarespace";
  // fallback: Title Case-ish
  return String(p || "").trim();
}
