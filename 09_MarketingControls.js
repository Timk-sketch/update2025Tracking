// =====================================================
// 09_MarketingControls.gs — Marketing Controls (Sidebar-only inputs)
// =====================================================

function ensureMarketingControlsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MARKETING_CONTROLS_SHEET);
  if (!sh) sh = ss.insertSheet(MARKETING_CONTROLS_SHEET);

  if (sh.getLastRow() < 1) {
    const rows = [
      ["Marketing Controls", ""],
      ["Marketing Spend ($) — total paid media + agency + creative for the selected period", 0],
      ["Contribution Margin (%) — margin AFTER COGS/fulfillment/fees/refunds/support (your definition of COGS)", 0.60],
      ["Target Spend % of Revenue — typical online businesses often land ~10%–25% depending on stage", 0.15]
    ];
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    sh.getRange(1, 1).setFontWeight("bold");
    sh.getRange(2, 2, 1, 1).setNumberFormat("$#,##0.00");
    sh.getRange(3, 2, 2, 1).setNumberFormat("0.00%");
    sh.autoResizeColumns(1, 2);
  }
  return sh;
}

function getMarketingControls_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(MARKETING_CONTROLS_SHEET) || ensureMarketingControlsSheet_();

  const values = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 2).getValues();
  const map = {};
  for (let i = 0; i < values.length; i++) {
    const k = s_(values[i][0]);
    const v = values[i][1];
    if (!k) continue;
    map[k] = v;
  }

  const spend = n_(map["Marketing Spend ($) — total paid media + agency + creative for the selected period"] ?? 0);
  const marginPct = n_(map["Contribution Margin (%) — margin AFTER COGS/fulfillment/fees/refunds/support (your definition of COGS)"] ?? 0.60);
  const targetSpendPct = n_(map["Target Spend % of Revenue — typical online businesses often land ~10%–25% depending on stage"] ?? 0.15);

  return { marketingSpend: spend, contributionMarginPct: marginPct, targetSpendPct: targetSpendPct };
}

function getMarketingControlsForSidebar() {
  const c = getMarketingControls_();
  return {
    marketingSpend: c.marketingSpend,
    contributionMarginPct: Math.round((c.contributionMarginPct || 0) * 100),
    targetSpendPct: Math.round((c.targetSpendPct || 0) * 100)
  };
}

function setMarketingControlsFromSidebar(marketingSpend, contributionMarginPctWhole, targetSpendPctWhole) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(MARKETING_CONTROLS_SHEET) || ensureMarketingControlsSheet_();

  const marginDec = n_(contributionMarginPctWhole) / 100;
  const targetDec = n_(targetSpendPctWhole) / 100;

  const rows = [
    ["Marketing Controls", ""],
    ["Marketing Spend ($) — total paid media + agency + creative for the selected period", n_(marketingSpend)],
    ["Contribution Margin (%) — margin AFTER COGS/fulfillment/fees/refunds/support (your definition of COGS)", marginDec],
    ["Target Spend % of Revenue — typical online businesses often land ~10%–25% depending on stage", targetDec]
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1).setFontWeight("bold");
  sh.getRange(2, 2, 1, 1).setNumberFormat("$#,##0.00");
  sh.getRange(3, 2, 2, 1).setNumberFormat("0.00%");
  sh.autoResizeColumns(1, 2);

  ss.toast("Marketing controls saved.", "Marketing", 4);
  return "Marketing controls saved.";
}
