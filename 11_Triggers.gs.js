// REMOVED: createRefundRefreshTriggerEvery4Hours() function
// This old trigger used refreshShopifyAdjustments() which has been deprecated
// Refund imports are now handled by:
// - automatedImportAndUpdate() for scheduled imports (uses importShopifyRefunds())
// - Manual buttons in sidebar: "Import Historical Refunds" and "Check Refunds Only"