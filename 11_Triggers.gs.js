function createRefundRefreshTriggerEvery4Hours() {
  // Remove any existing triggers for refreshShopifyAdjustments if you want to avoid duplicates:
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getHandlerFunction() === 'refreshShopifyAdjustments') {
      // Uncomment to delete existing: ScriptApp.deleteTrigger(t);
      // For safety we don't auto-delete here.
    }
  }
  ScriptApp.newTrigger('refreshShopifyAdjustments').timeBased().everyHours(4).create();
  return "Created trigger to run refreshShopifyAdjustments every 4 hours.";
}