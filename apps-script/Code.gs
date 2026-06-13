/**
 * Delivery Tool — Main Entry Point
 * 
 * This file handles menu creation, installation hooks,
 * and user license validation.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Delivery Tool')
    .addItem('Send Orders', 'sendPendingOrders')
    .addItem('Update Statuses', 'refreshAllStatuses')
    .addSeparator()
    .addItem('Stats Dashboard', 'showStatsSidebar')
    .addItem('Blacklist Manager', 'showBlacklistSidebar')
    .addSeparator()
    .addItem('Settings', 'showSettingsSidebar')
    .addItem('Help', 'showHelpSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetStructure(sheet);
}

function ensureSheetStructure(sheet) {
  var configSheet = sheet.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!configSheet) {
    configSheet = sheet.insertSheet(CONFIG.CONFIG_SHEET);
    configSheet.getRange('A1').setValue('Plan');
    configSheet.getRange('A2').setValue('standard');
    configSheet.getRange('B1').setValue('Provider API Keys (JSON)');
    configSheet.getRange('B2').setValue('{}');
  }

  var ordersSheet = sheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!ordersSheet) {
    ordersSheet = sheet.insertSheet(CONFIG.SHEET_NAME);
    var headers = [
      'Order ID', 'Customer', 'Phone', 'Address', 'Wilaya', 'Commune',
      'Product', 'Quantity', 'Price (DA)', 'Delivery (DA)', 'Provider',
      'Status', 'Tracking', 'Notes', 'Date'
    ];
    ordersSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    ordersSheet.setFrozenRows(1);
  }
}

function getUserPlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!configSheet) return 'standard';
  var plan = configSheet.getRange('A2').getValue().toString().toLowerCase().trim();
  return CONFIG.PLANS[plan] ? plan : 'standard';
}
