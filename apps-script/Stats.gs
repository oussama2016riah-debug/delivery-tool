/**
 * Delivery Tool — Statistics
 * 
 * Generates per-user and per-confirmatrice
 * performance reports.
 */

function refreshAllStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[CONFIG.COLUMNS.STATUS];
    var tracking = row[CONFIG.COLUMNS.TRACKING];
    var provider = row[CONFIG.COLUMNS.PROVIDER];

    if (status !== CONFIG.STATUS.SENT || !tracking || !provider) continue;

    var newStatus = checkProviderStatus(provider.toString().toLowerCase().trim(), tracking.toString().trim());
    if (newStatus !== 'Unknown' && newStatus !== status) {
      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue(newStatus);
      updated++;
    }
  }
}

function generateStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var stats = {
    total: 0,
    pending: 0,
    sent: 0,
    delivered: 0,
    failed: 0,
    cancelled: 0,
    totalRevenue: 0,
    totalDeliveryFees: 0,
    byProvider: {},
    byWilaya: {}
  };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = (row[CONFIG.COLUMNS.STATUS] || '').toString();
    var provider = (row[CONFIG.COLUMNS.PROVIDER] || '').toString();
    var wilaya = (row[CONFIG.COLUMNS.WILAYA] || '').toString();
    var price = parseFloat(row[CONFIG.COLUMNS.PRICE]) || 0;
    var delivery = parseFloat(row[CONFIG.COLUMNS.DELIVERY]) || 0;

    stats.total++;

    switch (status) {
      case CONFIG.STATUS.PENDING:   stats.pending++;   break;
      case CONFIG.STATUS.SENT:      stats.sent++;      break;
      case CONFIG.STATUS.DELIVERED: stats.delivered++; stats.totalRevenue += price; break;
      case CONFIG.STATUS.FAILED:    stats.failed++;    break;
      case CONFIG.STATUS.CANCELLED: stats.cancelled++; break;
    }

    stats.totalDeliveryFees += delivery;

    if (provider) {
      if (!stats.byProvider[provider]) stats.byProvider[provider] = { total: 0, delivered: 0, revenue: 0 };
      stats.byProvider[provider].total++;
      if (status === CONFIG.STATUS.DELIVERED) {
        stats.byProvider[provider].delivered++;
        stats.byProvider[provider].revenue += price;
      }
    }

    if (wilaya) {
      if (!stats.byWilaya[wilaya]) stats.byWilaya[wilaya] = { total: 0, delivered: 0 };
      stats.byWilaya[wilaya].total++;
      if (status === CONFIG.STATUS.DELIVERED) stats.byWilaya[wilaya].delivered++;
    }
  }

  // calculate rates
  stats.deliveryRate = stats.total > 0 ? ((stats.delivered / stats.total) * 100).toFixed(1) : 0;

  return stats;
}

function showStatsSidebar() {
  var stats = generateStats();
  if (!stats) {
    SpreadsheetApp.getUi().alert('No data found.');
    return;
  }

  var html = HtmlService.createHtmlOutputFromFile('StatsSidebar')
    .setTitle('📊 Dashboard')
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}
