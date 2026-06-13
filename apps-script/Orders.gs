/**
 * Delivery Tool — Order Processing
 * 
 * Reads orders from the sheet and dispatches them
 * to the selected delivery provider's API.
 */

function sendPendingOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No "' + CONFIG.SHEET_NAME + '" sheet found.');
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No orders to send.');
    return;
  }

  var plan = getUserPlan();
  var planLimits = CONFIG.PLANS[plan];
  var sent = 0, errors = [];

  // Track which providers we've used this run
  var providersUsed = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[CONFIG.COLUMNS.STATUS];

    if (status !== '' && status !== CONFIG.STATUS.PENDING) continue;

    var providerId = row[CONFIG.COLUMNS.PROVIDER];
    if (!providerId) {
      errors.push('Row ' + (i + 1) + ': no provider specified');
      continue;
    }

    providerId = providerId.toString().toLowerCase().trim();

    // Plan check: standard allows 1 provider total
    if (plan === 'standard' && !providersUsed[providerId]) {
      if (Object.keys(providersUsed).length >= 1) {
        errors.push('Row ' + (i + 1) + ': Standard plan — only 1 provider allowed');
        continue;
      }
    }

    providersUsed[providerId] = true;

    var order = buildOrderObject(row);

    try {
      var result = dispatchOrder(providerId, order);
      if (result.success) {
        sheet.getRange(i + 1, CONFIG.COLUMNS.TRACKING + 1).setValue(result.tracking);
        sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue(CONFIG.STATUS.SENT);
        sent++;
      } else {
        sheet.getRange(i + 1, CONFIG.COLUMNS.NOTES + 1).setValue('API Error: ' + result.error);
        sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue(CONFIG.STATUS.FAILED);
        errors.push('Row ' + (i + 1) + ': ' + result.error);
      }
    } catch (e) {
      sheet.getRange(i + 1, CONFIG.COLUMNS.NOTES + 1).setValue('Exception: ' + e.message);
      sheet.getRange(i + 1, CONFIG.COLUMNS.STATUS + 1).setValue(CONFIG.STATUS.FAILED);
      errors.push('Row ' + (i + 1) + ': ' + e.message);
    }
  }

  var msg = sent + ' order(s) sent successfully.';
  if (errors.length) msg += '\n\nErrors:\n' + errors.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

function buildOrderObject(row) {
  return {
    orderId:    row[CONFIG.COLUMNS.ORDER_ID],
    customer:   row[CONFIG.COLUMNS.CUSTOMER],
    phone:      row[CONFIG.COLUMNS.PHONE],
    address:    row[CONFIG.COLUMNS.ADDRESS],
    wilaya:     row[CONFIG.COLUMNS.WILAYA],
    commune:    row[CONFIG.COLUMNS.COMMUNE],
    product:    row[CONFIG.COLUMNS.PRODUCT],
    quantity:   row[CONFIG.COLUMNS.QUANTITY],
    price:      row[CONFIG.COLUMNS.PRICE],
    delivery:   row[CONFIG.COLUMNS.DELIVERY],
    provider:   row[CONFIG.COLUMNS.PROVIDER]
  };
}

function dispatchOrder(providerId, order) {
  switch (providerId) {
    case 'yalidine':  return sendYalidine(order);
    case 'zr':        return sendZR(order);
    case 'dhd':       return sendDHD(order);
    default:          return sendGeneric(providerId, order);
  }
}
