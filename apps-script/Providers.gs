/**
 * Delivery Tool — Provider Integrations
 * 
 * Each delivery company gets its own send function.
 * Follow the same pattern to add new providers.
 * 
 * Every send* function must return:
 *   { success: true, tracking: 'ABC123' }
 *   or
 *   { success: false, error: 'reason' }
 */

function getProviderApiKey(providerId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!configSheet) return null;
  try {
    var raw = configSheet.getRange('B2').getValue();
    var keys = JSON.parse(raw || '{}');
    return keys[providerId] || null;
  } catch (e) {
    return null;
  }
}

// ── Yalidine ──
function sendYalidine(order) {
  var apiKey = getProviderApiKey('yalidine');
  if (!apiKey) return { success: false, error: 'Yalidine API key not configured' };

  var payload = {
    api_key: apiKey,
    order_id: order.orderId,
    client_name: order.customer,
    client_phone: order.phone,
    client_address: order.address,
    wilaya: order.wilaya,
    commune: order.commune,
    product: order.product,
    quantity: order.quantity || 1,
    price: parseFloat(order.price) || 0,
    delivery_price: parseFloat(order.delivery) || 0
  };

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.yalidine + '/orders', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.tracking_number) {
      return { success: true, tracking: result.tracking_number };
    }
    return { success: false, error: result.message || 'Unknown Yalidine error' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── ZR Express ──
function sendZR(order) {
  var apiKey = getProviderApiKey('zr');
  if (!apiKey) return { success: false, error: 'ZR Express API key not configured' };

  var payload = {
    token: apiKey,
    reference: order.orderId,
    full_name: order.customer,
    phone: order.phone,
    address: order.address,
    wilaya: order.wilaya,
    product_name: order.product,
    total_amount: parseFloat(order.price) || 0
  };

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.zr + '/orders/create', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.tracking || result.code) {
      return { success: true, tracking: result.tracking || result.code };
    }
    return { success: false, error: result.error || 'Unknown ZR error' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── DHD ──
function sendDHD(order) {
  var apiKey = getProviderApiKey('dhd');
  if (!apiKey) return { success: false, error: 'DHD API key not configured' };

  var payload = {
    apikey: apiKey,
    commande_id: order.orderId,
    destinataire: order.customer,
    telephone: order.phone,
    adresse: order.address,
    wilaya: order.wilaya,
    produit: order.product,
    montant: parseFloat(order.price) || 0
  };

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.dhd + '/api/orders', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.numero_suivi) {
      return { success: true, tracking: result.numero_suivi };
    }
    return { success: false, error: result.message || 'Unknown DHD error' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Generic Provider Fallback ──
function sendGeneric(providerId, order) {
  var apiKey = getProviderApiKey(providerId);
  if (!apiKey) {
    return { success: false, error: providerId + ' API key not configured' };
  }

  var payload = {
    api_key: apiKey,
    order_id: order.orderId,
    customer_name: order.customer,
    customer_phone: order.phone,
    customer_address: order.address,
    city: order.wilaya + (order.commune ? ', ' + order.commune : ''),
    product_description: order.product,
    product_quantity: order.quantity || 1,
    amount: parseFloat(order.price) || 0,
    delivery_fee: parseFloat(order.delivery) || 0
  };

  // Generic fallback tries to POST to a standard endpoint
  var baseUrl = CONFIG.API[providerId];
  if (!baseUrl) {
    return { success: false, error: 'No API endpoint configured for ' + providerId };
  }

  try {
    var response = UrlFetchApp.fetch(baseUrl + '/orders', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    var tracking = result.tracking || result.tracking_number || result.id || result.order_id || result.code;
    if (tracking) {
      return { success: true, tracking: tracking.toString() };
    }
    return { success: false, error: JSON.stringify(result) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Status Checking ──
function checkProviderStatus(providerId, trackingNumber) {
  switch (providerId) {
    case 'yalidine': return checkYalidineStatus(trackingNumber);
    case 'zr':       return checkZRStatus(trackingNumber);
    case 'dhd':      return checkDHDStatus(trackingNumber);
    default:         return checkGenericStatus(providerId, trackingNumber);
  }
}

function checkYalidineStatus(tracking) {
  var apiKey = getProviderApiKey('yalidine');
  if (!apiKey) return 'Unknown';

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.yalidine + '/orders/' + tracking + '?api_key=' + apiKey, {
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    return result.status || 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

function checkZRStatus(tracking) {
  var apiKey = getProviderApiKey('zr');
  if (!apiKey) return 'Unknown';

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.zr + '/orders/status?token=' + apiKey + '&tracking=' + tracking, {
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    return result.status || 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

function checkDHDStatus(tracking) {
  var apiKey = getProviderApiKey('dhd');
  if (!apiKey) return 'Unknown';

  try {
    var response = UrlFetchApp.fetch(CONFIG.API.dhd + '/api/orders/' + tracking + '?apikey=' + apiKey, {
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    return result.statut || result.status || 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

function checkGenericStatus(providerId, tracking) {
  var apiKey = getProviderApiKey(providerId);
  if (!apiKey) return 'Unknown';

  var baseUrl = CONFIG.API[providerId];
  if (!baseUrl) return 'Unknown';

  try {
    var response = UrlFetchApp.fetch(baseUrl + '/orders/' + tracking + '?api_key=' + apiKey, {
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    return result.status || result.statut || result.state || 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}
