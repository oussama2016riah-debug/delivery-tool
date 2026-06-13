/**
 * Delivery Tool — UI Components
 * 
 * Sidebar HTML and helper functions
 * for the interactive interface.
 */

function getStatsJSON() {
  return JSON.stringify(generateStats());
}

function getBlacklist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Blacklist');
  if (!sheet) return '[]';

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) list.push({ phone: data[i][0], reason: data[i][1] || '' });
  }
  return JSON.stringify(list);
}

function showSettingsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SettingsSidebar')
    .setTitle('⚙️ Settings')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showHelpSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('HelpSidebar')
    .setTitle('❓ Help')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showBlacklistSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('BlacklistSidebar')
    .setTitle('🚫 Blacklist')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Settings IO ──
function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!sheet) return { keys: '{}', plan: 'standard' };

  return {
    keys: sheet.getRange('B2').getValue() || '{}',
    plan: sheet.getRange('A2').getValue() || 'standard'
  };
}

function saveSettings(keys, plan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CONFIG_SHEET);
    sheet.getRange('A1').setValue('Plan');
    sheet.getRange('B1').setValue('Provider API Keys (JSON)');
  }

  // validate JSON
  try { JSON.parse(keys); } catch (e) { throw new Error('Invalid JSON format'); }

  sheet.getRange('A2').setValue(plan);
  sheet.getRange('B2').setValue(keys);
}

// ── Blacklist IO ──
function addToBlacklist(phone, reason) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Blacklist');
  if (!sheet) {
    sheet = ss.insertSheet('Blacklist');
    sheet.getRange('A1').setValue('Phone');
    sheet.getRange('B1').setValue('Reason');
  }

  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1).setValue(phone);
  sheet.getRange(nextRow, 2).setValue(reason || '');
}

function removeFromBlacklist(index) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Blacklist');
  if (!sheet) return;

  var row = index + 2; // header + 0-indexed
  sheet.deleteRow(row);
}
