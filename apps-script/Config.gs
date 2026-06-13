/**
 * Delivery Tool — Configuration
 * 
 * All settings in one place. Edit this file to:
 * - Change which columns your sheet uses
 * - Add or remove delivery providers
 * - Set API keys (or reference a config sheet)
 */

var CONFIG = {
  // ── Sheet Setup ──
  SHEET_NAME: 'Orders',
  CONFIG_SHEET: 'Config',
  HEADER_ROW: 1,

  // ── Column Mapping (0-indexed) ──
  COLUMNS: {
    ORDER_ID:     0,
    CUSTOMER:     1,
    PHONE:        2,
    ADDRESS:      3,
    WILAYA:       4,
    COMMUNE:      5,
    PRODUCT:      6,
    QUANTITY:     7,
    PRICE:        8,
    DELIVERY:     9,
    PROVIDER:     10,
    STATUS:       11,
    TRACKING:     12,
    NOTES:        13,
    DATE:         14
  },

  // ── Status Values ──
  STATUS: {
    PENDING:    'Pending',
    SENT:       'Sent',
    DELIVERED:  'Delivered',
    FAILED:     'Failed',
    CANCELLED:  'Cancelled'
  },

  // ── Available Providers ──
  PROVIDERS: [
    { id: 'yalidine',   name: 'Yalidine',    apiVersion: 'v1' },
    { id: 'zr',         name: 'ZR Express',  apiVersion: 'v2' },
    { id: 'dhd',        name: 'DHD',         apiVersion: 'v1' },
    { id: 'allo',       name: 'Allo',        apiVersion: 'v1' },
    { id: '48hr',       name: '48HR',        apiVersion: 'v1' },
    { id: 'worldex',    name: 'WorldExpress',apiVersion: 'v1' },
    { id: 'anderson',   name: 'Anderson',    apiVersion: 'v1' },
    { id: 'guepex',     name: 'Guepex',      apiVersion: 'v1' },
    { id: 'samex',      name: 'Samex',       apiVersion: 'v1' },
    { id: 'ovred',      name: 'Ovred',       apiVersion: 'v1' },
    { id: 'navex',      name: 'Navex',       apiVersion: 'v1' },
    { id: 'elogistia',  name: 'Elogistia',   apiVersion: 'v1' },
    { id: 'colireli',   name: 'Colireli',    apiVersion: 'v1' },
    { id: 'assil',      name: 'Assil',       apiVersion: 'v1' },
    { id: 'areex',      name: 'Areex',       apiVersion: 'v1' },
    { id: 'ecom',       name: 'Ecom',        apiVersion: 'v1' },
    { id: 'noest',      name: 'Noest',       apiVersion: 'v1' },
    { id: 'salva',      name: 'Salva',       apiVersion: 'v1' },
    { id: 'expedia',    name: 'Expedia',     apiVersion: 'v1' }
  ],

  // ── API Base URLs ──
  API: {
    yalidine:  'https://api.yalidine.net',
    zr:        'https://api.zrexpress.dz',
    dhd:       'https://api.dhd.dz'
    // extend with real endpoints as you onboard each provider
  },

  // ── Plan Limits ──
  PLANS: {
    standard: {
      maxProviders: 1,
      maxConfirmatrices: 0,
      customPrices: false,
      blacklist: false
    },
    advanced: {
      maxProviders: 5,
      maxConfirmatrices: 5,
      customPrices: true,
      blacklist: true
    }
  }
};
