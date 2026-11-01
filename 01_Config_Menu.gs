/**
 * 01_Config_Menu.gs
 * System Configuration, Top Menu, Background Trigger Installers & Edit Listeners
 */

// ==========================================
// 1. SYSTEM CONFIGURATION
// ==========================================
const EBAY_CONFIG = {
  // 🔗 MASTER AUTH RELAY URL (Connects to Muhammad Haseeb's Private License Server)
  AUTH_RELAY_URL: 'LINK OF YOUR MASTERSHEET WHERE YOU ADD USERS',

  // Public eBay Orders REST API Endpoint
  API_ENDPOINT: 'https://api.ebay.com/sell/fulfillment/v1/order'
};

// ==========================================
// 2. SYNC CONDITION / MODE PREFERENCE
// ==========================================

/**
 * Saves user preference for background sync condition (NEW_AND_UPDATE or SYNC_ALL)
 */
function saveSyncModePreference(mode) {
  const cleanMode = (mode === 'SYNC_ALL') ? 'SYNC_ALL' : 'NEW_AND_UPDATE';
  PropertiesService.getScriptProperties().setProperty('SYNC_MODE', cleanMode);
  console.log('Saved sync condition mode preference:', cleanMode);
  return cleanMode;
}

/**
 * Retrieves the saved sync condition preference (defaults to NEW_AND_UPDATE)
 */
function getSyncModePreference() {
  try {
    return PropertiesService.getScriptProperties().getProperty('SYNC_MODE') || 'NEW_AND_UPDATE';
  } catch (e) {
    return 'NEW_AND_UPDATE';
  }
}

// ==========================================
// 3. UI MENUS
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  let modeTag = 'New & Update';
  try {
    const currentMode = getSyncModePreference();
    modeTag = currentMode === 'SYNC_ALL' ? 'Sync All' : 'New & Update';
  } catch (e) {}

  ui.createMenu('⚡ eBay Sync')
    .addItem('▶️ Sync Orders (Last 15 Days)', 'syncOrders15Days')
    .addItem('📅 Custom Date Range Sync...', 'openCustomDateRangeDialog')
    .addItem(`⚙️ Set Sync Condition [${modeTag}]...`, 'openSyncSettingsDialog')
    .addSeparator()
    .addItem('1️⃣ Run Condition 1: New Orders & Update Status', 'syncNewOrdersAndUpdateStatus')
    .addItem('2️⃣ Run Condition 2: Sync All (New + Previous)', 'syncAllOrdersFullHistory')
    .addSeparator()
    .addItem('🔃 Sort All Orders (Latest on Top)', 'manualSortOrders')
    .addItem('🔍 Clear Date Filter (Show All)', 'clearOrderDateFilter')
    .addItem('💱 Update All Currency Formats', 'reformatAllCurrencies')
    .addItem('🎨 Setup Status Chips & Today Highlight', 'setupStatusValidationAndHighlight')
    .addSeparator()
    .addItem('🛡️ Subscription & License Status', 'showSubscriptionStatusModal')
    .addSeparator()
    .addItem('⏰ Enable 15-Minute Auto Sync (Fast)', 'installTrigger15Min')
    .addItem('⏰ Enable 1-Hour Auto Sync (Standard)', 'installHourlyTrigger')
    .addItem('🛑 Disable Background Sync', 'removeHourlyTrigger')
    .addSeparator()
    .addItem('🔑 1. Authorize My eBay Store (Step 1)', 'openAuthorizationUrl')
    .addItem('🔄 2. Paste Redirect Code (Step 2)', 'promptExchangeAuthCode')
    .addToUi();

  ui.createMenu('📖 How to Use')
    .addItem('🚀 Quick Start Guide (3 Simple Steps)', 'showUserGuideModal')
    .addItem('❓ FAQs & Troubleshooting', 'showFaqModal')
    .addToUi();
}

// ==========================================
// 3. BACKGROUND TRIGGERS (15-MIN & 1-HOUR)
// ==========================================

/**
 * Installs background trigger to run every 15 minutes automatically 24/7
 */
function installTrigger15Min() {
  const ui = SpreadsheetApp.getUi();
  removeHourlyTriggerSilent();

  ScriptApp.newTrigger('syncEbayOrdersHourly')
    .timeBased()
    .everyMinutes(15)
    .create();

  ui.alert(
    '✓ Active: 15-Minute Auto Sync',
    'Google Cloud will now automatically check for new or updated eBay orders every 15 minutes 24/7 in the background.\n\nEven when this spreadsheet is closed, orders will sync!',
    ui.ButtonSet.OK
  );
}

/**
 * Installs background trigger to run every 1 hour automatically 24/7
 */
function installHourlyTrigger() {
  const ui = SpreadsheetApp.getUi();
  removeHourlyTriggerSilent();

  ScriptApp.newTrigger('syncEbayOrdersHourly')
    .timeBased()
    .everyHours(1)
    .create();

  ui.alert(
    '✓ Active: 1-Hour Auto Sync',
    'Google Cloud will now automatically check for new or updated eBay orders every 1 hour 24/7 in the background.',
    ui.ButtonSet.OK
  );
}

/**
 * Disables background auto-sync triggers
 */
function removeHourlyTrigger() {
  const ui = SpreadsheetApp.getUi();
  removeHourlyTriggerSilent();
  ui.alert('🛑 Disabled', 'Background triggers have been removed. Auto-sync is paused.', ui.ButtonSet.OK);
}

/**
 * Cleans up any existing triggers for syncEbayOrdersHourly silently
 */
function removeHourlyTriggerSilent() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'syncEbayOrdersHourly') {
      ScriptApp.deleteTrigger(t);
    }
  }
}

// ==========================================
// 4. ONEDIT LISTENER
// ==========================================
function onEdit(e) {
  if (!e || !e.range) return;

  try {
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Trigger date filter when Cell K1 (Row 1, Column 11) is edited
    if (row === 1 && col === 11) {
      applyOrderDateFilter(e);
    }
  } catch (err) {
    console.warn('onEdit warning:', err);
  }
}

