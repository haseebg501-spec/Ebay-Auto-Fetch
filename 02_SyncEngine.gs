/**
 * 02_SyncEngine.gs
 * Core eBay Fulfillment Sync Engine, Order Extraction, Status Detection & Sorting
 */

// ==========================================
// 1. SYNC CONDITIONS & TRIGGERS
// ==========================================

/**
 * Background Trigger Entry Point:
 * Automatically called by 15-minute or 1-hour time-based trigger.
 * Executes the user's active sync condition:
 * - Condition 1: Sync New Orders & Update Status
 * - Condition 2: Sync All (New + Previous)
 */
function syncEbayOrdersHourly() {
  const mode = getSyncModePreference();
  return runCoreSync({
    mode: mode,
    daysBack: 15,
    maxPages: 5,
    isBackground: true
  });
}

/**
 * 15-Day Sync Button:
 * Automatically executes the user's active condition preference!
 */
function syncOrders15Days() {
  const mode = getSyncModePreference();
  return runCoreSync({
    mode: mode,
    daysBack: 15,
    maxPages: 5,
    isBackground: false
  });
}

/**
 * Direct Condition 1 Runner: Sync New Orders & Update Status
 */
function syncNewOrdersAndUpdateStatus(isBackground) {
  return runCoreSync({
    mode: 'NEW_AND_UPDATE',
    daysBack: 15,
    maxPages: 5,
    isBackground: isBackground === true
  });
}

/**
 * Direct Condition 2 Runner: Sync All (New + Previous Orders)
 */
function syncAllOrdersFullHistory(isBackground) {
  return runCoreSync({
    mode: 'SYNC_ALL',
    daysBack: 89,
    maxPages: 10,
    isBackground: isBackground === true
  });
}

/**
 * Unified Core Sync Engine
 */
function runCoreSync(options) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(45000)) {
    console.warn('Sync already running. Skipping duplicate call.');
    return;
  }

  const props = PropertiesService.getScriptProperties();

  try {
    const token = getValidAccessToken();

    const now = new Date();
    const daysBack = options.daysBack || 15;
    const pastDaysMs = daysBack * 24 * 60 * 60 * 1000;
    const startTime = new Date(now.getTime() - pastDaysMs).toISOString();
    const endTime = now.toISOString();

    const isAll = options.mode === 'SYNC_ALL';
    const modeTitle = isAll ? 'Sync All (New + Previous Orders)' : 'Sync New Orders & Update Status';

    console.log(`Starting [${modeTitle}] window: ${startTime} to ${endTime}`);

    const endpoint = EBAY_CONFIG.API_ENDPOINT;
    const filterQuery = encodeURIComponent(`creationdate:[${startTime}..${endTime}]`);
    let url = `${endpoint}?filter=${filterQuery}&limit=50`;

    const sheet = getOrCreateOrdersSheet();
    let totalImported = 0;
    let totalUpdated = 0;
    let totalSkipped = 0;

    let hasMore = true;
    let pagesProcessed = 0;
    const maxPages = options.maxPages || 5;
    const allFetchedOrders = [];

    while (hasMore && pagesProcessed < maxPages) {
      pagesProcessed++;
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      });

      const statusCode = response.getResponseCode();
      if (statusCode !== 200) {
        let errDesc = response.getContentText();
        try {
          const errJson = JSON.parse(errDesc);
          if (errJson.errors && errJson.errors[0]?.message) {
            errDesc = errJson.errors[0].message;
          }
        } catch (e) {}
        throw new Error(`eBay API Error (${statusCode}): ${errDesc}`);
      }

      const data = JSON.parse(response.getContentText());
      const orders = data.orders || [];
      allFetchedOrders.push(...orders);

      if (data.next) {
        url = data.next;
      } else {
        hasMore = false;
      }
    }

    console.log(`[${modeTitle}] Fetched ${allFetchedOrders.length} total orders from eBay.`);

    // Sort chronologically ascending so newest sits on Row 6
    allFetchedOrders.sort((a, b) => {
      const dateA = a.creationDate ? new Date(a.creationDate).getTime() : 0;
      const dateB = b.creationDate ? new Date(b.creationDate).getTime() : 0;
      return dateA - dateB;
    });

    for (const order of allFetchedOrders) {
      const result = processAndSaveOrder(sheet, order, options.mode);
      if (result === 'INSERT') totalImported++;
      if (result === 'UPDATE') totalUpdated++;
      if (result === 'SKIPPED') totalSkipped++;
    }

    sortSheetOrdersDescending(sheet);

    // Save status
    props.setProperty('LAST_SYNC_TIME', endTime);
    props.setProperty('LAST_SYNC_STATUS', `Success: ${modeTitle} (${totalImported} new, ${totalUpdated} updated)`);
    props.deleteProperty('LAST_SYNC_ERROR');

    // UI feedback if manually triggered
    if (!options.isBackground) {
      try {
        SpreadsheetApp.getUi().alert(
          `✓ ${modeTitle} Complete`,
          `• New Orders Imported: ${totalImported}\n• Existing Orders Updated: ${totalUpdated}\n• Non-Pending/Shipped Skipped: ${totalSkipped}\n\nAll orders organized with latest on top.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (e) {}
    }

    return {
      status: 'success',
      mode: options.mode,
      imported: totalImported,
      updated: totalUpdated,
      skipped: totalSkipped,
      total: allFetchedOrders.length
    };

  } catch (err) {
    console.error('eBay Sync Error:', err.toString());
    props.setProperty('LAST_SYNC_ERROR', `${new Date().toISOString()}: ${err.toString()}`);
    if (!options.isBackground) {
      try {
        SpreadsheetApp.getUi().alert('❌ Sync Error', err.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {}
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. CUSTOM DATE RANGE SYNC
// ==========================================

/**
 * Custom Date Range Sync: Fetches orders between user-selected dates
 * Filtering strictly for Pending and Shipped orders.
 * Includes automatic guard rails:
 * 1. Caps end date to NOW (eBay API forbids future end dates).
 * 2. Caps start date to 89 days ago (eBay Fulfillment API has a strict 90-day limit).
 */
function syncEbayOrdersCustomRange(startDateStr, endDateStr, chosenMode) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error('Another sync is currently in progress. Please wait a moment and retry.');
  }

  const mode = chosenMode || getSyncModePreference();

  try {
    const token = getValidAccessToken();

    if (!startDateStr || !endDateStr) {
      throw new Error('Please select both Start Date and End Date.');
    }

    let start = new Date(startDateStr + 'T00:00:00.000Z');
    let end = new Date(endDateStr + 'T23:59:59.999Z');

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('Invalid dates selected.');
    }

    if (start.getTime() > end.getTime()) {
      throw new Error('Start Date cannot be after End Date.');
    }

    const now = new Date();

    // Guard Rail 1: End Date cannot be in the future (eBay API throws HTTP 400)
    if (end.getTime() > now.getTime()) {
      end = now;
    }

    // Guard Rail 2: eBay Fulfillment API creationdate cannot exceed 90 days
    const maxPastMs = 89 * 24 * 60 * 60 * 1000;
    const earliestAllowed = new Date(now.getTime() - maxPastMs);
    if (start.getTime() < earliestAllowed.getTime()) {
      start = earliestAllowed;
      console.warn('Start date adjusted to 89 days ago to comply with eBay API 90-day maximum limit.');
    }

    const startIso = start.toISOString();
    const endIso = end.toISOString();
    console.log(`Custom date range sync (${mode}): from ${startIso} to ${endIso}`);

    const endpoint = EBAY_CONFIG.API_ENDPOINT;
    const filterQuery = encodeURIComponent(`creationdate:[${startIso}..${endIso}]`);
    let url = `${endpoint}?filter=${filterQuery}&limit=50`;

    const sheet = getOrCreateOrdersSheet();
    let totalImported = 0;
    let totalUpdated = 0;
    let totalSkipped = 0;

    let hasMore = true;
    let pagesProcessed = 0;
    const allFetchedOrders = [];

    // Up to 10 pages (500 orders) in custom range
    while (hasMore && pagesProcessed < 10) {
      pagesProcessed++;
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      });

      const statusCode = response.getResponseCode();
      if (statusCode !== 200) {
        let errDesc = response.getContentText();
        try {
          const errJson = JSON.parse(errDesc);
          if (errJson.errors && errJson.errors[0]?.message) {
            errDesc = errJson.errors[0].message;
          }
        } catch (e) {}
        throw new Error(`eBay API Error (${statusCode}): ${errDesc}`);
      }

      const data = JSON.parse(response.getContentText());
      const orders = data.orders || [];
      allFetchedOrders.push(...orders);

      if (data.next) {
        url = data.next;
      } else {
        hasMore = false;
      }
    }

    console.log(`Custom sync: fetched ${allFetchedOrders.length} total orders from eBay.`);

    // Sort chronologically ascending
    allFetchedOrders.sort((a, b) => {
      const dateA = a.creationDate ? new Date(a.creationDate).getTime() : 0;
      const dateB = b.creationDate ? new Date(b.creationDate).getTime() : 0;
      return dateA - dateB;
    });

    for (const order of allFetchedOrders) {
      const result = processAndSaveOrder(sheet, order, mode);
      if (result === 'INSERT') totalImported++;
      if (result === 'UPDATE') totalUpdated++;
      if (result === 'SKIPPED') totalSkipped++;
    }

    sortSheetOrdersDescending(sheet);

    return {
      status: 'success',
      imported: totalImported,
      updated: totalUpdated,
      skipped: totalSkipped,
      total: allFetchedOrders.length,
      start: startDateStr,
      end: endDateStr
    };

  } catch (err) {
    console.error('Custom sync error:', err.toString());
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. ORDER ROW PROCESSING
// ==========================================

/**
 * Inserts or updates an individual order row in the sheet
 * Dynamically enforces the user's selected sync condition:
 * - NEW_AND_UPDATE: Updates existing rows, inserts new Pending/Shipped, skips new Delivered/Cancelled
 * - SYNC_ALL: Syncs EVERYTHING (Pending, Shipped, Delivered, Cancelled)
 */
function processAndSaveOrder(sheet, order, syncMode) {
  const mode = syncMode || getSyncModePreference();
  const props = PropertiesService.getScriptProperties();
  const orderId = (order.orderId || 'UNKNOWN').toString().trim();
  const creationDate = order.creationDate ? new Date(order.creationDate) : null;
  const buyerUsername = order.buyer?.username || '';
  
  // Buyer Details
  const buyerDetails = extractBuyerDetails(order);
  const buyerName = buyerDetails.buyerName;
  const contactPhone = buyerDetails.contactPhone;
  const fullAddress = buyerDetails.fullAddress;

  // Financials
  const pricing = order.pricingSummary || {};
  const payment = order.paymentSummary || {};

  const totalAmount = parseFloat(pricing.total?.value || 0.00);
  const cost = parseFloat(pricing.priceSubtotal?.value || 0.00);
  const netEarnings = parseFloat(payment.totalDueSeller?.value || (totalAmount - parseFloat(pricing.fee?.value || 0.00)));
  const deductions = parseFloat((totalAmount - netEarnings).toFixed(2));
  
  // Currency Detection
  const market = props.getProperty('EBAY_MARKETPLACE') || 'it';
  const defaultMarketCurrency = getMarketplaceDefaultCurrency(market);

  const detectedCurrency = (pricing.total && pricing.total.currency) ||
                           (pricing.priceSubtotal && pricing.priceSubtotal.currency) ||
                           (order.paymentSummary && order.paymentSummary.totalDueSeller?.currency) ||
                           (order.lineItems && order.lineItems[0]?.total?.currency) ||
                           defaultMarketCurrency ||
                           'USD';
  const currency = detectedCurrency.toUpperCase().trim();

  // Timestamps
  const now = new Date();
  const timeZone = (sheet.getParent && sheet.getParent()) ? sheet.getParent().getSpreadsheetTimeZone() : (Session.getScriptTimeZone() || 'GMT');
  const syncTimestampStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss');
  const orderDateStr = creationDate 
    ? Utilities.formatDate(creationDate, timeZone, 'yyyy-MM-dd') 
    : Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');

  // Clickable Order Source URL
  const sourceUrlFormula = `=HYPERLINK("https://www.ebay.${market}/sh/ord/details?orderid=${orderId}", "View Order #${orderId}")`;

  // Duplicate Check against Column B starting at Row 6
  const lastRow = sheet.getLastRow();
  let existingRow = null;

  if (lastRow >= 6) {
    const existingIds = sheet.getRange(6, 2, lastRow - 5, 1).getValues();
    for (let i = 0; i < existingIds.length; i++) {
      if (existingIds[i][0] && existingIds[i][0].toString().trim() === orderId) {
        existingRow = i + 6;
        break;
      }
    }
  }

  // Determine order status (Pending, Shipped, Delivered, Cancelled)
  const orderStatus = determineOrderStatus(order);

  // Sync Condition Filtering (Condition 1 vs Condition 2):
  if (mode === 'NEW_AND_UPDATE') {
    // Condition 1: Sync New Orders & Update Status
    // - Existing orders in sheet: ALWAYS update (refreshes Pending -> Shipped, Delivered, financials)
    // - Brand new orders: Only insert if status is Pending or Shipped! (Skip Delivered & Cancelled)
    if (!existingRow && orderStatus !== 'Pending' && orderStatus !== 'Shipped') {
      return 'SKIPPED';
    }
  } else {
    // Condition 2: Sync All (New + Previous)
    // - Imports ALL orders (Pending, Shipped, Delivered, Cancelled)
    // - No status is skipped!
  }

  // Exact 14 columns A to N
  const rowData = [
    syncTimestampStr,  // Col A: Timestamp (Date + Time)
    orderId,           // Col B: Order ID (Text)
    orderDateStr,      // Col C: Order Date (yyyy-MM-dd)
    buyerName,         // Col D: Buyer Name
    buyerUsername,     // Col E: Buyer Username
    contactPhone,      // Col F: Contact Phone (Text)
    fullAddress,       // Col G: Full Address
    cost,              // Col H: Item Cost
    deductions,        // Col I: Deductions / Fees
    totalAmount,       // Col J: Total Amount
    netEarnings,       // Col K: Net Earnings
    currency,          // Col L: Currency
    sourceUrlFormula,  // Col M: Clickable Link
    orderStatus        // Col N: Status Chip (Pending, Shipped)
  ];

  if (existingRow) {
    // Update existing row
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    formatDataRow(sheet, existingRow, currency);
    return 'UPDATE';
  } else {
    // Insert brand new order at Row 6
    if (sheet.getLastRow() < 6) {
      sheet.getRange(6, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.insertRowBefore(6);
      sheet.getRange(6, 1, 1, rowData.length).setValues([rowData]);
    }

    // Copy dropdown validation if present on row below
    try {
      const dropdownRule = sheet.getRange(7, 14).getDataValidation();
      if (dropdownRule) {
        sheet.getRange(6, 14).setDataValidation(dropdownRule);
      }
    } catch (e) {}

    formatDataRow(sheet, 6, currency);
    return 'INSERT';
  }
}

// ==========================================
// 4. ORDER STATUS & DETAILS EXTRACTION
// ==========================================

function determineOrderStatus(order) {
  const fulfillmentStatus = (order.orderFulfillmentStatus || '').toUpperCase();
  const cancelStatus = (order.cancelStatus?.cancelState || '').toUpperCase();

  if (cancelStatus === 'CANCELED' || cancelStatus === 'CANCEL_CLOSED_WITH_REFUND') {
    return 'Cancelled';
  }

  let isDelivered = false;
  const plans = order.fulfillmentStartPlans || [];
  for (const p of plans) {
    const trackings = p.shippingStep?.shipmentTracking || [];
    for (const t of trackings) {
      const delStat = (t.deliveryStatus || t.status || '').toUpperCase();
      if (delStat === 'DELIVERED') {
        isDelivered = true;
        break;
      }
    }
  }

  if (isDelivered) {
    return 'Delivered';
  }

  if (fulfillmentStatus === 'FULFILLED') {
    return 'Shipped';
  }

  return 'Pending';
}

function extractBuyerDetails(order) {
  let buyerName = '';
  let contactPhone = '';
  let fullAddress = '';

  const plans = order.fulfillmentStartPlans || [];
  if (plans.length > 0) {
    const shipStep = plans[0].shippingStep;
    if (shipStep && shipStep.shipTo) {
      const shipTo = shipStep.shipTo;
      buyerName = shipTo.fullName || shipTo.name || buyerName;
      contactPhone = shipTo.primaryPhone?.phoneNumber || shipTo.phone || contactPhone;

      const addr = shipTo.contactAddress;
      if (addr) {
        const parts = [
          addr.addressLine1,
          addr.addressLine2,
          addr.city,
          addr.stateOrProvince,
          addr.postalCode,
          addr.countryCode
        ].filter(p => p && p.toString().trim() !== '');
        fullAddress = parts.join(', ');
      }
    }
  }

  if (!buyerName && order.buyer) {
    buyerName = order.buyer.username || '';
  }

  return {
    buyerName: buyerName.trim(),
    contactPhone: contactPhone.trim(),
    fullAddress: fullAddress.trim()
  };
}

// ==========================================
// 5. SORTING ENGINE
// ==========================================

function sortSheetOrdersDescending(sheet) {
  if (!sheet) sheet = getOrCreateOrdersSheet();
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - 5;

  if (numRows > 1) {
    const dataRange = sheet.getRange(6, 1, numRows, 14);
    dataRange.sort([
      { column: 3, ascending: false }, // Column C: Order Date (latest first)
      { column: 1, ascending: false }  // Column A: Sync Timestamp (latest first)
    ]);
    console.log(`Sorted ${numRows} orders descending with newest on top.`);
  }
}

function manualSortOrders() {
  const sheet = getOrCreateOrdersSheet();
  sortSheetOrdersDescending(sheet);
  SpreadsheetApp.getUi().alert('✓ Completed', 'All orders have been sorted with the latest order on top (Row 6).', SpreadsheetApp.getUi().ButtonSet.OK);
}

