/**
 * 03_SheetFormatter.gs
 * Sheet Layout, Currency Formatting, Status Chips, Today's Highlight & Date Filter
 */

// ==========================================
// 1. SHEET ACCESS (BACKGROUND TRIGGER SAFE)
// ==========================================

/**
 * Robust sheet resolver that works 100% reliably in both interactive and
 * headless background triggers (when there is no active UI session).
 */
function getOrCreateOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('eBay Orders') || ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.getActiveSheet() || ss.getSheets()[0];
  }
  try {
    ensureOrdersSheetHeaders(sheet);
  } catch (e) {
    console.warn('ensureOrdersSheetHeaders notice:', e);
  }
  return sheet;
}

/**
 * Initializes clean, professional headers and KPI cards if starting on a new sheet
 */
function ensureOrdersSheetHeaders(sheet) {
  if (!sheet) return;

  // If Row 5 already has headers, preserve existing sheet!
  const existingHeader = sheet.getRange(5, 2).getValue();
  if (existingHeader && existingHeader.toString().trim() !== '') {
    return;
  }

  // Ensure sheet has at least 100 rows
  const currentRows = sheet.getMaxRows();
  if (currentRows < 100) {
    sheet.insertRowsAfter(currentRows, 100 - currentRows);
  }

  // 1. Configure Filter Search Box at J1:K1
  sheet.getRange('J1').setValue('🔍 Filter Date (K1):')
    .setFontWeight('bold')
    .setFontColor('#374151')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  const k1 = sheet.getRange('K1');
  k1.setNumberFormat('yyyy-mm-dd')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FEF3C7')
    .setFontWeight('bold');

  // 2. Configure KPI Cards at Row 2-3
  sheet.getRange('B2').setValue('TOTAL ORDERS').setFontWeight('bold').setFontColor('#6B7280').setFontSize(9);
  sheet.getRange('B3').setFormula('=COUNTA(B6:B)').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  sheet.getRange('D2').setValue('TOTAL REVENUE').setFontWeight('bold').setFontColor('#6B7280').setFontSize(9);
  sheet.getRange('D3').setFormula('=SUM(J6:J)').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  sheet.getRange('G2').setValue('TOTAL DEDUCTIONS').setFontWeight('bold').setFontColor('#6B7280').setFontSize(9);
  sheet.getRange('G3').setFormula('=SUM(I6:I)').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  sheet.getRange('J2').setValue('NET EARNINGS').setFontWeight('bold').setFontColor('#6B7280').setFontSize(9);
  sheet.getRange('J3').setFormula('=SUM(K6:K)').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // 3. Row 5 Table Headers (14 columns A to N)
  const headers = [
    'Sync Timestamp',
    'Order ID',
    'Order Date',
    'Buyer Name',
    'Buyer Username',
    'Contact Phone',
    'Shipping Address',
    'Item Cost',
    'Deductions',
    'Total Amount',
    'Net Earnings',
    'Currency',
    'eBay Link',
    'Order Status'
  ];

  const headerRange = sheet.getRange(5, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(5, 34);

  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 160); // Order ID
  sheet.setColumnWidth(3, 110); // Order Date
  sheet.setColumnWidth(4, 150); // Buyer Name
  sheet.setColumnWidth(5, 130); // Username
  sheet.setColumnWidth(6, 130); // Phone
  sheet.setColumnWidth(7, 240); // Address
  sheet.setColumnWidth(8, 110); // Cost
  sheet.setColumnWidth(9, 110); // Deductions
  sheet.setColumnWidth(10, 110); // Total
  sheet.setColumnWidth(11, 110); // Net
  sheet.setColumnWidth(12, 90);  // Currency
  sheet.setColumnWidth(13, 130); // Link
  sheet.setColumnWidth(14, 120); // Status
}

// ==========================================
// 2. DATA ROW & CURRENCY FORMATTING
// ==========================================

function formatDataRow(sheet, rowNum, currency) {
  try {
    sheet.getRange(rowNum, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Col A: Timestamp
    sheet.getRange(rowNum, 2).setNumberFormat('@');                   // Col B: Order ID (Text)
    sheet.getRange(rowNum, 3).setNumberFormat('yyyy-mm-dd');          // Col C: Order Date
    sheet.getRange(rowNum, 6).setNumberFormat('@');                   // Col F: Phone (Text)

    // Financial formatting
    const curr = currency || (sheet.getRange(rowNum, 12).getValue() || '').toString();
    const currencyFormat = getCurrencyNumberFormat(curr);
    sheet.getRange(rowNum, 8, 1, 4).setNumberFormat(currencyFormat);

    // Alignment
    sheet.getRange(rowNum, 1).setHorizontalAlignment('center');
    sheet.getRange(rowNum, 3).setHorizontalAlignment('center');
    sheet.getRange(rowNum, 12).setHorizontalAlignment('center');
    sheet.getRange(rowNum, 13).setHorizontalAlignment('center');
    sheet.getRange(rowNum, 14).setHorizontalAlignment('center');

    sheet.getRange(rowNum, 1, 1, 14).setVerticalAlignment('middle');
    sheet.getRange(rowNum, 7).setWrap(true);
  } catch (e) {
    console.warn('Formatting warning:', e);
  }
}

function getCurrencyNumberFormat(currencyCode) {
  const c = (currencyCode || '').toUpperCase().trim();
  switch (c) {
    case 'EUR': return '€#,##0.00';
    case 'GBP': return '£#,##0.00';
    case 'USD': return '$#,##0.00';
    case 'AUD': return 'A$#,##0.00';
    case 'CAD': return 'CA$#,##0.00';
    case 'CHF': return 'CHF #,##0.00';
    case 'PLN': return '#,##0.00 zł';
    default:    return '$#,##0.00';
  }
}

function getMarketplaceDefaultCurrency(market) {
  const m = (market || '').toLowerCase().trim();
  const map = {
    'it': 'EUR', 'de': 'EUR', 'fr': 'EUR', 'es': 'EUR',
    'ie': 'EUR', 'at': 'EUR', 'nl': 'EUR', 'be': 'EUR',
    'co.uk': 'GBP', 'com': 'USD', 'com.au': 'AUD',
    'ca': 'CAD', 'ch': 'CHF', 'pl': 'PLN'
  };
  return map[m] || 'USD';
}

function reformatAllCurrencies() {
  const sheet = getOrCreateOrdersSheet();
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - 5;

  const props = PropertiesService.getScriptProperties();
  const market = props.getProperty('EBAY_MARKETPLACE') || 'it';
  const defaultCurrency = getMarketplaceDefaultCurrency(market);

  try {
    const kpiFormat = getCurrencyNumberFormat(defaultCurrency);
    sheet.getRange('D3').setNumberFormat(kpiFormat);
    sheet.getRange('G3').setNumberFormat(kpiFormat);
    sheet.getRange('J3').setNumberFormat(kpiFormat);
  } catch (e) {}

  if (numRows <= 0) {
    SpreadsheetApp.getUi().alert('Notice', 'No order rows found starting at Row 6.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const currencyValues = sheet.getRange(6, 12, numRows, 1).getValues();
  let updatedCount = 0;

  for (let i = 0; i < numRows; i++) {
    const rowNum = i + 6;
    const rawCurrency = (currencyValues[i][0] || '').toString().trim();
    const curr = rawCurrency || defaultCurrency;
    const formatString = getCurrencyNumberFormat(curr);

    sheet.getRange(rowNum, 8, 1, 4).setNumberFormat(formatString);
    updatedCount++;
  }

  SpreadsheetApp.getUi().alert(
    '✓ Currency Update Complete',
    `Successfully updated financial formatting for ${updatedCount} orders to their native currency format!\n\nColumns H, I, J, K and KPI cards now display correct symbols.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================
// 3. STATUS VALIDATION, TODAY HIGHLIGHT & SEARCH BOX
// ==========================================

function setupStatusValidationAndHighlight() {
  const sheet = getOrCreateOrdersSheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Column N Dropdown Validation (Pending, Shipped, Delivered, Cancelled)
  const statusValues = ['Pending', 'Shipped', 'Delivered', 'Cancelled'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusValues, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(6, 14, 995, 1).setDataValidation(rule);

  // 2. Status Color Pills (Column N)
  const statusRange = [sheet.getRange('N6:N1000')];

  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setBold(true)
    .setRanges(statusRange)
    .build();

  const shippedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Shipped')
    .setBackground('#DBEAFE')
    .setFontColor('#1E40AF')
    .setBold(true)
    .setRanges(statusRange)
    .build();

  const deliveredRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Delivered')
    .setBackground('#DCFCE7')
    .setFontColor('#166534')
    .setBold(true)
    .setRanges(statusRange)
    .build();

  const cancelledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancelled')
    .setBackground('#FEE2E2')
    .setFontColor('#991B1B')
    .setBold(true)
    .setRanges(statusRange)
    .build();

  // 3. Today's Highlight Rule (Columns A to M)
  const todayFormula = '=AND($C6<>"", OR(IFERROR(INT($C6)=TODAY(), FALSE), IFERROR(LEFT(TRIM($C6),10)=TEXT(TODAY(),"yyyy-mm-dd"), FALSE)))';
  const todayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(todayFormula)
    .setBackground('#E8F8F0')
    .setRanges([sheet.getRange('A6:M1000')])
    .build();

  sheet.setConditionalFormatRules([pendingRule, shippedRule, deliveredRule, cancelledRule, todayRule]);

  // 4. Clean Existing Rows
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - 5;
  if (numRows > 0) {
    sheet.getRange(6, 3, numRows, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(6, 14, numRows, 1).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // 5. Configure Filter Date Search Box at Cells J1:K1
  try {
    sheet.getRange('J1').setValue('🔍 Filter Date (K1):')
      .setFontWeight('bold')
      .setFontColor('#374151')
      .setHorizontalAlignment('right')
      .setVerticalAlignment('middle');

    const k1 = sheet.getRange('K1');
    k1.setNumberFormat('yyyy-mm-dd')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#FEF3C7')
      .setFontWeight('bold');
  } catch (e) {}

  ui.alert(
    '✓ Setup Complete',
    '1. Status dropdown configured (Pending, Shipped, Delivered, Cancelled).\n' +
    '2. Status colored chips activated (Pending=Amber, Shipped=Blue, Delivered=Green, Cancelled=Red).\n' +
    '3. Today\'s orders highlight rule activated.\n' +
    '4. Column C (Order Date) converted to date-only (yyyy-mm-dd).\n' +
    '5. Date Filter Box activated at Cell K1 (Type date to filter, clear to show all).',
    ui.ButtonSet.OK
  );
}

// ==========================================
// 4. CELL K1 AUTO DATE FILTER
// ==========================================

function applyOrderDateFilter(e) {
  const ss = (e && e.source) ? e.source : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = (e && e.range) ? e.range.getSheet() : ss.getActiveSheet();

  const sheetName = sheet.getName();
  if (sheetName !== 'eBay Orders' && sheetName !== 'Orders' && !sheetName.toLowerCase().includes('order')) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return;

  const range = (e && e.range) ? e.range : sheet.getRange(1, 11);
  const filterVal = range.getValue();
  const displayVal = range.getDisplayValue().trim();

  // If Cell K1 is empty or cleared, show all orders
  if (!filterVal || displayVal === '') {
    const filter = sheet.getFilter();
    if (filter) {
      filter.removeColumnFilterCriteria(3);
    }
    return;
  }

  // Ensure filter covers table headers (Row 5) down to last data row
  let filter = sheet.getFilter();
  if (filter) {
    const fRange = filter.getRange();
    if (fRange.getRow() !== 5 || fRange.getLastRow() < lastRow) {
      filter.remove();
      filter = null;
    }
  }

  if (!filter) {
    const tableRange = sheet.getRange(5, 1, lastRow - 4, 14);
    filter = tableRange.createFilter();
  }

  // Smart Date Parsing without timezone shift:
  // 1. If display value is already YYYY-MM-DD (e.g. 2026-10-01), use it directly! (No Date object conversion!)
  // 2. If a Date object, format using the Spreadsheet's native timezone (ss.getSpreadsheetTimeZone).
  let targetDateStr = '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(displayVal)) {
    targetDateStr = displayVal;
  } else if (filterVal instanceof Date) {
    const ssTz = ss.getSpreadsheetTimeZone() || 'GMT';
    targetDateStr = Utilities.formatDate(filterVal, ssTz, 'yyyy-MM-dd');
  } else {
    targetDateStr = displayVal;
  }

  if (!targetDateStr) {
    filter.removeColumnFilterCriteria(3);
    return;
  }

  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains(targetDateStr)
    .build();

  filter.setColumnFilterCriteria(3, criteria);
}

function clearOrderDateFilter() {
  const sheet = getOrCreateOrdersSheet();
  sheet.getRange(1, 11).clearContent();
  const filter = sheet.getFilter();
  if (filter) {
    filter.removeColumnFilterCriteria(3);
  }
  SpreadsheetApp.getUi().alert('✓ Filter Cleared', 'All orders are now visible.', SpreadsheetApp.getUi().ButtonSet.OK);
}

