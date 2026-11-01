/**
 * eBay OAuth Relay & Commercial Subscription Server
 * (Private - Stored directly in Muhammad Haseeb's Master Sheet or Google Account)
 * 
 * Functions:
 * 1. Safely stores eBay App ID and private Cert ID.
 * 2. Manages Master Subscribers Database (Gmail, Name, Registration Date, Status, Expiry, Marketplace).
 * 3. Auto-registers clients when they run Step 1 in dd-mmm-yyyy date format.
 * 4. Validates monthly subscription access before issuing eBay access tokens.
 */

// YOUR PRIVATE MASTER CREDENTIALS & CONFIGURATION
const SECRETS = {
  APP_ID: 'EBAY DEVELOPER APP ID',
  CERT_ID: 'YOUR_PRODUCTION_CERT_ID_HERE', // <-- Keep your Production Cert ID here
  RU_NAME: 'EBAY DEVELOPER PRODUCTION RU NAME',
  SCOPES: 'https://api.ebay.com/oauth/api_scope/sell.fulfillment https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly',
  OAUTH_ENDPOINT: 'https://api.ebay.com/identity/v1/oauth2/token',

  // 📊 Master Subscribers Google Sheet ID
  // (Auto-detected if script is pasted directly into your Master Sheet!)
  MASTER_SHEET_ID: 'MASTER SHEET ID'
};

/**
 * Constructs the secure eBay OAuth Authorization URL for a specific marketplace
 * Kept strictly private on your Relay Server so buyers never see your App ID or RuName!
 */
function generateAuthUrl(marketplace) {
  const market = (marketplace || 'it').toLowerCase().trim();
  const scope = encodeURIComponent(SECRETS.SCOPES);
  return `https://auth.ebay.${market}/oauth2/authorize?client_id=${SECRETS.APP_ID.trim()}&response_type=code&redirect_uri=${SECRETS.RU_NAME.trim()}&scope=${scope}&prompt=login`;
}

/**
 * ⚡ STEP 1 SETUP & AUTHORIZATION FUNCTION
 * 
 * RUN THIS ONCE IN APPS SCRIPT (Select from toolbar dropdown & click ▶️ Run)
 * 
 * Why this is necessary:
 * When you run this function in the editor, Google prompts you to "Review Permissions"
 * and click "Allow", which grants the `https://www.googleapis.com/auth/spreadsheets`
 * permission so your Web App never throws the "permission to call openById" error!
 */
function setupAndAuthorizeRelay() {
  const ss = getMasterSpreadsheet();
  const sheet = ss.getSheetByName('Subscribers') || ss.getActiveSheet();
  
  // Set clean tab name if it's default Sheet1
  if (sheet.getName().toLowerCase() === 'sheet1') {
    sheet.setName('Subscribers');
  }

  initMasterSheetHeaders(sheet);

  Logger.log('====================================================');
  Logger.log('✅ AUTHORIZATION SUCCESSFUL!');
  Logger.log('Connected to Master Spreadsheet: ' + ss.getName());
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('Subscribers Sheet: ' + sheet.getName());
  Logger.log('====================================================');
  Logger.log('CRITICAL DEPLOYMENT STEP:');
  Logger.log('1. Click: Deploy > Manage Deployments');
  Logger.log('2. Click the Pencil icon (Edit)');
  Logger.log('3. Under "Version", select "New version"');
  Logger.log('4. Click "Deploy"');
  Logger.log('====================================================');

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Setup & Authorization Successful!',
      'Master Sheet is connected and headers are configured!\n\n' +
      'FINAL STEP TO ACTIVATE:\n' +
      'Click "Deploy" > "Manage Deployments" > Edit (Pencil icon) > select "New version" > click "Deploy".\n\n' +
      'This activates the spreadsheet permissions for your Web App URL!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    // Headless execution
  }
}

/**
 * Safely accesses the Master Spreadsheet:
 * 1. Container-bound: Uses getActiveSpreadsheet() if pasted directly inside the Master Sheet.
 * 2. Standalone fallback: Uses openById(SECRETS.MASTER_SHEET_ID).
 */
function getMasterSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    console.log('Not container-bound, using openById fallback');
  }

  if (SECRETS.MASTER_SHEET_ID && !SECRETS.MASTER_SHEET_ID.includes('YOUR_ID')) {
    return SpreadsheetApp.openById(SECRETS.MASTER_SHEET_ID.trim());
  }

  throw new Error('Master Spreadsheet not accessible. Please ensure EbayAuthRelay is pasted inside your Master Sheet, or set MASTER_SHEET_ID.');
}

/**
 * Initializes beautiful, professional headers in the Master Sheet if empty
 */
function initMasterSheetHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Gmail / User Email',
      'Client Name',
      'Registration Date',
      'Status',
      'Expiry Date',
      'Marketplace',
      'Last Active',
      'Notes'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Header Styling
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0F9D58'); // Google Green
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);
    
    // Set Column text formatting
    sheet.getRange(2, 1, 500, headers.length).setNumberFormat('@');

    // Column widths
    sheet.setColumnWidth(1, 250); // Col A: Gmail
    sheet.setColumnWidth(2, 170); // Col B: Name
    sheet.setColumnWidth(3, 150); // Col C: Reg Date (dd-mmm-yyyy)
    sheet.setColumnWidth(4, 110); // Col D: Status
    sheet.setColumnWidth(5, 150); // Col E: Expiry Date (dd-mmm-yyyy)
    sheet.setColumnWidth(6, 120); // Col F: Marketplace
    sheet.setColumnWidth(7, 150); // Col G: Last Active
    sheet.setColumnWidth(8, 250); // Col H: Notes

    // Dropdown validation on Status (Col D)
    const statusValues = ['Pending', 'Active', 'Suspended', 'Expired', 'Trial'];
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusValues, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 4, 500, 1).setDataValidation(rule);

    // Status Pill Colors (Col D)
    const statusRange = [sheet.getRange('D2:D500')];

    // 1. Pending Rule (Yellow / Amber) - Default for new registrations!
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setBold(true)
      .setRanges(statusRange)
      .build();

    // 2. Active Rule (Green) - Set by Muhammad Haseeb once payment is received
    const activeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Active')
      .setBackground('#DCFCE7')
      .setFontColor('#166534')
      .setBold(true)
      .setRanges(statusRange)
      .build();

    // 3. Suspended Rule (Red)
    const suspendedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Suspended')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setBold(true)
      .setRanges(statusRange)
      .build();

    // 4. Expired Rule (Grey)
    const expiredRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Expired')
      .setBackground('#E5E7EB')
      .setFontColor('#374151')
      .setBold(true)
      .setRanges(statusRange)
      .build();

    // 5. Trial Rule (Blue)
    const trialRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Trial')
      .setBackground('#DBEAFE')
      .setFontColor('#1E40AF')
      .setBold(true)
      .setRanges(statusRange)
      .build();

    sheet.setConditionalFormatRules([pendingRule, activeRule, suspendedRule, expiredRule, trialRule]);
  }
}

/**
 * Formats a Date object as dd-mmm-yyyy (e.g. 01-Nov-2026)
 */
function formatDateDdMmmYyyy(date) {
  if (!date) return '';
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

/**
 * Safely parses any date string (dd-mmm-yyyy or ISO) into a Date object
 */
function parseDateSafe(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const str = val.toString().trim();
  
  const parts = str.split('-');
  if (parts.length === 3) {
    const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    const d = parseInt(parts[0], 10);
    const m = parts[1].toLowerCase();
    const y = parseInt(parts[2], 10);
    if (!isNaN(d) && months[m] !== undefined && !isNaN(y)) {
      return new Date(y, months[m], d);
    }
  }
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Automatically records or updates a subscriber in Muhammad Haseeb's Master Sheet
 * Formats registration and expiry dates strictly as dd-mmm-yyyy (e.g. 01-Nov-2026)
 */
function recordOrUpdateSubscriber(userEmail, clientName, market) {
  const cleanEmail = (userEmail || '').toLowerCase().trim();
  if (!cleanEmail) {
    throw new Error('Email is required to record subscriber.');
  }

  const ss = getMasterSpreadsheet();
  const sheet = ss.getSheetByName('Subscribers') || ss.getActiveSheet();
  initMasterSheetHeaders(sheet);

  const data = sheet.getDataRange().getValues();
  const headerRow = (data[0] || []).map(h => (h || '').toString().toLowerCase().trim());
  
  let emailCol = headerRow.findIndex(h => h.includes('email') || h.includes('gmail') || h.includes('user'));
  let nameCol = headerRow.findIndex(h => h.includes('name') || h.includes('client') || h.includes('customer'));
  let regCol = headerRow.findIndex(h => h.includes('reg') || h.includes('join') || h.includes('start'));
  let statusCol = headerRow.findIndex(h => h.includes('status') || h.includes('state'));
  let expiryCol = headerRow.findIndex(h => h.includes('expir') || h.includes('valid') || h.includes('end'));
  let marketCol = headerRow.findIndex(h => h.includes('market') || h.includes('store') || h.includes('country'));
  let lastActiveCol = headerRow.findIndex(h => h.includes('active') || h.includes('last'));
  let notesCol = headerRow.findIndex(h => h.includes('note') || h.includes('remark'));

  // Sensible defaults if headers are generic
  if (emailCol === -1) emailCol = 0;
  if (nameCol === -1) nameCol = 1;
  if (regCol === -1) regCol = 2;
  if (statusCol === -1) statusCol = 3;
  if (expiryCol === -1) expiryCol = 4;
  if (marketCol === -1) marketCol = 5;
  if (lastActiveCol === -1) lastActiveCol = 6;
  if (notesCol === -1) notesCol = 7;

  // Format today's date strictly as dd-mmm-yyyy (e.g. 01-Nov-2026)
  const now = new Date();
  const regDateFormatted = formatDateDdMmmYyyy(now);
  
  // Default 30-day initial subscription period
  const expiryDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const expiryDateFormatted = formatDateDdMmmYyyy(expiryDate);

  const cleanName = (clientName || '').trim() || 'New Subscriber';
  const cleanMarket = (market || 'it').toLowerCase().trim();

  // Check if subscriber already exists
  for (let i = 1; i < data.length; i++) {
    const rowEmail = (data[i][emailCol] || '').toString().toLowerCase().trim();
    if (rowEmail === cleanEmail) {
      const rowNum = i + 1;
      
      // Update Name if it was blank or generic
      const currentValName = (data[i][nameCol] || '').toString().trim();
      if (cleanName && (!currentValName || currentValName === 'New Subscriber' || currentValName === 'Subscriber')) {
        sheet.getRange(rowNum, nameCol + 1).setValue(cleanName);
      }
      
      // Update Marketplace if provided
      if (marketCol !== -1 && cleanMarket) {
        sheet.getRange(rowNum, marketCol + 1).setValue(cleanMarket);
      }
      
      // Update Last Active date
      if (lastActiveCol !== -1) {
        sheet.getRange(rowNum, lastActiveCol + 1).setValue(regDateFormatted);
      }
      
      const currentStatus = (data[i][statusCol] || 'Pending').toString().trim();
      const currentExpiry = (data[i][expiryCol] || 'Pending Payment').toString().trim();

      return {
        action: 'EXISTING_UPDATED',
        email: cleanEmail,
        name: cleanName || data[i][nameCol],
        status: currentStatus,
        expiry: currentExpiry,
        registered_date: data[i][regCol] || regDateFormatted
      };
    }
  }

  // Not found: Append Brand New Subscriber Row
  // Default status is strictly PENDING so new users do not get free active access!
  const totalCols = Math.max(8, headerRow.length);
  const newRow = new Array(totalCols).fill('');
  newRow[emailCol] = cleanEmail;
  newRow[nameCol] = cleanName;
  newRow[regCol] = regDateFormatted;          // dd-mmm-yyyy (e.g. 01-Nov-2026)
  newRow[statusCol] = 'Pending';              // ⏳ PENDING by default
  newRow[expiryCol] = 'Pending Payment';      // Muhammad Haseeb sets expiry once paid
  newRow[marketCol] = cleanMarket;
  newRow[lastActiveCol] = regDateFormatted;
  newRow[notesCol] = 'New Registration - Payment Pending';

  sheet.appendRow(newRow);

  const newRowNum = sheet.getLastRow();
  sheet.getRange(newRowNum, 1, 1, newRow.length).setNumberFormat('@');
  sheet.getRange(newRowNum, 1, 1, newRow.length).setVerticalAlignment('middle');
  sheet.getRange(newRowNum, 3, 1, 1).setHorizontalAlignment('center'); // Reg Date
  sheet.getRange(newRowNum, 4, 1, 1).setHorizontalAlignment('center'); // Status
  sheet.getRange(newRowNum, 5, 1, 1).setHorizontalAlignment('center'); // Expiry
  sheet.getRange(newRowNum, 6, 1, 1).setHorizontalAlignment('center'); // Market

  return {
    action: 'NEW_REGISTERED',
    email: cleanEmail,
    name: cleanName,
    status: 'Pending',
    expiry: 'Pending Payment',
    registered_date: regDateFormatted
  };
}

/**
 * Validates subscriber Gmail against Master Google Sheet
 */
function checkLicense(userEmail) {
  const cleanEmail = (userEmail || '').toLowerCase().trim();
  if (!cleanEmail) {
    return {
      valid: false,
      message: '❌ License Error: Google Account Gmail could not be verified. Please enter your registered Gmail.'
    };
  }

  try {
    const ss = getMasterSpreadsheet();
    const sheet = ss.getSheetByName('Subscribers') || ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();

    if (!data || data.length <= 1) {
      return {
        valid: false,
        message: 'Master Subscribers Sheet is empty. No subscribers registered yet.'
      };
    }

    const headerRow = data[0].map(h => (h || '').toString().toLowerCase().trim());
    let emailCol = headerRow.findIndex(h => h.includes('email') || h.includes('gmail') || h.includes('user'));
    let nameCol = headerRow.findIndex(h => h.includes('name') || h.includes('client') || h.includes('customer'));
    let statusCol = headerRow.findIndex(h => h.includes('status') || h.includes('state'));
    let expiryCol = headerRow.findIndex(h => h.includes('expir') || h.includes('valid') || h.includes('end'));
    let lastActiveCol = headerRow.findIndex(h => h.includes('active') || h.includes('last'));

    if (emailCol === -1) emailCol = 0;
    if (nameCol === -1) nameCol = 1;
    if (statusCol === -1) statusCol = 3;
    if (expiryCol === -1) expiryCol = 4;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = (data[i][emailCol] || '').toString().toLowerCase().trim();
      if (rowEmail === cleanEmail) {
        const name = (nameCol !== -1 && data[i][nameCol] ? data[i][nameCol] : 'Subscriber').toString().trim();
        const status = (statusCol !== -1 && data[i][statusCol] ? data[i][statusCol] : 'Active').toString().trim();
        const expiryRaw = expiryCol !== -1 ? data[i][expiryCol] : null;
        const expiryDate = parseDateSafe(expiryRaw);

        // Update Last Active timestamp on check
        if (lastActiveCol !== -1) {
          try {
            sheet.getRange(i + 1, lastActiveCol + 1).setValue(formatDateDdMmmYyyy(new Date()));
          } catch (e) {}
        }

        // Check active status
        if (status.toLowerCase() !== 'active') {
          if (status.toLowerCase() === 'pending') {
            return {
              valid: false,
              status: 'Pending',
              name: name,
              email: cleanEmail,
              expiry: expiryRaw || 'Pending Payment',
              message: `⏳ Subscription Pending Approval: Account [${cleanEmail}] abhi Pending hai. Muhammad Haseeb se contact karein taake payment verify hone par aapka status "Active" kiya ja sake.`
            };
          }
          return {
            valid: false,
            status: status,
            name: name,
            email: cleanEmail,
            expiry: expiryRaw || 'N/A',
            message: `🛑 Subscription Inactive: Account [${cleanEmail}] status is "${status}". Please contact Muhammad Haseeb to activate your subscription.`
          };
        }

        // Check expiry date
        if (expiryDate && !isNaN(expiryDate.getTime())) {
          const today = new Date();
          today.setHours(0, 0, 0, 0);

          const expFormatted = formatDateDdMmmYyyy(expiryDate);
          if (today > expiryDate) {
            return {
              valid: false,
              status: 'Expired',
              name: name,
              email: cleanEmail,
              expiry: expFormatted,
              message: `⏰ Subscription Expired: Your plan for [${cleanEmail}] expired on ${expFormatted}. Please pay your monthly renewal fee to Muhammad Haseeb to resume.`
            };
          }

          return {
            valid: true,
            status: 'Active',
            name: name,
            email: cleanEmail,
            expiry: expFormatted
          };
        }

        return {
          valid: true,
          status: 'Active',
          name: name,
          email: cleanEmail,
          expiry: 'Lifetime / Ongoing'
        };
      }
    }

    // Email not found in Master Sheet
    return {
      valid: false,
      message: `🚫 Unlicensed Copy: Google account [${cleanEmail}] is not registered in the Master Subscriber database. Please contact Muhammad Haseeb to subscribe.`
    };

  } catch (err) {
    console.error('License check error:', err);
    return {
      valid: false,
      message: 'Subscription Server Error: ' + err.toString()
    };
  }
}

/**
 * Main Web App POST Handler
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No post data received.');
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const userEmail = (payload.user_email || '').toString().toLowerCase().trim();

    // 0. License & Subscription Verification Action
    if (action === 'CHECK_LICENSE') {
      const license = checkLicense(userEmail);
      return createJsonResponse(license);
    }

    // 1. Auto-Record Subscriber & Generate Secure OAuth URL on Step 1
    if (action === 'RECORD_SUBSCRIBER') {
      const clientName = (payload.client_name || '').toString().trim();
      const marketplace = (payload.marketplace || 'it').toString().trim();
      const result = recordOrUpdateSubscriber(userEmail, clientName, marketplace);
      const authUrl = generateAuthUrl(marketplace);

      return createJsonResponse({
        status: 'success',
        data: result,
        auth_url: authUrl
      });
    }

    // 1b. Direct OAuth URL Request
    if (action === 'GET_AUTH_URL') {
      const marketplace = (payload.marketplace || 'it').toString().trim();
      return createJsonResponse({
        status: 'success',
        auth_url: generateAuthUrl(marketplace)
      });
    }

    // 2. Exchange temporary authorization code for 18-month Refresh Token
    if (action === 'EXCHANGE_CODE') {
      // Ensure user is recorded in Master Sheet
      if (userEmail) {
        recordOrUpdateSubscriber(userEmail, payload.client_name, payload.marketplace);
      }

      // License Check
      const license = checkLicense(userEmail);
      const isPending = !license.valid && license.status === 'Pending';

      // If user is suspended or expired (not pending registration), block exchange
      if (!license.valid && !isPending) {
        return createJsonResponse({ status: 'error', message: license.message });
      }

      const authCode = payload.code;
      if (!authCode) throw new Error('Authorization code is missing.');

      const basicAuth = Utilities.base64Encode(`${SECRETS.APP_ID.trim()}:${SECRETS.CERT_ID.trim()}`);
      const ebayResponse = UrlFetchApp.fetch(SECRETS.OAUTH_ENDPOINT, {
        method: 'POST',
        headers: {
          'Authorization': 'Basic ' + basicAuth,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
          grant_type: 'authorization_code',
          code: authCode,
          redirect_uri: SECRETS.RU_NAME
        },
        muteHttpExceptions: true
      });

      const data = JSON.parse(ebayResponse.getContentText());
      if (ebayResponse.getResponseCode() !== 200 || !data.refresh_token) {
        return createJsonResponse({
          status: 'error',
          message: data.error_description || ebayResponse.getContentText()
        });
      }

      return createJsonResponse({
        status: 'success',
        refresh_token: data.refresh_token,
        access_token: data.access_token,
        expires_in: data.expires_in,
        is_pending: isPending,
        license_status: isPending ? 'Pending' : (license.status || 'Active'),
        pending_message: isPending ? license.message : ''
      });
    }

    // 3. Refresh 2-hour Access Token using buyer's Refresh Token
    if (action === 'REFRESH_TOKEN') {
      // License Check
      const license = checkLicense(userEmail);
      if (!license.valid) {
        return createJsonResponse({ status: 'error', message: license.message });
      }

      const refreshToken = payload.refresh_token;
      if (!refreshToken) throw new Error('Refresh token is missing.');

      const basicAuth = Utilities.base64Encode(`${SECRETS.APP_ID.trim()}:${SECRETS.CERT_ID.trim()}`);
      const ebayResponse = UrlFetchApp.fetch(SECRETS.OAUTH_ENDPOINT, {
        method: 'POST',
        headers: {
          'Authorization': 'Basic ' + basicAuth,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          scope: 'https://api.ebay.com/oauth/api_scope/sell.fulfillment https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly'
        },
        muteHttpExceptions: true
      });

      const data = JSON.parse(ebayResponse.getContentText());
      if (ebayResponse.getResponseCode() !== 200 || !data.access_token) {
        return createJsonResponse({
          status: 'error',
          message: data.error_description || ebayResponse.getContentText()
        });
      }

      return createJsonResponse({
        status: 'success',
        access_token: data.access_token,
        expires_in: data.expires_in
      });
    }

    throw new Error('Invalid action requested.');

  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    });
  }
}

/**
 * Health Check GET Endpoint
 */
function doGet(e) {
  return createJsonResponse({
    status: 'online',
    service: 'eBay OAuth Relay & Master Subscription Server',
    server_time: formatDateDdMmmYyyy(new Date()),
    message: 'Relay server is active and operational.'
  });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Optional Manual Test Function:
 * Run this from Apps Script editor to test adding a subscriber manually!
 */
function testRecordSubscriber() {
  const result = recordOrUpdateSubscriber('testbuyer@gmail.com', 'Test Buyer', 'it');
  Logger.log('Record result: ' + JSON.stringify(result));
}
