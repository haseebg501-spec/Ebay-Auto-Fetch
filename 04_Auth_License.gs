/**
 * 04_Auth_License.gs
 * OAuth 2.0 Engine, Background Token Refresh, License Verification & Store Connection (Steps 1 & 2)
 */

// ==========================================
// 1. TOKEN REFRESH & BACKGROUND ENGINE
// ==========================================

/**
 * Safely calls Muhammad Haseeb's Master Relay Web App.
 * Automatically catches HTTP 404 Not Found, Google sign-in pages, and HTML responses
 * so users never see the cryptic "Unexpected token <" error!
 */
function callAuthRelayServer(payload) {
  const relayUrl = EBAY_CONFIG.AUTH_RELAY_URL;
  if (!relayUrl || relayUrl.includes('PASTE_YOUR_AUTH_RELAY_URL_HERE')) {
    throw new Error('AUTH_RELAY_URL is not configured. Please paste your Master Relay Web App URL in 01_Config_Menu.gs.');
  }

  const response = UrlFetchApp.fetch(relayUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const rawText = (response.getContentText() || '').trim();

  // 1. Check for 404 Not Found
  if (statusCode === 404) {
    throw new Error(
      'Master Relay Web App URL Not Found (HTTP 404).\n\n' +
      'Google could not find this Web App. When you deploy EbayAuthRelay.gs in your Master Sheet, Google generates a Web App URL ending in /exec.\n\n' +
      'Please check your Master Sheet > Deploy > Manage Deployments, copy the Web App URL, and paste it into AUTH_RELAY_URL in 01_Config_Menu.gs.'
    );
  }

  // 2. Check for HTML response (Google Sign-In page or permission block)
  if (rawText.startsWith('<') || rawText.includes('<!DOCTYPE') || rawText.includes('<html')) {
    if (rawText.includes('accounts.google.com') || rawText.includes('Sign in')) {
      throw new Error(
        'Master Relay Permission Error (Google Sign-in Required).\n\n' +
        'Your Web App is currently restricted. Please edit your deployment in Master Sheet:\n' +
        '1. Click Deploy > Manage Deployments > Edit (Pencil)\n' +
        '2. Set "Who has access" to "Anyone" (NOT "Only myself")\n' +
        '3. Select "New version" and click Deploy.'
      );
    }
    throw new Error(
      `Master Relay returned an HTML page (Status ${statusCode}) instead of JSON data.\n\n` +
      'Please verify that your Master Sheet Web App is deployed with "Who has access: Anyone".'
    );
  }

  // 3. Parse JSON safely
  let data;
  try {
    data = JSON.parse(rawText);
  } catch (err) {
    throw new Error(`Failed to parse Relay Server response (Status ${statusCode}): ${rawText.slice(0, 150)}`);
  }

  return data;
}

/**
 * Retrieves a valid eBay Access Token.
 * Automatically refreshes the 2-hour token via Master Relay when expired.
 * 100% safe for background time-based triggers!
 */
function getValidAccessToken() {
  const props = PropertiesService.getScriptProperties();

  const cachedToken = props.getProperty('CACHED_ACCESS_TOKEN');
  const tokenExpiry = parseInt(props.getProperty('ACCESS_TOKEN_EXPIRY') || '0', 10);
  const nowMs = Date.now();

  // If cached token has > 5 minutes remaining, use it
  if (cachedToken && tokenExpiry > nowMs + 300000) {
    return cachedToken;
  }

  const refreshToken = props.getProperty('EBAY_REFRESH_TOKEN');
  if (!refreshToken) {
    throw new Error('eBay store not connected yet. Please click "⚡ eBay Sync > 🔑 1. Authorize My eBay Store" from the menu.');
  }

  const userEmail = getLicensedUserEmail();

  const data = callAuthRelayServer({
    action: 'REFRESH_TOKEN',
    refresh_token: refreshToken.trim(),
    user_email: userEmail
  });

  if (data.status === 'error' || !data.access_token) {
    throw new Error('Failed to refresh eBay token: ' + (data.message || JSON.stringify(data)));
  }

  const freshToken = data.access_token;
  const expiresInMs = (data.expires_in || 7200) * 1000;

  props.setProperty('CACHED_ACCESS_TOKEN', freshToken);
  props.setProperty('ACCESS_TOKEN_EXPIRY', (nowMs + expiresInMs).toString());

  return freshToken;
}

// ==========================================
// 2. LICENSED GMAIL & USER IDENTITY
// ==========================================

/**
 * Safely retrieves user's registered Gmail across all properties and session scopes.
 * Checks ScriptProperties FIRST so that 15-minute and 1-hour background triggers
 * always have access to the email even without an active browser session!
 */
function getLicensedUserEmail() {
  let email = '';

  // 1. ScriptProperties (Persistent globally, always available in background triggers)
  try {
    email = PropertiesService.getScriptProperties().getProperty('LICENSED_EMAIL') || '';
  } catch (e) {}

  // 2. DocumentProperties
  if (!email) {
    try {
      email = PropertiesService.getDocumentProperties().getProperty('LICENSED_EMAIL') || '';
    } catch (e) {}
  }

  // 3. UserProperties
  if (!email) {
    try {
      email = PropertiesService.getUserProperties().getProperty('LICENSED_EMAIL') || '';
    } catch (e) {}
  }

  // 4. Session Active User (Interactive session only)
  if (!email) {
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {}
  }

  // 5. Effective User
  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch (e) {}
  }

  // 6. Spreadsheet Owner
  if (!email) {
    try {
      email = SpreadsheetApp.getActiveSpreadsheet().getOwner()?.getEmail() || '';
    } catch (e) {}
  }

  return (email || '').toLowerCase().trim();
}

function getClientName() {
  let name = '';
  try {
    name = PropertiesService.getScriptProperties().getProperty('CLIENT_NAME') || '';
  } catch (e) {}

  if (!name) {
    try {
      name = PropertiesService.getDocumentProperties().getProperty('CLIENT_NAME') || '';
    } catch (e) {}
  }

  return (name || '').trim();
}

function checkRemoteLicense(email) {
  if (!EBAY_CONFIG.AUTH_RELAY_URL || EBAY_CONFIG.AUTH_RELAY_URL.includes('PASTE_YOUR_AUTH_RELAY_URL_HERE')) {
    return { valid: true, status: 'Active (Development)', expiry: 'No Relay URL Set' };
  }
  const cleanEmail = (email || getLicensedUserEmail()).toLowerCase().trim();
  try {
    const data = callAuthRelayServer({
      action: 'CHECK_LICENSE',
      user_email: cleanEmail
    });
    return data;
  } catch (err) {
    return { valid: false, message: err.message || err.toString() };
  }
}

function promptChangeLicensedEmail() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = getLicensedUserEmail();
  const prompt = ui.prompt(
    '🔑 Set / Update Registered Gmail',
    `Current Registered Gmail: ${currentEmail || '(None set)'}\n\nEnter the Gmail address you registered with the seller:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() === ui.Button.OK) {
    const newEmail = prompt.getResponseText().trim().toLowerCase();
    if (newEmail && newEmail.includes('@')) {
      PropertiesService.getScriptProperties().setProperty('LICENSED_EMAIL', newEmail);
      PropertiesService.getDocumentProperties().setProperty('LICENSED_EMAIL', newEmail);
      PropertiesService.getUserProperties().setProperty('LICENSED_EMAIL', newEmail);
      showSubscriptionStatusModal();
    } else {
      ui.alert('Invalid Email', 'Please enter a valid Gmail address (e.g. name@gmail.com).', ui.ButtonSet.OK);
    }
  }
}

function showSubscriptionStatusModal() {
  const ui = SpreadsheetApp.getUi();
  let userEmail = getLicensedUserEmail();

  if (!userEmail) {
    const prompt = ui.prompt(
      '🔑 License Activation',
      'Please enter the Gmail address you used for your subscription:',
      ui.ButtonSet.OK_CANCEL
    );

    if (prompt.getSelectedButton() === ui.Button.OK) {
      userEmail = prompt.getResponseText().trim().toLowerCase();
      if (userEmail && userEmail.includes('@')) {
        PropertiesService.getScriptProperties().setProperty('LICENSED_EMAIL', userEmail);
        PropertiesService.getDocumentProperties().setProperty('LICENSED_EMAIL', userEmail);
        PropertiesService.getUserProperties().setProperty('LICENSED_EMAIL', userEmail);
      }
    }
  }

  if (!userEmail) {
    ui.alert('Notice', 'No registered Gmail provided. Please click "Change / Set Gmail" to activate your license.', ui.ButtonSet.OK);
    return;
  }

  const info = checkRemoteLicense(userEmail);

  const isPending = (info.status === 'Pending') || (!info.valid && info.message && info.message.includes('Pending'));
  let statusColor = info.valid ? '#0F9D58' : (isPending ? '#D97706' : '#D93025');
  let statusBadge = info.valid ? (info.status || 'ACTIVE') : (isPending ? 'PENDING APPROVAL' : 'INACTIVE / EXPIRED');
  const expiry = info.expiry || (isPending ? 'Pending Payment' : 'N/A');
  const name = info.name || getClientName() || (info.valid ? 'Active Subscriber' : (isPending ? 'Pending Subscriber' : 'Unregistered User'));

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 20px; color: #1f2937; line-height: 1.5; background: #f9fafb; font-size: 13px; }
        .card { border: 1px solid #e5e7eb; border-radius: 10px; padding: 18px; background: #ffffff; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; border-bottom: 1px solid #f3f4f6; padding-bottom: 8px; }
        .row:last-child { border: none; margin-bottom: 0; padding-bottom: 0; }
        .label { font-weight: 600; color: #6b7280; }
        .val { font-weight: 700; color: #111827; }
        .badge { background: ${statusColor}; color: white; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; }
        .alert-box { background: ${info.valid ? '#E8F8F0' : '#FEE2E2'}; border: 1px solid ${info.valid ? '#A7F3D0' : '#FCA5A5'}; color: ${info.valid ? '#065F46' : '#991B1B'}; padding: 12px; border-radius: 8px; font-size: 12px; line-height: 1.4; margin-bottom: 14px; }
        .btn-change { display: inline-block; width: 100%; text-align: center; background: #f3f4f6; color: #374151; padding: 8px; border-radius: 6px; font-size: 12px; font-weight: 600; border: 1px solid #d1d5db; cursor: pointer; text-decoration: none; }
        .btn-change:hover { background: #e5e7eb; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="row">
          <span class="label">Google Account:</span>
          <span class="val">${userEmail}</span>
        </div>
        <div class="row">
          <span class="label">Client Name:</span>
          <span class="val">${name}</span>
        </div>
        <div class="row">
          <span class="label">License Status:</span>
          <span class="badge">${statusBadge}</span>
        </div>
        <div class="row">
          <span class="label">Valid Until:</span>
          <span class="val" style="color: ${statusColor};">${expiry}</span>
        </div>
      </div>

      <div class="alert-box">
        ${info.valid 
          ? '<b>✓ License Verified:</b> Your Google Sheet is authorized to receive 24/7 background eBay order syncs.'
          : '<b>🛑 Attention:</b> ' + (info.message || 'Subscription inactive. Please contact the administrator.')}
      </div>

      <button class="btn-change" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).promptChangeLicensedEmail()">✏️ Change / Update Registered Gmail</button>
    </body>
    </html>
  `).setWidth(440).setHeight(380);

  ui.showModalDialog(html, '🛡️ Subscription & License Status');
}

// ==========================================
// 3. STEP 1 & STEP 2 AUTHORIZATION
// ==========================================

/**
 * Step 1: Shows interactive Marketplace & Registration Dialog before opening Auth
 */
function openAuthorizationUrl() {
  const ui = SpreadsheetApp.getUi();
  const currentMarket = PropertiesService.getScriptProperties().getProperty('EBAY_MARKETPLACE') || 'it';
  const savedEmail = getLicensedUserEmail();
  const savedName = getClientName();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 20px; line-height: 1.5; color: #111827; background: #ffffff; font-size: 13px; }
        h3 { font-size: 16px; font-weight: 700; margin-bottom: 6px; color: #111827; }
        p { font-size: 12.5px; color: #4b5563; margin-bottom: 16px; }
        .form-group { margin-bottom: 12px; }
        label { font-size: 11.5px; font-weight: 700; color: #374151; display: block; margin-bottom: 5px; text-transform: uppercase; letter-spacing: 0.5px; }
        input[type="text"], input[type="email"], select { width: 100%; padding: 9px 12px; border-radius: 6px; border: 1px solid #d1d5db; font-size: 13.5px; color: #111827; background: #f9fafb; outline: none; }
        input:focus, select:focus { border-color: #0F9D58; box-shadow: 0 0 0 2px rgba(15,157,88,0.2); background: #ffffff; }
        .btn-container { text-align: center; margin-top: 16px; margin-bottom: 12px; }
        .btn { display: inline-block; width: 100%; padding: 11px 20px; background: #0F9D58; color: #ffffff; font-weight: 700; font-size: 13.5px; border-radius: 6px; border: none; cursor: pointer; transition: background 0.15s ease; box-shadow: 0 2px 6px rgba(15,157,88,0.25); }
        .btn:hover { background: #0b8043; }
        .btn:disabled { background: #9ca3af; cursor: not-allowed; }
        .hint-box { background: #f3f4f6; border: 1px solid #e5e7eb; padding: 10px 12px; border-radius: 6px; font-size: 11.5px; color: #4b5563; line-height: 1.4; }
        .status-box { display: none; margin-top: 12px; padding: 10px 12px; border-radius: 6px; font-size: 12px; line-height: 1.4; }
        .status-loading { display: block; background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af; text-align: center; }
        .status-success { display: block; background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
        .status-warning { display: block; background: #fffbeb; border: 1px solid #fef3c7; color: #92400e; }
        .status-error { display: block; background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
      </style>
    </head>
    <body>
      <h3>Step 1: Connect eBay Store</h3>
      <p>Confirm your registration details and select your eBay marketplace:</p>

      <div class="form-group">
        <label for="clientName">Your Name</label>
        <input type="text" id="clientName" placeholder="e.g. John Doe" value="${savedName}">
      </div>

      <div class="form-group">
        <label for="clientEmail">Your Gmail (License Account)</label>
        <input type="email" id="clientEmail" placeholder="e.g. yourname@gmail.com" value="${savedEmail}">
      </div>

      <div class="form-group">
        <label for="marketSelect">eBay Marketplace</label>
        <select id="marketSelect">
          <option value="it" ${currentMarket === 'it' ? 'selected' : ''}>🇮🇹 Italy (ebay.it)</option>
          <option value="com" ${currentMarket === 'com' ? 'selected' : ''}>🇺🇸 United States (ebay.com)</option>
          <option value="co.uk" ${currentMarket === 'co.uk' ? 'selected' : ''}>🇬🇧 United Kingdom (ebay.co.uk)</option>
          <option value="de" ${currentMarket === 'de' ? 'selected' : ''}>🇩🇪 Germany (ebay.de)</option>
          <option value="fr" ${currentMarket === 'fr' ? 'selected' : ''}>🇫🇷 France (ebay.fr)</option>
          <option value="es" ${currentMarket === 'es' ? 'selected' : ''}>🇪🇸 Spain (ebay.es)</option>
          <option value="com.au" ${currentMarket === 'com.au' ? 'selected' : ''}>🇦🇺 Australia (ebay.com.au)</option>
          <option value="ca" ${currentMarket === 'ca' ? 'selected' : ''}>🇨🇦 Canada (ebay.ca)</option>
          <option value="ie" ${currentMarket === 'ie' ? 'selected' : ''}>🇮🇪 Ireland (ebay.ie)</option>
          <option value="at" ${currentMarket === 'at' ? 'selected' : ''}>🇦🇹 Austria (ebay.at)</option>
          <option value="ch" ${currentMarket === 'ch' ? 'selected' : ''}>🇨🇭 Switzerland (ebay.ch)</option>
          <option value="nl" ${currentMarket === 'nl' ? 'selected' : ''}>🇳🇱 Netherlands (ebay.nl)</option>
          <option value="be" ${currentMarket === 'be' ? 'selected' : ''}>🇧🇪 Belgium (ebay.be)</option>
          <option value="pl" ${currentMarket === 'pl' ? 'selected' : ''}>🇵🇱 Poland (ebay.pl)</option>
        </select>
      </div>

      <div class="btn-container">
        <button class="btn" id="btnAuth" onclick="startAuthorization()">Authorize My eBay Store &rarr;</button>
      </div>

      <div id="statusBox" class="status-box"></div>

      <div class="hint-box" style="margin-top: 12px;">
        🛡️ <b>Automatic License Registration:</b> Clicking authorize will register your account with Muhammad Haseeb's Master Sheet and open the native eBay sign-in portal.
      </div>

      <script>
        function startAuthorization() {
          const nameInput = document.getElementById('clientName');
          const emailInput = document.getElementById('clientEmail');
          const marketSelect = document.getElementById('marketSelect');
          const btn = document.getElementById('btnAuth');
          const statusBox = document.getElementById('statusBox');

          const name = (nameInput.value || '').trim();
          const email = (emailInput.value || '').trim();
          const market = marketSelect.value;

          if (!name) {
            alert('Please enter your Name.');
            nameInput.focus();
            return;
          }

          if (!email || email.indexOf('@') === -1) {
            alert('Please enter a valid Gmail address (e.g. name@gmail.com).');
            emailInput.focus();
            return;
          }

          btn.disabled = true;
          btn.innerText = '⏳ Connecting & Registering...';
          statusBox.className = 'status-box status-loading';
          statusBox.innerHTML = 'Connecting to Master License Server and registering subscriber...';
          statusBox.style.display = 'block';

          google.script.run
            .withSuccessHandler(function(res) {
              btn.disabled = false;
              btn.innerText = 'Authorize My eBay Store &rarr;';

              if (res && res.authUrl) {
                window.open(res.authUrl, '_blank');

                if (res.status === 'warning') {
                  statusBox.className = 'status-box status-warning';
                  statusBox.innerHTML = '⚠️ <b>Note:</b> ' + (res.warning || 'Server note') + '<br><br><b>eBay login tab opened!</b> Sign in on eBay, copy the redirect URL, and run Step 2.';
                } else {
                  statusBox.className = 'status-box status-success';
                  statusBox.innerHTML = '✅ <b>Registration Successful!</b><br>Account registered in Master Database.<br><br><b>eBay sign-in tab opened!</b> Click "Agree and Continue", then copy the URL and run <b>Step 2</b> from the menu.';
                }
              } else {
                statusBox.className = 'status-box status-error';
                statusBox.innerHTML = '❌ Error: Failed to generate authorization URL.';
              }
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.innerText = 'Authorize My eBay Store &rarr;';
              statusBox.className = 'status-box status-error';
              statusBox.innerHTML = '❌ Connection Error: ' + err.message;
            })
            .processStep1Authorization(name, email, market);
        }
      </script>
    </body>
    </html>
  `).setWidth(480).setHeight(480);
  ui.showModalDialog(html, 'eBay Store Authorization (Step 1)');
}

function processStep1Authorization(name, email, market) {
  const cleanName = (name || '').trim();
  const cleanEmail = (email || '').toLowerCase().trim();
  const cleanMarket = (market || 'it').toLowerCase().trim();

  if (!cleanEmail || !cleanEmail.includes('@')) {
    throw new Error('Please enter a valid Gmail address.');
  }

  // Save to ScriptProperties (guarantees availability in background triggers!)
  PropertiesService.getScriptProperties().setProperty('LICENSED_EMAIL', cleanEmail);
  PropertiesService.getDocumentProperties().setProperty('LICENSED_EMAIL', cleanEmail);
  PropertiesService.getUserProperties().setProperty('LICENSED_EMAIL', cleanEmail);

  if (cleanName) {
    PropertiesService.getScriptProperties().setProperty('CLIENT_NAME', cleanName);
    PropertiesService.getDocumentProperties().setProperty('CLIENT_NAME', cleanName);
    PropertiesService.getUserProperties().setProperty('CLIENT_NAME', cleanName);
  }

  PropertiesService.getScriptProperties().setProperty('EBAY_MARKETPLACE', cleanMarket);
  saveMarketplacePreference(cleanMarket);

  const data = callAuthRelayServer({
    action: 'RECORD_SUBSCRIBER',
    user_email: cleanEmail,
    client_name: cleanName,
    marketplace: cleanMarket
  });

  if (data.status === 'error' || !data.auth_url) {
    throw new Error(data.message || 'Failed to generate authorization URL from Master Relay Server.');
  }

  return {
    status: 'success',
    authUrl: data.auth_url
  };
}

function saveMarketplacePreference(market) {
  PropertiesService.getScriptProperties().setProperty('EBAY_MARKETPLACE', market);
  try {
    const sheet = getOrCreateOrdersSheet();
    const defaultCurrency = getMarketplaceDefaultCurrency(market);
    const kpiFormat = getCurrencyNumberFormat(defaultCurrency);
    sheet.getRange('D3').setNumberFormat(kpiFormat);
    sheet.getRange('G3').setNumberFormat(kpiFormat);
    sheet.getRange('J3').setNumberFormat(kpiFormat);
  } catch (e) {
    console.warn('KPI format update warning:', e);
  }
}

function promptExchangeAuthCode() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  if (!EBAY_CONFIG.AUTH_RELAY_URL || EBAY_CONFIG.AUTH_RELAY_URL.includes('PASTE_YOUR_AUTH_RELAY_URL_HERE')) {
    ui.alert('Setup Required', 'AUTH_RELAY_URL is not set.', ui.ButtonSet.OK);
    return;
  }

  const prompt = ui.prompt(
    'Step 2: Connect eBay Store',
    'Paste the URL from your browser address bar after clicking "Agree and Continue":',
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  let rawInput = prompt.getResponseText().trim();

  let authCode = rawInput;
  if (rawInput.includes('code=')) {
    const match = rawInput.match(/[?&]code=([^&]+)/);
    if (match) {
      authCode = decodeURIComponent(match[1]);
    }
  }

  try {
    const userEmail = getLicensedUserEmail();
    const clientName = getClientName();
    const market = PropertiesService.getScriptProperties().getProperty('EBAY_MARKETPLACE') || 'it';

    const data = callAuthRelayServer({
      action: 'EXCHANGE_CODE',
      code: authCode,
      user_email: userEmail,
      client_name: clientName,
      marketplace: market
    });

    if (data.status === 'error' || !data.refresh_token) {
      throw new Error(data.message || 'Token exchange failed.');
    }

    props.setProperty('EBAY_REFRESH_TOKEN', data.refresh_token);
    props.setProperty('CACHED_ACCESS_TOKEN', data.access_token);
    props.setProperty('ACCESS_TOKEN_EXPIRY', (Date.now() + (data.expires_in || 7200) * 1000).toString());

    if (data.is_pending) {
      ui.alert(
        '⏳ eBay Linked (Pending Approval)',
        'Your eBay store has been linked successfully!\n\n' +
        'NOTE: Your subscription is currently "PENDING APPROVAL" in Muhammad Haseeb\'s Master Sheet.\n\n' +
        'Please contact Muhammad Haseeb to activate your paid subscription so you can start syncing orders.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('🎉 SUCCESS!', 'Your eBay store is now connected!\n\nYou can now click "Sync Orders (Last 15 Days)" or enable the 15-Minute Auto Sync.', ui.ButtonSet.OK);
    }

  } catch (err) {
    ui.alert('❌ Connection Failed', err.toString(), ui.ButtonSet.OK);
  }
}

