/**
 * 05_UserGuides.gs
 * Interactive Modals: Sync Condition Settings, Custom Date Range, Quick Start Guide & FAQs
 */

// ==========================================
// 0. SYNC CONDITION / MODE SELECTOR MODAL
// ==========================================

function openSyncSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentMode = getSyncModePreference();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 20px; color: #1f2937; background: #ffffff; line-height: 1.5; font-size: 13px; }
        h3 { font-size: 16px; font-weight: 700; color: #111827; margin-bottom: 6px; }
        p.desc { font-size: 12.5px; color: #4b5563; margin-bottom: 16px; }
        .mode-card {
          border: 1.5px solid #e5e7eb;
          border-radius: 8px;
          padding: 14px;
          margin-bottom: 12px;
          cursor: pointer;
          transition: border-color 0.15s, background 0.15s;
        }
        .mode-card:hover { border-color: #0F9D58; background: #f0fdf4; }
        .mode-card.selected { border-color: #0F9D58; background: #f0fdf4; }
        .radio-header { display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 13.5px; color: #111827; margin-bottom: 4px; }
        .radio-header input { transform: scale(1.15); accent-color: #0F9D58; }
        .mode-details { font-size: 12px; color: #4b5563; margin-left: 24px; line-height: 1.4; }
        .badge { display: inline-block; background: #DCFCE7; color: #166534; font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-left: 6px; }
        .btn-container { display: flex; gap: 10px; margin-top: 20px; }
        .btn { flex: 1; padding: 11px; border-radius: 6px; font-size: 13px; font-weight: 700; border: none; cursor: pointer; text-align: center; }
        .btn-save { background: #f3f4f6; color: #374151; border: 1px solid #d1d5db; }
        .btn-save:hover { background: #e5e7eb; }
        .btn-sync { background: #0F9D58; color: #ffffff; }
        .btn-sync:hover { background: #0b8043; }
        .status-msg { display: none; text-align: center; margin-top: 12px; font-size: 12px; font-weight: 600; padding: 8px; border-radius: 6px; }
      </style>
    </head>
    <body>
      <h3>⚙️ Choose Sync Mode Condition</h3>
      <p class="desc">Select your preferred sync condition for background auto-sync and manual runs:</p>

      <div class="mode-card ${currentMode === 'NEW_AND_UPDATE' ? 'selected' : ''}" onclick="selectMode('NEW_AND_UPDATE')">
        <div class="radio-header">
          <input type="radio" name="syncMode" id="mode1" value="NEW_AND_UPDATE" ${currentMode === 'NEW_AND_UPDATE' ? 'checked' : ''}>
          <label for="mode1">1) Sync New Orders & Update Status <span class="badge">Recommended</span></label>
        </div>
        <div class="mode-details">
          Fast and smart incremental sync (Last 15 Days). Inserts brand new incoming orders, and updates existing pending orders when marked shipped or modified on eBay.
        </div>
      </div>

      <div class="mode-card ${currentMode === 'SYNC_ALL' ? 'selected' : ''}" onclick="selectMode('SYNC_ALL')">
        <div class="radio-header">
          <input type="radio" name="syncMode" id="mode2" value="SYNC_ALL" ${currentMode === 'SYNC_ALL' ? 'checked' : ''}>
          <label for="mode2">2) Sync All (New + Previous Orders)</label>
        </div>
        <div class="mode-details">
          Full historical scan of all past orders up to 90 days. Inserts any previous missing orders and refreshes all existing order statuses.
        </div>
      </div>

      <div id="statusMsg" class="status-msg"></div>

      <div class="btn-container">
        <button class="btn btn-save" onclick="savePreferenceOnly()">💾 Save as Default</button>
        <button class="btn btn-sync" onclick="saveAndRunSync()">⚡ Save & Sync Now</button>
      </div>

      <script>
        function selectMode(val) {
          document.getElementById(val === 'NEW_AND_UPDATE' ? 'mode1' : 'mode2').checked = true;
          document.querySelectorAll('.mode-card').forEach(c => c.classList.remove('selected'));
          if (event && event.currentTarget) event.currentTarget.classList.add('selected');
        }

        function getSelected() {
          return document.querySelector('input[name="syncMode"]:checked').value;
        }

        function savePreferenceOnly() {
          const mode = getSelected();
          showStatus('Saving preference...', '#eff6ff', '#1e40af');
          google.script.run
            .withSuccessHandler(function() {
              showStatus('✓ Preference saved! Background triggers will use this condition.', '#f0fdf4', '#166534');
              setTimeout(function(){ google.script.host.close(); }, 1200);
            })
            .saveSyncModePreference(mode);
        }

        function saveAndRunSync() {
          const mode = getSelected();
          showStatus('Saving preference and starting sync...', '#eff6ff', '#1e40af');
          google.script.run.saveSyncModePreference(mode);
          if (mode === 'SYNC_ALL') {
            google.script.run.syncAllOrdersFullHistory(false);
          } else {
            google.script.run.syncNewOrdersAndUpdateStatus(false);
          }
          google.script.host.close();
        }

        function showStatus(txt, bg, color) {
          const m = document.getElementById('statusMsg');
          m.style.display = 'block';
          m.style.background = bg;
          m.style.color = color;
          m.innerHTML = txt;
        }
      </script>
    </body>
    </html>
  `).setWidth(490).setHeight(430);

  ui.showModalDialog(html, 'eBay Sync Condition Settings');
}

// ==========================================
// 1. CUSTOM DATE RANGE SELECTOR MODAL
// ==========================================

function openCustomDateRangeDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentMode = getSyncModePreference();

  const now = new Date();
  const past30 = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  const ssTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'GMT';
  const defaultEnd = Utilities.formatDate(now, ssTz, 'yyyy-MM-dd');
  const defaultStart = Utilities.formatDate(past30, ssTz, 'yyyy-MM-dd');

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 20px; color: #1f2937; background: #ffffff; line-height: 1.5; font-size: 13px; }
        h3 { font-size: 16px; font-weight: 700; color: #111827; margin-bottom: 6px; }
        p { font-size: 12.5px; color: #4b5563; margin-bottom: 16px; }
        .form-group { margin-bottom: 12px; }
        label { display: block; font-size: 11.5px; font-weight: 700; color: #374151; margin-bottom: 5px; text-transform: uppercase; letter-spacing: 0.5px; }
        input[type="date"], select { width: 100%; padding: 9px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 13.5px; color: #111827; background: #f9fafb; outline: none; }
        input[type="date"]:focus, select:focus { border-color: #0F9D58; box-shadow: 0 0 0 2px rgba(15,157,88,0.2); }
        .hint-card { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; padding: 10px 12px; border-radius: 6px; font-size: 11.5px; margin-top: 10px; margin-bottom: 16px; line-height: 1.4; }
        .btn-container { text-align: center; }
        .btn { width: 100%; padding: 11px 20px; background: #0F9D58; color: #ffffff; font-size: 14px; font-weight: 700; border: none; border-radius: 6px; cursor: pointer; transition: background 0.15s ease; }
        .btn:hover { background: #0b8043; }
        .btn:disabled { background: #9ca3af; cursor: not-allowed; }
        .status-box { display: none; margin-top: 14px; padding: 12px; border-radius: 6px; font-size: 12px; line-height: 1.4; }
        .status-loading { display: block; background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af; text-align: center; }
        .status-success { display: block; background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
        .status-error { display: block; background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
      </style>
    </head>
    <body>
      <h3>📅 Custom Date Range Sync</h3>
      <p>Select the date range and sync condition to fetch orders from your eBay store:</p>

      <div class="form-group">
        <label for="startDate">From (Start Date)</label>
        <input type="date" id="startDate" value="${defaultStart}">
      </div>

      <div class="form-group">
        <label for="endDate">To (End Date)</label>
        <input type="date" id="endDate" value="${defaultEnd}">
      </div>

      <div class="form-group">
        <label for="syncConditionSelect">Sync Condition</label>
        <select id="syncConditionSelect">
          <option value="NEW_AND_UPDATE" ${currentMode === 'NEW_AND_UPDATE' ? 'selected' : ''}>1) Sync New Orders & Update Status</option>
          <option value="SYNC_ALL" ${currentMode === 'SYNC_ALL' ? 'selected' : ''}>2) Sync All (New + Previous Orders)</option>
        </select>
      </div>

      <div class="hint-card">
        🛡️ <b>Note:</b> eBay allows retrieving up to 90 days of order history. End date is automatically kept within current time.
      </div>

      <div class="btn-container">
        <button class="btn" id="btnSync" onclick="startCustomSync()">📥 Fetch Orders in Range</button>
      </div>

      <div id="statusBox" class="status-box"></div>

      <script>
        function startCustomSync() {
          const startVal = document.getElementById('startDate').value;
          const endVal = document.getElementById('endDate').value;
          const modeVal = document.getElementById('syncConditionSelect').value;
          const btn = document.getElementById('btnSync');
          const statusBox = document.getElementById('statusBox');

          if (!startVal || !endVal) {
            alert('Please select both Start Date and End Date.');
            return;
          }

          if (startVal > endVal) {
            alert('Start Date cannot be after End Date.');
            return;
          }

          btn.disabled = true;
          btn.innerText = '⏳ Syncing... Please wait';
          statusBox.className = 'status-box status-loading';
          statusBox.innerHTML = 'Connecting to eBay Fulfillment API and syncing orders... This may take up to 30 seconds.';

          google.script.run
            .withSuccessHandler(function(res) {
              btn.disabled = false;
              btn.innerText = '📥 Fetch Orders in Range';
              statusBox.className = 'status-box status-success';
              statusBox.innerHTML = '<b>✓ Sync Complete!</b><br>' +
                '• New Orders Imported: <b>' + res.imported + '</b><br>' +
                '• Existing Orders Updated: <b>' + res.updated + '</b><br>' +
                '• Non-Pending/Shipped Skipped: <b>' + res.skipped + '</b><br><br>' +
                '<small>You can now close this window.</small>';
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.innerText = '📥 Fetch Orders in Range';
              statusBox.className = 'status-box status-error';
              statusBox.innerHTML = '<b>❌ Error:</b> ' + (err.message || err);
            })
            .syncEbayOrdersCustomRange(startVal, endVal, modeVal);
        }
      </script>
    </body>
    </html>
  `).setWidth(460).setHeight(480);

  ui.showModalDialog(html, 'eBay Date Range Sync');
}

// ==========================================
// 2. QUICK START USER GUIDE MODAL
// ==========================================

function showUserGuideModal() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 22px; color: #1f2937; background: #ffffff; line-height: 1.5; font-size: 13px; }
        h2 { font-size: 18px; font-weight: 700; color: #111827; margin-bottom: 4px; display: flex; align-items: center; gap: 8px; }
        p.subtitle { color: #6b7280; font-size: 13px; margin-bottom: 20px; }
        .step { display: flex; gap: 14px; margin-bottom: 18px; align-items: flex-start; }
        .step-num { flex-shrink: 0; width: 28px; height: 28px; border-radius: 50%; background: #0F9D58; color: #ffffff; font-weight: 700; display: flex; align-items: center; justify-content: center; font-size: 13px; }
        .step-body h4 { font-size: 14px; font-weight: 700; color: #111827; margin-bottom: 3px; }
        .step-body p { color: #4b5563; font-size: 12.5px; line-height: 1.4; }
        .tip-card { background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 12px 14px; margin-top: 10px; color: #166534; font-size: 12px; }
        .btn-close { display: block; width: 100%; margin-top: 20px; padding: 10px; background: #111827; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; text-align: center; }
      </style>
    </head>
    <body>
      <h2>🚀 Quick Start Guide</h2>
      <p class="subtitle">Connect your eBay store and start auto-syncing in 3 simple steps:</p>

      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <h4>Authorize My eBay Store (Step 1)</h4>
          <p>Click <code>⚡ eBay Sync</code> &rarr; <code>🔑 1. Authorize My eBay Store</code>. Enter your Name, Gmail, and select your marketplace (Italy, UK, US, etc.). Click Authorize to sign in on eBay.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">2</div>
        <div class="step-body">
          <h4>Paste Redirect URL (Step 2)</h4>
          <p>After clicking "Agree and Continue" on eBay, copy the URL from your browser address bar. Return to this sheet and click <code>🔄 2. Paste Redirect Code</code>.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">3</div>
        <div class="step-body">
          <h4>Turn On Auto-Sync</h4>
          <p>Click <code>⏰ Enable 15-Minute Auto Sync</code> or <code>1-Hour Auto Sync</code>. Google Cloud will now sync your orders in the background 24/7!</p>
        </div>
      </div>

      <div class="tip-card">
        💡 <b>Tip:</b> Enter any date in Cell <b>K1</b> to instantly filter orders for that date. Delete K1 to show all orders!
      </div>

      <button class="btn-close" onclick="google.script.host.close()">Got it!</button>
    </body>
    </html>
  `).setWidth(480).setHeight(460);

  ui.showModalDialog(html, 'Quick Start Guide');
}

// ==========================================
// 3. FAQS & TROUBLESHOOTING MODAL
// ==========================================

function showFaqModal() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 22px; color: #1f2937; background: #ffffff; line-height: 1.5; font-size: 13px; }
        h2 { font-size: 18px; font-weight: 700; color: #111827; margin-bottom: 14px; }
        .faq-item { margin-bottom: 14px; border-bottom: 1px solid #f3f4f6; padding-bottom: 12px; }
        .faq-q { font-weight: 700; color: #111827; margin-bottom: 3px; font-size: 13px; }
        .faq-a { color: #4b5563; font-size: 12.5px; line-height: 1.4; }
        .btn-close { display: block; width: 100%; margin-top: 16px; padding: 10px; background: #111827; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; text-align: center; }
      </style>
    </head>
    <body>
      <h2>❓ FAQs & Troubleshooting</h2>

      <div class="faq-item">
        <div class="faq-q">Which orders get synced?</div>
        <div class="faq-a">It depends on your active condition setting under <code>⚡ eBay Sync</code> &rarr; <code>⚙️ Set Sync Condition</code>:<br>• <b>Condition 1 (New & Update):</b> Syncs new Pending and Shipped orders, and refreshes existing orders when updated.<br>• <b>Condition 2 (Sync All):</b> Syncs all orders without skipping, including past Delivered and Cancelled orders.</div>
      </div>

      <div class="faq-item">
        <div class="faq-q">Does background sync work when spreadsheet is closed?</div>
        <div class="faq-a">Yes! Once you enable 15-Minute or 1-Hour Auto Sync, Google's Cloud servers run the sync automatically 24/7 in the background without needing your browser or computer open.</div>
      </div>

      <div class="faq-item">
        <div class="faq-q">How does the Cell K1 Date Filter work?</div>
        <div class="faq-a">Simply type or paste any date into Cell K1 (e.g. <code>2026-11-01</code>, or <code>2026-10</code> for the entire month). The sheet will immediately filter Column C. Clear Cell K1 to reveal all orders again.</div>
      </div>

      <div class="faq-item">
        <div class="faq-q">How does currency formatting work?</div>
        <div class="faq-a">The sheet automatically formats all prices, fees, and earnings in the native currency of each order (€ for Italy/EU, £ for UK, $ for US, A$ for Australia, etc.). You can also click <code>⚡ eBay Sync</code> &rarr; <code>💱 Update All Currency Formats</code> to update historical rows.</div>
      </div>

      <div class="faq-item">
        <div class="faq-q">Are latest orders always on top?</div>
        <div class="faq-a">Yes! All orders are sorted chronologically descending so the newest sale is always at Row 6. You can also re-sort anytime using <code>🔃 Sort All Orders (Latest on Top)</code>.</div>
      </div>

      <div class="faq-item">
        <div class="faq-q">How do I change my registered subscription Gmail?</div>
        <div class="faq-a">Click <code>⚡ eBay Sync</code> &rarr; <code>🛡️ Subscription & License Status</code> &rarr; <code>✏️ Change / Update Registered Gmail</code>.</div>
      </div>

      <button class="btn-close" onclick="google.script.host.close()">Close</button>
    </body>
    </html>
  `).setWidth(500).setHeight(520);

  ui.showModalDialog(html, 'FAQs & Troubleshooting');
}

