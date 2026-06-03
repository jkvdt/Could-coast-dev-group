// ─── Plaid credential helpers ─────────────────────────────────────────────────
function getPlaidConfig() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET.SETTINGS);
  const props  = PropertiesService.getScriptProperties();

  // Prefer encrypted script properties over the Settings sheet (safer)
  const clientId = props.getProperty('PLAID_CLIENT_ID') ||
    (settingsSheet ? settingsSheet.getRange('B5').getValue() : '');
  const secret   = props.getProperty('PLAID_SECRET') ||
    (settingsSheet ? settingsSheet.getRange('B6').getValue() : '');
  const env      = props.getProperty('PLAID_ENV') ||
    (settingsSheet ? settingsSheet.getRange('B3').getValue() : 'production');

  return { clientId, secret, baseUrl: `https://${env}.plaid.com` };
}

function showPlaidSettings() {
  const ui   = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const idResult = ui.prompt('Plaid Client ID',
    'Enter your Plaid client_id (from dashboard.plaid.com → Team Settings → Keys):',
    ui.ButtonSet.OK_CANCEL);
  if (idResult.getSelectedButton() !== ui.Button.OK) return;

  const secretResult = ui.prompt('Plaid Secret',
    'Enter your Plaid secret (production or sandbox):',
    ui.ButtonSet.OK_CANCEL);
  if (secretResult.getSelectedButton() !== ui.Button.OK) return;

  const envResult = ui.prompt('Plaid Environment',
    'Enter environment — type "sandbox" (for testing) or "production":',
    ui.ButtonSet.OK_CANCEL);
  if (envResult.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty('PLAID_CLIENT_ID', idResult.getResponseText().trim());
  props.setProperty('PLAID_SECRET',    secretResult.getResponseText().trim());
  props.setProperty('PLAID_ENV',       envResult.getResponseText().trim() || 'production');

  ui.alert('✅ Plaid credentials saved securely.\nNow use "Connect a bank account…" to link your first account.');
}

// ─── Link Token (creates the token Plaid Link needs to open) ──────────────────
function createLinkToken() {
  const cfg = getPlaidConfig();
  if (!cfg.clientId || !cfg.secret) {
    return { error: 'Plaid credentials not set. Use Settings → Set Plaid API credentials first.' };
  }

  const resp = _plaidPost(cfg, '/link/token/create', {
    user: { client_user_id: Session.getEffectiveUser().getEmail() || 'budget-tracker-user' },
    client_name: 'Google Sheets Budget Tracker',
    products: ['transactions'],
    country_codes: ['US'],
    language: 'en',
  });

  if (resp.error_code) return { error: `Plaid error: ${resp.error_message}` };
  return { linkToken: resp.link_token };
}

// ─── Exchange public token → access token, save account ──────────────────────
function exchangePublicToken(publicToken, metadata) {
  const cfg = getPlaidConfig();
  const resp = _plaidPost(cfg, '/item/public_token/exchange', {
    public_token: publicToken,
  });

  if (resp.error_code) return { error: `Plaid error: ${resp.error_message}` };

  const accessToken = resp.access_token;
  const itemId      = resp.item_id;

  // Store access token securely in Script Properties (never in the sheet itself)
  const propKey = `PLAID_ACCESS_${itemId}`;
  PropertiesService.getScriptProperties().setProperty(propKey, accessToken);

  // Record account info in Accounts sheet
  _saveAccountsFromMetadata(metadata, itemId, propKey);

  // Immediately pull 90 days of transactions
  syncPlaidItem(accessToken, itemId);

  return { ok: true, itemId };
}

function _saveAccountsFromMetadata(metadata, itemId, propKey) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ACCOUNTS);
  const inst  = metadata.institution ? metadata.institution.name : 'Unknown';

  (metadata.accounts || []).forEach(acct => {
    sheet.appendRow([
      acct.name || acct.official_name || '',
      inst,
      acct.subtype || acct.type || '',
      acct.mask || '',
      '',
      new Date(),
      itemId,
      propKey, // key into Script Properties — not the token itself
    ]);
  });
}

// ─── Sync one item ────────────────────────────────────────────────────────────
function syncPlaidItem(accessToken, itemId) {
  const cfg    = getPlaidConfig();
  const props  = PropertiesService.getScriptProperties();
  const cursorKey = `PLAID_CURSOR_${itemId}`;
  let cursor   = props.getProperty(cursorKey) || '';

  const txRows = [];
  let hasMore  = true;

  while (hasMore) {
    const body = { access_token: accessToken };
    if (cursor) body.cursor = cursor;

    const resp = _plaidPost(cfg, '/transactions/sync', body);
    if (resp.error_code) throw new Error(`Plaid sync error: ${resp.error_message}`);

    for (const tx of (resp.added || [])) txRows.push(_plaidTxToRow(tx));
    // We ignore 'modified' and 'removed' for now (dedup by ID handles most cases)

    cursor  = resp.next_cursor || cursor;
    hasMore = resp.has_more;
  }

  props.setProperty(cursorKey, cursor);
  const count = insertTransactions(txRows);

  // Update balances
  _updateBalances(cfg, accessToken, itemId);
  return count;
}

function _plaidTxToRow(tx) {
  const date   = new Date(tx.date);
  const name   = tx.merchant_name || tx.name || '';
  const amount = -(tx.amount); // Plaid: positive = outflow (expense), we negate
  const [cat, sub] = categorize(name);
  return [date, tx.account_id, name, amount, cat, sub, '', tx.transaction_id, 'Plaid', ''];
}

function _updateBalances(cfg, accessToken, itemId) {
  const resp = _plaidPost(cfg, '/accounts/balance/get', { access_token: accessToken });
  if (resp.error_code) return;

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET.ACCOUNTS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids      = sheet.getRange(2, 7, lastRow - 1, 1).getValues(); // itemId col
  const now      = new Date();

  for (const acct of (resp.accounts || [])) {
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === itemId) {
        const balance = acct.balances.current ?? acct.balances.available ?? '';
        sheet.getRange(i + 2, 5).setValue(balance);
        sheet.getRange(i + 2, 6).setValue(now);
        break;
      }
    }
  }
}

// ─── Sync all connected accounts ──────────────────────────────────────────────
function syncAllPlaidAccounts() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET.ACCOUNTS);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No connected accounts. Use "Connect a bank account…" first.');
    return;
  }

  const props   = PropertiesService.getScriptProperties();
  const rows    = sheet.getRange(2, 7, sheet.getLastRow() - 1, 2).getValues(); // itemId, propKey
  let totalNew  = 0;

  for (const [itemId, propKey] of rows) {
    if (!itemId || !propKey) continue;
    const accessToken = props.getProperty(propKey);
    if (!accessToken) continue;
    try {
      totalNew += syncPlaidItem(accessToken, itemId);
    } catch(e) {
      Logger.log(`Sync failed for item ${itemId}: ${e.message}`);
    }
  }

  refreshDashboard();
  SpreadsheetApp.getUi().alert(`Sync complete. ${totalNew} new transaction(s) added.`);
}

function showConnectedAccounts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ACCOUNTS);
  ss.setActiveSheet(sheet);
}

// ─── Plaid Link sidebar ────────────────────────────────────────────────────────
function showPlaidConnect() {
  const result = createLinkToken();
  if (result.error) {
    SpreadsheetApp.getUi().alert(`Error: ${result.error}`);
    return;
  }
  const html = HtmlService.createTemplateFromFile('PlaidSidebar');
  html.linkToken = result.linkToken;
  html.plaidEnv  = (getPlaidConfig().baseUrl.includes('sandbox') ? 'sandbox' : 'production');
  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Connect Bank Account').setWidth(400));
}

// ─── HTTP helper ─────────────────────────────────────────────────────────────
function _plaidPost(cfg, path, body) {
  const payload = Object.assign({ client_id: cfg.clientId, secret: cfg.secret }, body);
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  const resp = UrlFetchApp.fetch(cfg.baseUrl + path, options);
  return JSON.parse(resp.getContentText());
}
