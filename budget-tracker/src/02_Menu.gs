function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Budget Tracker')
    .addItem('🔄 Refresh Dashboard',        'refreshDashboard')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📥 Import Transactions')
      .addItem('Import from CSV file…',     'showCsvImportDialog')
      .addItem('Add transaction manually…', 'showAddTransactionDialog'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🏦 Live Bank Sync (Plaid)')
      .addItem('Connect a bank account…',   'showPlaidConnect')
      .addItem('Sync transactions now',     'syncAllPlaidAccounts')
      .addItem('View connected accounts',   'showConnectedAccounts'))
    .addSeparator()
    .addItem('📊 Regenerate Annual Summary', 'buildAnnualSummary')
    .addItem('🏷️  Re-categorize transactions','runAutoCategorizer')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️  Settings')
      .addItem('Set Plaid API credentials…','showPlaidSettings')
      .addItem('Set budget year…',          'showYearSetting')
      .addItem('Reset / Reinstall sheets',  'confirmReinstall'))
    .addToUi();
}

function showAddTransactionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddTransactionDialog')
    .setWidth(480).setHeight(400).setTitle('Add Transaction');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Transaction');
}

function showYearSetting() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty('BUDGET_YEAR') || new Date().getFullYear();
  const result = ui.prompt('Budget Year', `Enter the budget year (currently ${current}):`, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt(result.getResponseText(), 10);
  if (isNaN(year) || year < 2000 || year > 2100) {
    ui.alert('Invalid year. Please enter a 4-digit year.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('BUDGET_YEAR', String(year));
  ui.alert(`Budget year set to ${year}. Run "Regenerate Annual Summary" to update.`);
}

function getBudgetYear() {
  return parseInt(PropertiesService.getScriptProperties().getProperty('BUDGET_YEAR') || new Date().getFullYear(), 10);
}

function confirmReinstall() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Reinstall Sheets',
    'This will recreate all sheets with default structure. Existing transaction data will NOT be deleted. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.YES) setupSpreadsheet();
}
