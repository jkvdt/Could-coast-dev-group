// Entry point — run once to build the whole spreadsheet.
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupSettings(ss);
  _setupCategories(ss);
  _setupTransactions(ss);
  _setupBudget(ss);
  _setupAnnual(ss);
  _setupAccounts(ss);
  _setupDashboard(ss);
  refreshDashboard();
  SpreadsheetApp.getUi().alert('✅ Budget Tracker is ready!\n\nNext steps:\n1. Update your monthly budgets in the "Budget" tab.\n2. Import transactions via the Budget Tracker menu.\n3. (Optional) Connect live bank accounts via Plaid.');
}

// ─── Settings sheet ────────────────────────────────────────────────────────────
function _setupSettings(ss) {
  let sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEET.SETTINGS);
  sheet.clearContents();

  const rows = [
    ['Setting', 'Value', 'Notes'],
    ['Budget Year', new Date().getFullYear(), 'Year to track'],
    ['Currency Symbol', '$', 'Shown on dashboard'],
    ['Plaid Environment', 'production', 'production or sandbox'],
    ['Plaid Client ID', '', 'From dashboard.plaid.com'],
    ['Plaid Secret', '', 'Keep this private'],
    ['Auto-sync on open', 'FALSE', 'TRUE to sync Plaid on every open'],
    ['Starting Net Worth', 0, 'Optional baseline for net worth chart'],
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  _styleHeaderRow(sheet, 1, 3);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  sheet.getRange('B5:B6').setFontWeight('bold').setBackground('#fff3cd'); // highlight creds
}

// ─── Categories sheet ──────────────────────────────────────────────────────────
function _setupCategories(ss) {
  let sheet = ss.getSheetByName(SHEET.CATEGORIES);
  if (!sheet) sheet = ss.insertSheet(SHEET.CATEGORIES);
  sheet.clearContents();

  const headers = [['Category', 'Subcategory', 'Type', 'Monthly Budget']];
  const data = DEFAULT_CATEGORIES.map(r => [r[0], r[1], r[2], r[3]]);
  sheet.getRange(1, 1, 1, 4).setValues(headers);
  sheet.getRange(2, 1, data.length, 4).setValues(data);

  _styleHeaderRow(sheet, 1, 4);
  sheet.setColumnWidths(1, 4, [180, 200, 80, 130]);
  sheet.getRange(2, 4, data.length, 1).setNumberFormat('"$"#,##0.00');

  // Alternate row shading
  for (let i = 0; i < data.length; i++) {
    if (i % 2 === 0) sheet.getRange(i + 2, 1, 1, 4).setBackground(COLOR.ALT_ROW);
  }
}

// ─── Transactions sheet ────────────────────────────────────────────────────────
function _setupTransactions(ss) {
  let sheet = ss.getSheetByName(SHEET.TRANSACTIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.TRANSACTIONS);
  } else if (sheet.getLastRow() >= 1) {
    // Sheet already has data — only fix headers/format, don't wipe rows
    _applyTransactionFormat(sheet);
    return;
  }

  const headers = [['Date','Account','Description','Amount','Category','Subcategory','Notes','Transaction ID','Source','Month']];
  sheet.getRange(1, 1, 1, 10).setValues(headers);
  _applyTransactionFormat(sheet);
}

function _applyTransactionFormat(sheet) {
  _styleHeaderRow(sheet, 1, 10);
  sheet.setColumnWidths(1, 10, [100, 150, 260, 90, 150, 150, 180, 160, 80, 80]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Date format
  sheet.getRange(2, TX_COL.DATE, sheet.getMaxRows() - 1, 1).setNumberFormat('mm/dd/yyyy');
  // Amount format
  sheet.getRange(2, TX_COL.AMOUNT, sheet.getMaxRows() - 1, 1).setNumberFormat('"$"#,##0.00;[RED]"$"(#,##0.00)');
  // Month formula locked to column A
  // We insert this as a note rather than hardcoded formula; rows are populated by import/Plaid scripts.

  // Data validation: Category dropdown from Categories sheet
  const catRange = sheet.getRange(2, TX_COL.CATEGORY, sheet.getMaxRows() - 1, 1);
  const catRule  = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CATEGORIES)
        .getRange('A2:A1000'), true)
    .build();
  catRange.setDataValidation(catRule);
}

// ─── Budget sheet ──────────────────────────────────────────────────────────────
function _setupBudget(ss) {
  let sheet = ss.getSheetByName(SHEET.BUDGET);
  if (!sheet) sheet = ss.insertSheet(SHEET.BUDGET);
  sheet.clearContents().clearFormats();

  const year = getBudgetYear();

  // Build header row
  const headerRow = ['Category', 'Subcategory', 'Type'];
  MONTHS.forEach(m => headerRow.push(`${m} Budget`, `${m} Actual`, `${m} Var`));
  headerRow.push('Annual Budget', 'Annual Actual', 'Annual Var');
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  _styleHeaderRow(sheet, 1, headerRow.length);

  // Populate categories
  const catData = DEFAULT_CATEGORIES.map(r => {
    const row = [r[0], r[1], r[2]];
    MONTHS.forEach(() => { row.push(r[3], 0, 0); }); // budget, actual placeholder, variance
    row.push(r[3] * 12, 0, 0); // annual
    return row;
  });
  sheet.getRange(2, 1, catData.length, headerRow.length).setValues(catData);

  // Replace actual/variance placeholder values with formulas
  _injectBudgetFormulas(sheet, year, catData.length);

  // Formatting
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.setColumnWidths(1, 3, [160, 180, 75]);
  for (let m = 0; m < 12; m++) {
    const base = BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 1;
    sheet.setColumnWidth(base, 100);     // Budget
    sheet.setColumnWidth(base + 1, 100); // Actual
    sheet.setColumnWidth(base + 2, 90);  // Variance
  }

  // Number format for all dollar columns
  const dollarRange = sheet.getRange(2, 4, catData.length, headerRow.length - 3);
  dollarRange.setNumberFormat('"$"#,##0.00');

  // Color variance columns: red if negative
  for (let m = 0; m < 13; m++) { // 12 months + annual
    const varCol = m < 12
      ? BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 3
      : BUDGET_ANNUAL_START_COL + 2;
    const varRange = sheet.getRange(2, varCol, catData.length, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor(COLOR.OVER_BUDGET)
      .setRanges([varRange]).build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }
}

function _injectBudgetFormulas(sheet, year, numCats) {
  const txSheet = `'${SHEET.TRANSACTIONS}'`;

  for (let row = 2; row <= numCats + 1; row++) {
    const catRef  = `$A${row}`;
    const subRef  = `$B${row}`;
    const typeRef = `$C${row}`;

    for (let m = 0; m < 12; m++) {
      const monthNum = m + 1;
      const monthStr = `"${year}-${String(monthNum).padStart(2,'0')}"`;
      const budCol  = BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 1;
      const actCol  = budCol + 1;
      const varCol  = budCol + 2;

      // Actual = SUMIFS on Transactions by category, subcategory, month
      // Income: sum positive amounts; Expenses: sum absolute of negative amounts
      const sumFormula =
        `IF(${typeRef}="Income",` +
          `SUMIFS(${txSheet}!$D:$D,${txSheet}!$E:$E,${catRef},${txSheet}!$F:$F,${subRef},${txSheet}!$J:$J,${monthStr}),` +
          `-SUMIFS(${txSheet}!$D:$D,${txSheet}!$E:$E,${catRef},${txSheet}!$F:$F,${subRef},${txSheet}!$J:$J,${monthStr}))`;

      sheet.getRange(row, actCol).setFormula(`=${sumFormula}`);
      // Variance = Budget - Actual (positive = under budget = good)
      sheet.getRange(row, varCol).setFormula(
        `=${_colLetter(budCol)}${row}-${_colLetter(actCol)}${row}`);
    }

    // Annual actuals & variance
    const annActCol = BUDGET_ANNUAL_START_COL + 1;
    const annVarCol = BUDGET_ANNUAL_START_COL + 2;
    const annBudCol = BUDGET_ANNUAL_START_COL;

    // Sum all 12 monthly actuals
    const actCols = Array.from({length:12}, (_,m) =>
      `${_colLetter(BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 2)}${row}`
    ).join('+');
    sheet.getRange(row, annActCol).setFormula(`=${actCols}`);
    sheet.getRange(row, annVarCol).setFormula(
      `=${_colLetter(annBudCol)}${row}-${_colLetter(annActCol)}${row}`);
  }
}

// ─── Annual Summary sheet ──────────────────────────────────────────────────────
function _setupAnnual(ss) {
  let sheet = ss.getSheetByName(SHEET.ANNUAL);
  if (!sheet) sheet = ss.insertSheet(SHEET.ANNUAL);
  buildAnnualSummary();
}

// ─── Accounts sheet ────────────────────────────────────────────────────────────
function _setupAccounts(ss) {
  let sheet = ss.getSheetByName(SHEET.ACCOUNTS);
  if (!sheet) sheet = ss.insertSheet(SHEET.ACCOUNTS);
  sheet.clearContents();

  const headers = [['Account Name', 'Institution', 'Type', 'Last 4', 'Current Balance', 'Last Synced', 'Plaid Item ID', 'Plaid Access Token Key']];
  sheet.getRange(1, 1, 1, 8).setValues(headers);
  _styleHeaderRow(sheet, 1, 8);
  sheet.setColumnWidths(1, 8, [160, 160, 100, 70, 130, 140, 180, 220]);
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setNumberFormat('"$"#,##0.00');
  // Hide the access token key column — it stores a property key, not the token itself
  sheet.hideColumns(8);
}

// ─── Dashboard sheet ───────────────────────────────────────────────────────────
function _setupDashboard(ss) {
  let sheet = ss.getSheetByName(SHEET.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEET.DASHBOARD, 0);
  sheet.clearContents().clearFormats();
  sheet.setTabColor('#1a1a2e');
  refreshDashboard(); // will be called again after setup, that's fine
}

// ─── Utilities ─────────────────────────────────────────────────────────────────
function _styleHeaderRow(sheet, row, numCols) {
  const range = sheet.getRange(row, 1, 1, numCols);
  range.setBackground(COLOR.HEADER_BG)
       .setFontColor(COLOR.HEADER_FG)
       .setFontWeight('bold')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 28);
}

function _colLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
