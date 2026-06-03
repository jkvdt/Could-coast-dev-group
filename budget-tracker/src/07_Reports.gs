// Rebuilds the Annual Summary sheet — a 12-month grid with income vs expenses.
function buildAnnualSummary() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET.ANNUAL);
  if (!sheet) sheet = ss.insertSheet(SHEET.ANNUAL);
  sheet.clearContents().clearFormats();

  const year    = getBudgetYear();
  const budSheet = ss.getSheetByName(SHEET.BUDGET);
  if (!budSheet) return;

  const headers = ['Category', 'Subcategory', 'Type', 'Annual Budget', 'Annual Actual', 'Variance', ...MONTHS];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  _styleHeaderRow(sheet, 1, headers.length);

  // Pull all budget rows
  const lastBudRow = budSheet.getLastRow();
  if (lastBudRow < 2) return;

  const budData = budSheet.getRange(2, 1, lastBudRow - 1, BUDGET_ANNUAL_START_COL + 2).getValues();
  const annData = [];

  for (const row of budData) {
    const cat    = row[0];
    const sub    = row[1];
    const type   = row[2];
    if (!cat) continue;

    const annBud = row[BUDGET_ANNUAL_START_COL - 1];
    const annAct = row[BUDGET_ANNUAL_START_COL];
    const annVar = row[BUDGET_ANNUAL_START_COL + 1];

    const monthActuals = [];
    for (let m = 0; m < 12; m++) {
      const actCol = BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 2; // 1-based → 0-based
      monthActuals.push(row[actCol - 1]);
    }

    annData.push([cat, sub, type, annBud, annAct, annVar, ...monthActuals]);
  }

  if (annData.length === 0) return;
  sheet.getRange(2, 1, annData.length, headers.length).setValues(annData);

  // Formatting
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.setColumnWidths(1, 3, [160, 180, 75]);
  sheet.getRange(2, 4, annData.length, headers.length - 3).setNumberFormat('"$"#,##0.00');

  // Color income rows green, expense rows light red
  for (let i = 0; i < annData.length; i++) {
    const type = annData[i][2];
    const color = type === 'Income' ? COLOR.INCOME_BG : (i % 2 === 0 ? '#ffffff' : COLOR.ALT_ROW);
    sheet.getRange(i + 2, 1, 1, headers.length).setBackground(color);
  }

  // Totals row
  const totalRow = sheet.getLastRow() + 2;
  sheet.getRange(totalRow, 1).setValue('NET (Income − Expenses)').setFontWeight('bold');
  // Net = sum income actuals − sum expense actuals for each month
  for (let m = 0; m < 12; m++) {
    const colLetter = _colLetter(7 + m);
    const incomeSum = `SUMIF($C$2:$C${totalRow-2},"Income",${colLetter}$2:${colLetter}${totalRow-2})`;
    const expSum    = `SUMIF($C$2:$C${totalRow-2},"Expense",${colLetter}$2:${colLetter}${totalRow-2})`;
    sheet.getRange(totalRow, 7 + m).setFormula(`=${incomeSum}-${expSum}`);
  }
  // Annual net
  sheet.getRange(totalRow, 4).setFormula(
    `=SUMIF($C$2:$C${totalRow-2},"Income",$D$2:$D${totalRow-2})-SUMIF($C$2:$C${totalRow-2},"Expense",$D$2:$D${totalRow-2})`);
  sheet.getRange(totalRow, 5).setFormula(
    `=SUMIF($C$2:$C${totalRow-2},"Income",$E$2:$E${totalRow-2})-SUMIF($C$2:$C${totalRow-2},"Expense",$E$2:$E${totalRow-2})`);
  sheet.getRange(totalRow, 6).setFormula(`=D${totalRow}-E${totalRow}`);

  sheet.getRange(totalRow, 4, 1, headers.length - 3)
    .setNumberFormat('"$"#,##0.00')
    .setFontWeight('bold')
    .setBackground(COLOR.SECTION_BG);

  SpreadsheetApp.getUi().alert('Annual Summary updated.');
}
