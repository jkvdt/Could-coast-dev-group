// ─── Update the monthly budget amount for a category ─────────────────────────
function setBudget(category, subcategory, monthIndex, amount) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.BUDGET);
  if (!sheet) return false;

  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === category && data[i][1] === subcategory) {
      const budCol = BUDGET_HEADER_COLS + monthIndex * BUDGET_COLS_PER_MONTH + 1;
      sheet.getRange(i + 2, budCol).setValue(amount);
      // Update annual budget (sum of 12 monthly budgets)
      const annBudCol = BUDGET_ANNUAL_START_COL;
      const monthlys  = Array.from({length: 12}, (_, m) =>
        sheet.getRange(i + 2, BUDGET_HEADER_COLS + m * BUDGET_COLS_PER_MONTH + 1).getValue()
      );
      sheet.getRange(i + 2, annBudCol).setValue(monthlys.reduce((a,b) => a + b, 0));
      return true;
    }
  }
  return false;
}

// Returns a summary object with totals for the given month (0-based index).
function getMonthSummary(monthIndex) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const budSheet = ss.getSheetByName(SHEET.BUDGET);
  if (!budSheet || budSheet.getLastRow() < 2) return null;

  const budCol = BUDGET_HEADER_COLS + monthIndex * BUDGET_COLS_PER_MONTH + 1;
  const actCol = budCol + 1;
  const data   = budSheet.getRange(2, 1, budSheet.getLastRow() - 1, actCol).getValues();

  let totalIncomeBud = 0, totalIncomeAct = 0;
  let totalExpBud    = 0, totalExpAct    = 0;

  for (const row of data) {
    const type = row[2];
    const bud  = row[budCol - 1] || 0;
    const act  = row[actCol - 1] || 0;
    if (type === 'Income') { totalIncomeBud += bud; totalIncomeAct += act; }
    else                   { totalExpBud    += bud; totalExpAct    += act; }
  }

  return {
    month:       monthIndex,
    incomeBud:   totalIncomeBud,
    incomeAct:   totalIncomeAct,
    expenseBud:  totalExpBud,
    expenseAct:  totalExpAct,
    netBud:      totalIncomeBud - totalExpBud,
    netAct:      totalIncomeAct - totalExpAct,
    savingsRate: totalIncomeAct > 0 ? (totalIncomeAct - totalExpAct) / totalIncomeAct : 0,
  };
}
