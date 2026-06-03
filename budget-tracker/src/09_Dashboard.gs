function refreshDashboard() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  let dash       = ss.getSheetByName(SHEET.DASHBOARD);
  if (!dash) dash = ss.insertSheet(SHEET.DASHBOARD, 0);
  dash.clearContents().clearFormats();

  const year     = getBudgetYear();
  const now      = new Date();
  const month    = now.getMonth(); // 0-based
  const monthStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM');
  const monthName = MONTHS[month];
  const currency = _getSetting(ss, 'Currency Symbol') || '$';

  // ── Title ──────────────────────────────────────────────────────────────────
  dash.getRange('A1:J1').merge()
    .setValue(`💰 Personal Budget Tracker  ·  ${monthName} ${year}`)
    .setFontSize(18).setFontWeight('bold')
    .setBackground(COLOR.HEADER_BG).setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 50);

  dash.getRange('A2').setValue(`Last updated: ${Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMM d, yyyy h:mm a')}`);
  dash.getRange('A2:J2').setFontColor('#888888').setFontSize(10);

  // ── Account balances ───────────────────────────────────────────────────────
  _dashSection(dash, 4, 1, 'ACCOUNT BALANCES');
  const acctSheet = ss.getSheetByName(SHEET.ACCOUNTS);
  if (acctSheet && acctSheet.getLastRow() >= 2) {
    const accts = acctSheet.getRange(2, 1, acctSheet.getLastRow() - 1, 5).getValues()
      .filter(r => r[0]);
    dash.getRange(5, 1).setValue('Account').setFontWeight('bold');
    dash.getRange(5, 2).setValue('Type').setFontWeight('bold');
    dash.getRange(5, 3).setValue('Balance').setFontWeight('bold');
    let totalBal = 0;
    for (let i = 0; i < accts.length; i++) {
      dash.getRange(6 + i, 1).setValue(accts[i][0]);
      dash.getRange(6 + i, 2).setValue(accts[i][2]);
      dash.getRange(6 + i, 3).setValue(accts[i][4]).setNumberFormat(`"${currency}"#,##0.00`);
      totalBal += Number(accts[i][4]) || 0;
    }
    const totRow = 6 + accts.length;
    dash.getRange(totRow, 1).setValue('Total').setFontWeight('bold');
    dash.getRange(totRow, 3).setValue(totalBal).setNumberFormat(`"${currency}"#,##0.00`).setFontWeight('bold');
  } else {
    dash.getRange(5, 1).setValue('No accounts connected. Use Budget Tracker → Live Bank Sync to connect.')
      .setFontColor('#888888').setFontStyle('italic');
  }

  // ── This month: budget vs actual ───────────────────────────────────────────
  const budSection = 4;
  const budCol     = 5;
  _dashSection(dash, budSection, budCol, `${monthName.toUpperCase()} BUDGET vs ACTUAL`);

  const budSheet = ss.getSheetByName(SHEET.BUDGET);
  if (budSheet && budSheet.getLastRow() >= 2) {
    const monthBudActData = _getMonthBudAct(budSheet, month);
    let r = budSection + 1;
    dash.getRange(r, budCol).setValue('Category').setFontWeight('bold');
    dash.getRange(r, budCol + 1).setValue('Budget').setFontWeight('bold');
    dash.getRange(r, budCol + 2).setValue('Spent').setFontWeight('bold');
    dash.getRange(r, budCol + 3).setValue('Remaining').setFontWeight('bold');
    r++;

    // Group by category for display
    const byCategory = {};
    for (const row of monthBudActData) {
      const [cat, , type, bud, act] = row;
      if (type === 'Income') continue;
      if (!byCategory[cat]) byCategory[cat] = { bud: 0, act: 0 };
      byCategory[cat].bud += bud;
      byCategory[cat].act += act;
    }

    let totalBud = 0, totalAct = 0, rowIdx = 0;
    for (const [cat, vals] of Object.entries(byCategory)) {
      if (vals.bud === 0 && vals.act === 0) continue;
      const remaining = vals.bud - vals.act;
      dash.getRange(r, budCol).setValue(cat);
      dash.getRange(r, budCol + 1).setValue(vals.bud).setNumberFormat(`"${currency}"#,##0.00`);
      dash.getRange(r, budCol + 2).setValue(vals.act).setNumberFormat(`"${currency}"#,##0.00`);
      dash.getRange(r, budCol + 3).setValue(remaining).setNumberFormat(`"${currency}"#,##0.00`);
      if (remaining < 0) {
        dash.getRange(r, budCol + 3).setFontColor(COLOR.OVER_BUDGET);
        dash.getRange(r, budCol).setFontColor(COLOR.OVER_BUDGET);
      }
      if (rowIdx % 2 === 0) dash.getRange(r, budCol, 1, 4).setBackground(COLOR.ALT_ROW);
      totalBud += vals.bud;
      totalAct += vals.act;
      r++;
      rowIdx++;
    }
    // Totals
    dash.getRange(r, budCol).setValue('TOTAL').setFontWeight('bold');
    dash.getRange(r, budCol + 1).setValue(totalBud).setNumberFormat(`"${currency}"#,##0.00`).setFontWeight('bold');
    dash.getRange(r, budCol + 2).setValue(totalAct).setNumberFormat(`"${currency}"#,##0.00`).setFontWeight('bold');
    const totalRem = totalBud - totalAct;
    dash.getRange(r, budCol + 3).setValue(totalRem).setNumberFormat(`"${currency}"#,##0.00`).setFontWeight('bold');
    if (totalRem < 0) dash.getRange(r, budCol + 3).setFontColor(COLOR.OVER_BUDGET);
    else dash.getRange(r, budCol + 3).setFontColor(COLOR.UNDER_BUDGET);
  }

  // ── Recent transactions ────────────────────────────────────────────────────
  const recentRow = dash.getLastRow() + 3;
  _dashSection(dash, recentRow, 1, 'RECENT TRANSACTIONS (last 15)');
  const txSheet = ss.getSheetByName(SHEET.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const hdr = ['Date', 'Account', 'Description', 'Amount', 'Category'];
    dash.getRange(recentRow + 1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold');
    const allTx = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues()
      .filter(r => r[0])
      .sort((a, b) => new Date(b[0]) - new Date(a[0]))
      .slice(0, 15);
    if (allTx.length > 0) {
      dash.getRange(recentRow + 2, 1, allTx.length, 5).setValues(allTx);
      dash.getRange(recentRow + 2, 1, allTx.length, 1).setNumberFormat('mm/dd/yyyy');
      dash.getRange(recentRow + 2, 4, allTx.length, 1).setNumberFormat(`"${currency}"#,##0.00;[RED]"${currency}"(#,##0.00)`);
    }
  }

  // ── Column widths ──────────────────────────────────────────────────────────
  dash.setColumnWidths(1, 10, [160, 120, 120, 120, 160, 120, 120, 120, 120, 120]);
  dash.setFrozenRows(1);
  dash.setTabColor('#1a1a2e');
}

function _dashSection(sheet, row, col, title) {
  sheet.getRange(row, col, 1, 4).merge()
    .setValue(title)
    .setBackground(COLOR.SECTION_BG)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#1a1a2e');
  sheet.setRowHeight(row, 26);
}

function _getMonthBudAct(budSheet, monthIndex) {
  const lastRow = budSheet.getLastRow();
  if (lastRow < 2) return [];
  const budCol = BUDGET_HEADER_COLS + monthIndex * BUDGET_COLS_PER_MONTH + 1;
  const actCol = budCol + 1;
  const data   = budSheet.getRange(2, 1, lastRow - 1, actCol).getValues();
  return data.map(r => [r[0], r[1], r[2], r[budCol - 1] || 0, r[actCol - 1] || 0]);
}

function _getSetting(ss, key) {
  const sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) return null;
  const data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const row   = data.find(r => r[0] === key);
  return row ? row[1] : null;
}
