function categorize(description) {
  const d = (description || '').trim();
  for (const [pattern, cat] of CATEGORY_RULES) {
    if (pattern.test(d)) return cat;
  }
  return ['Uncategorized', 'Uncategorized'];
}

// Runs over every Uncategorized row in Transactions and attempts to fill it in.
function runAutoCategorizer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.TRANSACTIONS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, TX_COL.MONTH);
  const data = range.getValues();
  let changed = 0;

  for (let i = 0; i < data.length; i++) {
    const cat    = data[i][TX_COL.CATEGORY - 1];
    const subcat = data[i][TX_COL.SUBCAT - 1];
    if (cat && cat !== 'Uncategorized') continue;

    const name = data[i][TX_COL.NAME - 1];
    const [newCat, newSub] = categorize(name);
    data[i][TX_COL.CATEGORY - 1] = newCat;
    data[i][TX_COL.SUBCAT - 1]   = newSub;
    changed++;
  }

  if (changed > 0) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert(`Updated ${changed} transaction(s).`);
  } else {
    SpreadsheetApp.getUi().alert('No uncategorized transactions found.');
  }
}

// Returns sorted unique [Category, Subcategory] pairs from the Categories sheet.
function getCategoryList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.CATEGORIES);
  if (!sheet || sheet.getLastRow() < 2) return DEFAULT_CATEGORIES.map(r => [r[0], r[1]]);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().filter(r => r[0]);
}
