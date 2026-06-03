// ─── Add a single transaction (called from dialog) ────────────────────────────
function addTransaction(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.TRANSACTIONS);
  const date  = new Date(data.date);
  const month = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM');

  sheet.appendRow([
    date,
    data.account    || '',
    data.name       || '',
    parseFloat(data.amount) || 0,
    data.category   || 'Uncategorized',
    data.subcat     || 'Uncategorized',
    data.notes      || '',
    data.txId       || `manual-${Date.now()}`,
    data.source     || 'Manual',
    month,
  ]);
  return true;
}

// ─── Bulk insert rows (used by CSV import & Plaid sync) ───────────────────────
function insertTransactions(rows) {
  if (!rows || rows.length === 0) return 0;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.TRANSACTIONS);
  const tz    = ss.getSpreadsheetTimeZone();

  // Build set of existing transaction IDs for deduplication
  const lastRow = sheet.getLastRow();
  const existingIds = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, TX_COL.TX_ID, lastRow - 1, 1)
      .getValues().forEach(r => { if (r[0]) existingIds.add(r[0]); });
  }

  const toAppend = [];
  for (const row of rows) {
    const txId = row[TX_COL.TX_ID - 1];
    if (txId && existingIds.has(txId)) continue; // skip duplicates
    existingIds.add(txId);

    // Compute Month column if not provided
    if (!row[TX_COL.MONTH - 1] && row[TX_COL.DATE - 1]) {
      const d = row[TX_COL.DATE - 1] instanceof Date
        ? row[TX_COL.DATE - 1]
        : new Date(row[TX_COL.DATE - 1]);
      row[TX_COL.MONTH - 1] = Utilities.formatDate(d, tz, 'yyyy-MM');
    }
    toAppend.push(row);
  }

  if (toAppend.length === 0) return 0;
  sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 10).setValues(toAppend);
  return toAppend.length;
}

// ─── Client-callable wrapper for the Add Transaction dialog ───────────────────
function clientAddTransaction(formData) {
  try {
    addTransaction(formData);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── Return category/subcategory pairs for dialog dropdowns ───────────────────
function getCategories() {
  return getCategoryList();
}
