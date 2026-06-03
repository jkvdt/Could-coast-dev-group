function showCsvImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CsvImportDialog')
    .setWidth(560).setHeight(480).setTitle('Import Transactions from CSV');
  SpreadsheetApp.getUi().showModalDialog(html, 'Import CSV');
}

// Called from the CsvImportDialog with the raw CSV text + user mapping choices.
function importCsv(csvText, mapping, accountName, source) {
  try {
    const rows = _parseCsv(csvText);
    if (rows.length < 2) return { ok: false, error: 'CSV appears to be empty or has no data rows.' };

    const headers = rows[0].map(h => h.toLowerCase().trim());
    const dataRows = rows.slice(1);

    // Resolve column indices from user mapping or auto-detect
    const colDate   = _resolveCol(headers, mapping.date,   ['date','transaction date','posted date','trans. date']);
    const colDesc   = _resolveCol(headers, mapping.desc,   ['description','merchant','name','payee','transaction']);
    const colAmount = _resolveCol(headers, mapping.amount, ['amount','transaction amount','debit','credit']);
    const colDebit  = _resolveCol(headers, mapping.debit,  ['debit','withdrawal','debit amount']);
    const colCredit = _resolveCol(headers, mapping.credit, ['credit','deposit','credit amount']);

    if (colDate === -1) return { ok: false, error: 'Could not find a Date column. Check your column mapping.' };
    if (colDesc === -1) return { ok: false, error: 'Could not find a Description column.' };
    if (colAmount === -1 && colDebit === -1) return { ok: false, error: 'Could not find an Amount or Debit column.' };

    const txRows = [];
    for (const r of dataRows) {
      if (r.every(c => !c.trim())) continue; // skip blank rows

      const dateRaw = r[colDate] ? r[colDate].trim() : '';
      if (!dateRaw) continue;
      const date = new Date(dateRaw);
      if (isNaN(date.getTime())) continue;

      const name = colDesc !== -1 ? r[colDesc].trim() : '';
      let amount = 0;

      if (colAmount !== -1) {
        // Single amount column: negative = expense, positive = income
        amount = _parseAmount(r[colAmount]);
      } else {
        // Separate debit/credit columns
        const debit  = colDebit  !== -1 ? _parseAmount(r[colDebit])  : 0;
        const credit = colCredit !== -1 ? _parseAmount(r[colCredit]) : 0;
        // Debits reduce balance → negative; credits add → positive
        amount = credit - debit;
      }

      const [cat, sub] = categorize(name);
      const txId = `csv-${Utilities.computeDigest(
        Utilities.DigestAlgorithm.MD5,
        `${dateRaw}|${name}|${amount}`
      ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0')).join('')}`;

      txRows.push([
        date, accountName || 'Imported', name, amount, cat, sub, '', txId, source || 'CSV', '',
      ]);
    }

    const inserted = insertTransactions(txRows);
    refreshDashboard();
    return { ok: true, total: txRows.length, inserted };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _parseCsv(text) {
  const rows = [];
  let row = [], field = '', inQuote = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuote) {
      if (ch === '"' && text[i+1] === '"') { field += '"'; i++; }
      else if (ch === '"') inQuote = false;
      else field += ch;
    } else if (ch === '"') {
      inQuote = true;
    } else if (ch === ',') {
      row.push(field); field = '';
    } else if (ch === '\n') {
      row.push(field); rows.push(row); row = []; field = '';
    } else if (ch === '\r') {
      // ignore CR
    } else {
      field += ch;
    }
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(c => c.trim()));
}

function _resolveCol(headers, userChoice, autoKeywords) {
  if (userChoice !== undefined && userChoice !== '' && userChoice !== null) {
    const idx = parseInt(userChoice, 10);
    if (!isNaN(idx)) return idx;
    const lower = String(userChoice).toLowerCase().trim();
    const found = headers.indexOf(lower);
    if (found !== -1) return found;
  }
  for (const kw of autoKeywords) {
    const idx = headers.findIndex(h => h.includes(kw));
    if (idx !== -1) return idx;
  }
  return -1;
}

function _parseAmount(raw) {
  if (!raw) return 0;
  const cleaned = String(raw).replace(/[$,\s]/g, '').replace(/\((.+)\)/, '-$1');
  return parseFloat(cleaned) || 0;
}

// Returns first 5 rows of CSV + detected headers for preview in dialog
function previewCsv(csvText) {
  try {
    const rows = _parseCsv(csvText);
    return { ok: true, headers: rows[0] || [], preview: rows.slice(1, 6) };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
