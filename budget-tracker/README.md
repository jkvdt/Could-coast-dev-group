# Google Sheets Budget Tracker

A Tiller-style personal finance tracker with:
- Monthly & annual budget vs. actual views
- Automatic transaction categorization (55+ rules)
- CSV import from any bank (Chase, BofA, Amex, Capital One, etc.)
- Live bank/credit card sync via Plaid (optional)
- Dashboard with account balances, budget status, recent transactions

---

## Quick Start (10 minutes)

### 1. Create a new Google Spreadsheet

Go to [sheets.new](https://sheets.new) and create a blank spreadsheet.
Name it something like **Personal Budget 2026**.

### 2. Open Apps Script

In your spreadsheet: **Extensions → Apps Script**

### 3. Deploy via clasp (recommended)

Install [clasp](https://github.com/google/clasp) if you haven't:

```bash
npm install -g @google/clasp
clasp login
```

Copy `.clasp.json.template` to `.clasp.json` and fill in your Script ID
(found in Apps Script: Project Settings → Script ID):

```bash
cp .clasp.json.template .clasp.json
# edit .clasp.json and replace YOUR_SCRIPT_ID_HERE
```

Push the code:

```bash
cd budget-tracker
clasp push
```

### 4. Alternative: copy-paste manually

If you prefer not to use clasp, open each `.gs` file from the `src/` folder
and paste its contents into new script files in the Apps Script editor.
Do the same for the `.html` files (use **File → New → HTML file**).
File name must match exactly (e.g. `PlaidSidebar`, `CsvImportDialog`, `AddTransactionDialog`).

### 5. Run Setup

Back in your spreadsheet, refresh the page. You'll see a new
**💰 Budget Tracker** menu. Click:

```
💰 Budget Tracker → (wait for menu to load) → Settings → Reset / Reinstall sheets
```

Or open Apps Script and run `setupSpreadsheet()` directly.

This creates all sheets and populates the default budget categories.

---

## Using the Tracker

### Option A — CSV Import (works immediately, no API needed)

1. Log into your bank or credit card website
2. Download a CSV export (usually under Statements or Activity)
3. In your spreadsheet: **Budget Tracker → Import Transactions → Import from CSV file…**
4. Paste the CSV, click "Preview & Map Columns", adjust column mapping if needed, then Import

Supported automatically: Chase, Bank of America, Wells Fargo, Capital One, Citi, American Express, and most others.

### Option B — Live Bank Sync via Plaid (pulls transactions automatically)

Plaid is the same service Tiller uses. For personal use it's free or very low cost.

#### Get Plaid credentials

1. Sign up at [dashboard.plaid.com](https://dashboard.plaid.com)
2. Create an app → go to **Team Settings → Keys**
3. Copy your `client_id` and `secret` (use the **Production** keys for real accounts, or **Sandbox** for testing)

#### Enter credentials in the spreadsheet

```
💰 Budget Tracker → Settings → Set Plaid API credentials…
```

Enter your client_id, secret, and environment (`production` or `sandbox`).
Credentials are stored encrypted in Google's Script Properties — never in the sheet itself.

#### Connect accounts

```
💰 Budget Tracker → Live Bank Sync (Plaid) → Connect a bank account…
```

A sidebar opens Plaid Link. Search for your bank, log in, select accounts.
The script immediately pulls 90 days of transaction history.

#### Sync going forward

```
💰 Budget Tracker → Live Bank Sync (Plaid) → Sync transactions now
```

Uses Plaid's incremental sync cursor — only pulls new transactions each time.
Run this weekly or set up a time-based trigger in Apps Script to automate it.

---

## Sheets Overview

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | Monthly snapshot — balances, budget status, recent transactions |
| **Transactions** | Every transaction. Edit Category/Subcategory here. |
| **Budget** | Set monthly budgets per category. Actuals auto-calculate via SUMIFS. |
| **Annual Summary** | 12-month grid — regenerate via menu |
| **Accounts** | Connected accounts and balances |
| **Categories** | Master category list — add custom categories here |
| **Settings** | Year, currency symbol, Plaid config |

---

## Customizing Categories

1. Open the **Categories** sheet
2. Add a new row: Category, Subcategory, Type (Income or Expense), Monthly Budget
3. The new category immediately appears in the Budget sheet dropdowns and transaction dialogs

To add auto-categorization rules, edit `CATEGORY_RULES` in `00_Config.gs`.

---

## Setting Up Automatic Sync (optional)

In Apps Script editor: **Triggers (clock icon) → Add Trigger**

- Function: `syncAllPlaidAccounts`
- Event source: Time-driven
- Type: Week timer, every Monday at 7am

---

## Vs. Tiller / Monarch / Rocket Money

| | This tracker | Tiller | Monarch | Rocket Money |
|--|--|--|--|--|
| Cost | Free (+ Plaid ~$0 personal) | $79/yr | $100/yr | Free–$72/yr |
| Live bank feeds | ✅ Plaid | ✅ Plaid | ✅ | ✅ |
| Google Sheets | ✅ | ✅ | ❌ | ❌ |
| Fully customizable | ✅ | Partially | ❌ | ❌ |
| Subscription tracker | Manual | Manual | ✅ | ✅ Best-in-class |
| You own the data | ✅ | ✅ | ❌ | ❌ |
| Setup effort | ~10 min | ~5 min | 2 min | 2 min |

**Recommendation:** Use this if you want full control and a Sheets-based workflow.
Use **Monarch** if you want the best app experience for behavioral budgeting.
Use **Rocket Money** if your biggest pain point is finding and canceling subscriptions.
