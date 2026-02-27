# Personal Finance Management Web App (Google Apps Script)

A responsive personal finance web app backed by Google Sheets, implemented as a Google Apps Script Web App.

## Features
- Dashboard with balance/income/expense/savings cards and charts
- Expense & income entry forms
- Bank statement import (CSV/TSV) with preview and confirm
- Automatic expense categorization using keyword rules
- Transaction management (filter, edit, delete, bulk category update)
- Financial reports (monthly summary, category breakdown, daily trend)
- Settings section for category, account, and categorization-rule creation

## Google Sheets schema
The app creates and uses:
- `Transactions`
- `Categories`
- `Category_Rules`
- `Accounts`

## Backend APIs
- `getDashboardData()`
- `addExpense(payload)`
- `addIncome(payload)`
- `getTransactions(filters)`
- `updateTransaction(payload)`
- `deleteTransaction(id)`
- `bulkUpdateCategory(payload)`
- `uploadBankStatement(payload)`
- `categorizeTransactions()`
- `getCategories()` / `addCategory(payload)`
- `getAccounts()` / `addAccount(payload)`
- `getCategoryRules()` / `addCategoryRule(payload)`
- `getFinancialReports()`

## Deploy
1. Create/open a Google Spreadsheet.
2. Open **Extensions → Apps Script** and add:
   - `Code.gs`
   - `Index.html`
   - `appsscript.json`
3. If using a **standalone** Apps Script project, set Script Property `SPREADSHEET_ID`.
4. Deploy as **Web App**:
   - Execute as: **Me**
   - Access: **Anyone with link**

> Open the deployment URL directly (do not open `Index.html` as a local file), because the UI depends on `google.script.run`.
