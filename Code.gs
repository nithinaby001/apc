const SHEETS = {
  TRANSACTIONS: 'Transactions',
  CATEGORIES: 'Categories',
  RULES: 'Category_Rules',
  ACCOUNTS: 'Accounts'
};

const TRANSACTION_HEADERS = ['ID','Date','Description','Amount','Type','Category','Subcategory','Account','Source','Notes','Created Timestamp'];
const CATEGORY_HEADERS = ['Category ID','Category Name','Type','Parent Category','Active'];
const RULE_HEADERS = ['Keyword','Category','Subcategory'];
const ACCOUNT_HEADERS = ['Account ID','Account Name','Bank','Currency','Opening Balance'];

function doGet() {
  setupDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Personal Finance Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const ss = getDbSpreadsheet();
  ensureSheet(ss, SHEETS.TRANSACTIONS, TRANSACTION_HEADERS);
  ensureSheet(ss, SHEETS.CATEGORIES, CATEGORY_HEADERS);
  ensureSheet(ss, SHEETS.RULES, RULE_HEADERS);
  ensureSheet(ss, SHEETS.ACCOUNTS, ACCOUNT_HEADERS);
  seedDefaults(ss);
}

function getDbSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('No active spreadsheet found. Set Script Property SPREADSHEET_ID for standalone deployments.');
  }
  return SpreadsheetApp.openById(id);
}

function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const headerRange = sh.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0];
  if (existing.every((v) => !v)) {
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function seedDefaults(ss) {
  const cat = ss.getSheetByName(SHEETS.CATEGORIES);
  if (cat.getLastRow() <= 1) {
    cat.getRange(2, 1, 12, 5).setValues([
      ['CAT-EXP-FOOD', 'Food', 'Expense', '', true],
      ['CAT-EXP-TRANSPORT', 'Transport', 'Expense', '', true],
      ['CAT-EXP-RENT', 'Rent', 'Expense', '', true],
      ['CAT-EXP-UTILITIES', 'Utilities', 'Expense', '', true],
      ['CAT-EXP-GROCERIES', 'Groceries', 'Expense', 'Food', true],
      ['CAT-EXP-ENTERTAINMENT', 'Entertainment', 'Expense', '', true],
      ['CAT-EXP-MEDICAL', 'Medical', 'Expense', '', true],
      ['CAT-INC-SALARY', 'Salary', 'Income', '', true],
      ['CAT-INC-BONUS', 'Bonus', 'Income', '', true],
      ['CAT-INC-INVESTMENT', 'Investment', 'Income', '', true],
      ['CAT-INC-FREELANCE', 'Freelance', 'Income', '', true],
      ['CAT-EXP-UNCAT', 'Uncategorized', 'Expense', '', true]
    ]);
  }
  const rules = ss.getSheetByName(SHEETS.RULES);
  if (rules.getLastRow() <= 1) {
    rules.getRange(2, 1, 5, 3).setValues([
      ['uber', 'Transport', 'Ride'],
      ['talabat', 'Food', 'Delivery'],
      ['carrefour', 'Groceries', 'Supermarket'],
      ['netflix', 'Entertainment', 'Streaming'],
      ['amazon', 'Shopping', 'Online']
    ]);
  }
}

function addExpense(payload) { return addTransaction(Object.assign({}, payload, { type: 'Expense', source: payload.source || 'Manual' })); }
function addIncome(payload) { return addTransaction(Object.assign({}, payload, { type: 'Income', source: payload.source || 'Manual' })); }

function addTransaction(payload) {
  validateTx(payload);
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.TRANSACTIONS);
  const id = payload.id || ('TX-' + Utilities.getUuid().slice(0, 8));
  sh.appendRow([
    id,
    normalizeDate(payload.date),
    payload.description || '',
    Number(payload.amount || 0),
    payload.type,
    payload.category || '',
    payload.subcategory || '',
    payload.account || '',
    payload.source || 'Manual',
    payload.notes || '',
    new Date()
  ]);
  return { success: true, id };
}

function validateTx(payload) {
  if (!payload || !payload.type) throw new Error('Transaction type is required.');
  if (!payload.date) throw new Error('Transaction date is required.');
  if (!payload.description) throw new Error('Transaction description is required.');
  if (!Number(payload.amount)) throw new Error('Amount must be greater than zero.');
}

function getTransactions(filters) {
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.TRANSACTIONS);
  if (sh.getLastRow() <= 1) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, TRANSACTION_HEADERS.length).getValues();
  const txs = rows.map((r) => ({
    id: r[0], date: formatAsIsoDate(r[1]), description: r[2], amount: Number(r[3] || 0), type: r[4],
    category: r[5], subcategory: r[6], account: r[7], source: r[8], notes: r[9], createdTimestamp: r[10]
  }));
  return filterTransactions(txs, filters || {});
}

function filterTransactions(txs, f) {
  return txs.filter((tx) => {
    if (f.type && tx.type !== f.type) return false;
    if (f.category && tx.category !== f.category) return false;
    if (f.account && tx.account !== f.account) return false;
    if (f.search && !(`${tx.description} ${tx.notes}`.toLowerCase().includes(String(f.search).toLowerCase()))) return false;
    if (f.startDate && new Date(tx.date) < new Date(f.startDate)) return false;
    if (f.endDate && new Date(tx.date) > new Date(f.endDate)) return false;
    return true;
  });
}

function updateTransaction(payload) {
  validateTx(payload);
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.TRANSACTIONS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === payload.id) {
      sh.getRange(i + 1, 2, 1, 9).setValues([[
        normalizeDate(payload.date), payload.description, Number(payload.amount), payload.type,
        payload.category || '', payload.subcategory || '', payload.account || '', payload.source || 'Manual', payload.notes || ''
      ]]);
      return { success: true };
    }
  }
  return { success: false, message: 'Transaction not found' };
}

function deleteTransaction(id) {
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.TRANSACTIONS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Transaction not found' };
}

function bulkUpdateCategory(payload) {
  const ids = payload.ids || [];
  const category = payload.category || '';
  const subcategory = payload.subcategory || '';
  if (!ids.length || !category) return { success: false, updated: 0 };
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.TRANSACTIONS);
  const rows = sh.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < rows.length; i++) {
    if (ids.indexOf(rows[i][0]) > -1) {
      sh.getRange(i + 1, 6).setValue(category);
      sh.getRange(i + 1, 7).setValue(subcategory);
      updated += 1;
    }
  }
  return { success: true, updated };
}

function uploadBankStatement(payload) {
  const rows = payload.rows || [];
  const imported = [];
  rows.forEach((r) => {
    const debit = Number(r.debit || 0);
    const credit = Number(r.credit || 0);
    if (!debit && !credit) return;
    const type = debit > 0 ? 'Expense' : 'Income';
    const amount = debit > 0 ? debit : credit;
    const cat = categorizeByDescription(r.description, type);
    const tx = {
      date: normalizeImportedDate(r.date), description: r.description || '', amount, type,
      category: cat.category, subcategory: cat.subcategory, account: payload.account || '', source: 'Bank Import', notes: payload.notes || ''
    };
    addTransaction(tx);
    imported.push(tx);
  });
  return { success: true, count: imported.length, imported };
}

function normalizeImportedDate(v) {
  if (!v) return new Date();
  const raw = String(v).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const p = raw.split('/');
    return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return new Date(raw + 'T00:00:00');
  return new Date(raw);
}

function categorizeTransactions() {
  const txs = getTransactions();
  let updatedCount = 0;
  txs.forEach((tx) => {
    if (!tx.category || tx.category === 'Uncategorized') {
      const cat = categorizeByDescription(tx.description, tx.type);
      if (cat.category) {
        tx.category = cat.category;
        tx.subcategory = cat.subcategory;
        updateTransaction(tx);
        updatedCount += 1;
      }
    }
  });
  return { success: true, updatedCount };
}

function categorizeByDescription(description, type) {
  const desc = String(description || '').toLowerCase();
  const rules = getCategoryRules();
  for (let i = 0; i < rules.length; i++) {
    const key = String(rules[i].keyword || '').toLowerCase();
    if (key && desc.indexOf(key) > -1) {
      return { category: rules[i].category || defaultCategory(type), subcategory: rules[i].subcategory || '' };
    }
  }
  return { category: defaultCategory(type), subcategory: '' };
}

function defaultCategory(type) {
  return type === 'Income' ? 'Salary' : 'Uncategorized';
}

function getDashboardData() {
  const txs = getTransactions();
  const now = new Date();
  const mStart = new Date(now.getFullYear(), now.getMonth(), 1);
  let income = 0, expense = 0;
  const monthMap = {}, catMap = {};

  txs.forEach((tx) => {
    const d = new Date(tx.date);
    const ym = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
    if (!monthMap[ym]) monthMap[ym] = { income: 0, expense: 0 };
    if (tx.type === 'Income') monthMap[ym].income += tx.amount; else {
      monthMap[ym].expense += tx.amount;
      catMap[tx.category || 'Uncategorized'] = (catMap[tx.category || 'Uncategorized'] || 0) + tx.amount;
    }
    if (d >= mStart) {
      if (tx.type === 'Income') income += tx.amount; else expense += tx.amount;
    }
  });

  const months = Object.keys(monthMap).sort();
  const categoryBreakdown = Object.keys(catMap).map((k) => ({ category: k, amount: catMap[k] })).sort((a, b) => b.amount - a.amount);

  return {
    summary: {
      currentBalance: getOpeningBalanceTotal() + txs.reduce((s, tx) => s + (tx.type === 'Income' ? tx.amount : -tx.amount), 0),
      totalIncomeMonth: income,
      totalExpensesMonth: expense,
      savings: income - expense
    },
    charts: {
      expenseTrend: months.map((m) => ({ month: m, value: monthMap[m].expense })),
      incomeExpense: months.map((m) => ({ month: m, income: monthMap[m].income, expense: monthMap[m].expense })),
      categoryBreakdown,
      topSpendingCategories: categoryBreakdown.slice(0, 5)
    },
    recentTransactions: txs.sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 10)
  };
}

function getFinancialReports() {
  const txs = getTransactions();
  const monthly = {}, category = {}, daily = {};
  txs.forEach((tx) => {
    const m = tx.date.slice(0, 7);
    if (!monthly[m]) monthly[m] = { month: m, income: 0, expense: 0, savings: 0 };
    if (tx.type === 'Income') monthly[m].income += tx.amount; else {
      monthly[m].expense += tx.amount;
      category[tx.category || 'Uncategorized'] = (category[tx.category || 'Uncategorized'] || 0) + tx.amount;
      daily[tx.date] = (daily[tx.date] || 0) + tx.amount;
    }
    monthly[m].savings = monthly[m].income - monthly[m].expense;
  });
  return {
    monthlySummary: Object.values(monthly).sort((a, b) => a.month > b.month ? 1 : -1),
    categoryReport: Object.keys(category).map((k) => ({ category: k, total: category[k] })).sort((a, b) => b.total - a.total),
    dailyExpenseTrend: Object.keys(daily).sort().map((d) => ({ date: d, total: daily[d] }))
  };
}

function getCategories() {
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.CATEGORIES);
  if (sh.getLastRow() <= 1) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map((r) => ({ id: r[0], name: r[1], type: r[2], parentCategory: r[3], active: r[4] }));
}

function addCategory(payload) {
  if (!payload.name || !payload.type) throw new Error('Category name and type are required.');
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.CATEGORIES);
  sh.appendRow(['CAT-' + Utilities.getUuid().slice(0, 8), payload.name, payload.type, payload.parentCategory || '', payload.active !== false]);
  return { success: true };
}

function getAccounts() {
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.ACCOUNTS);
  if (sh.getLastRow() <= 1) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map((r) => ({ id: r[0], name: r[1], bank: r[2], currency: r[3], openingBalance: Number(r[4] || 0) }));
}

function addAccount(payload) {
  if (!payload.name) throw new Error('Account name is required.');
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.ACCOUNTS);
  sh.appendRow(['ACC-' + Utilities.getUuid().slice(0, 8), payload.name, payload.bank || '', payload.currency || 'USD', Number(payload.openingBalance || 0)]);
  return { success: true };
}

function getCategoryRules() {
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.RULES);
  if (sh.getLastRow() <= 1) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter((r) => r[0])
    .map((r) => ({ keyword: r[0], category: r[1], subcategory: r[2] }));
}

function addCategoryRule(payload) {
  if (!payload.keyword || !payload.category) throw new Error('Keyword and category are required.');
  const sh = getDbSpreadsheet().getSheetByName(SHEETS.RULES);
  sh.appendRow([payload.keyword, payload.category, payload.subcategory || '']);
  return { success: true };
}

function getOpeningBalanceTotal() {
  return getAccounts().reduce((s, a) => s + Number(a.openingBalance || 0), 0);
}

function normalizeDate(v) {
  const d = v ? new Date(v) : new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatAsIsoDate(v) {
  if (!v) return '';
  return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
