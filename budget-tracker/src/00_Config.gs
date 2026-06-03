// ─── Sheet names ───────────────────────────────────────────────────────────────
const SHEET = {
  DASHBOARD:    'Dashboard',
  TRANSACTIONS: 'Transactions',
  BUDGET:       'Budget',
  ANNUAL:       'Annual Summary',
  ACCOUNTS:     'Accounts',
  CATEGORIES:   'Categories',
  SETTINGS:     'Settings',
};

// ─── Transaction column indices (1-based) ─────────────────────────────────────
const TX_COL = {
  DATE:     1,
  ACCOUNT:  2,
  NAME:     3,
  AMOUNT:   4,
  CATEGORY: 5,
  SUBCAT:   6,
  NOTES:    7,
  TX_ID:    8,
  SOURCE:   9,
  MONTH:    10,
};

// ─── Budget column indices (1-based) ──────────────────────────────────────────
// Layout: Category | Subcategory | Type | Jan Bud | Jan Act | Jan Var | Feb … | Annual Bud | Annual Act | Annual Var
const BUDGET_HEADER_COLS = 3; // Category, Subcategory, Type
const BUDGET_COLS_PER_MONTH = 3; // Budget, Actual, Variance
const BUDGET_ANNUAL_START_COL = BUDGET_HEADER_COLS + (12 * BUDGET_COLS_PER_MONTH) + 1;

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// ─── Default categories ───────────────────────────────────────────────────────
const DEFAULT_CATEGORIES = [
  // [Category, Subcategory, Type, DefaultMonthlyBudget]
  ['Income', 'Salary / Wages',       'Income',  5000],
  ['Income', 'Freelance',            'Income',     0],
  ['Income', 'Investment Returns',   'Income',     0],
  ['Income', 'Other Income',         'Income',     0],

  ['Housing', 'Rent / Mortgage',     'Expense', 1500],
  ['Housing', 'Electric',            'Expense',   80],
  ['Housing', 'Gas / Heat',          'Expense',   60],
  ['Housing', 'Water / Sewer',       'Expense',   40],
  ['Housing', 'Internet',            'Expense',   60],
  ['Housing', 'Phone',               'Expense',   80],
  ['Housing', 'Home Insurance',      'Expense',   75],
  ['Housing', 'HOA / Condo Fees',    'Expense',    0],
  ['Housing', 'Maintenance / Repairs','Expense',  50],

  ['Transportation', 'Car Payment',  'Expense',  400],
  ['Transportation', 'Gas',          'Expense',  120],
  ['Transportation', 'Car Insurance','Expense',  110],
  ['Transportation', 'Parking / Tolls','Expense', 30],
  ['Transportation', 'Public Transit','Expense',   0],
  ['Transportation', 'Uber / Lyft',  'Expense',   40],
  ['Transportation', 'Car Maintenance','Expense',  30],

  ['Food', 'Groceries',              'Expense',  400],
  ['Food', 'Restaurants',            'Expense',  200],
  ['Food', 'Coffee Shops',           'Expense',   60],
  ['Food', 'Food Delivery',          'Expense',   80],
  ['Food', 'Alcohol / Bars',         'Expense',   40],

  ['Health', 'Health Insurance',     'Expense',  250],
  ['Health', 'Doctor / Dental',      'Expense',   50],
  ['Health', 'Prescriptions',        'Expense',   20],
  ['Health', 'Gym / Fitness',        'Expense',   40],

  ['Subscriptions', 'Streaming',     'Expense',   50],
  ['Subscriptions', 'Software / Apps','Expense',  30],
  ['Subscriptions', 'News / Media',  'Expense',   20],
  ['Subscriptions', 'Other Subscriptions','Expense',0],

  ['Entertainment', 'Movies / Events','Expense',  50],
  ['Entertainment', 'Games',         'Expense',   20],
  ['Entertainment', 'Hobbies',       'Expense',   50],
  ['Entertainment', 'Sports / Recreation','Expense',40],

  ['Shopping', 'Clothing',           'Expense',  100],
  ['Shopping', 'Electronics',        'Expense',   50],
  ['Shopping', 'Amazon / Online',    'Expense',  100],
  ['Shopping', 'Home Goods',         'Expense',   50],

  ['Personal Care', 'Haircuts / Salon','Expense', 40],
  ['Personal Care', 'Cosmetics / Grooming','Expense',30],

  ['Education', 'Student Loans',     'Expense',    0],
  ['Education', 'Tuition / Courses', 'Expense',    0],
  ['Education', 'Books',             'Expense',   20],

  ['Travel', 'Flights',              'Expense',    0],
  ['Travel', 'Hotels / Lodging',     'Expense',    0],
  ['Travel', 'Vacation Activities',  'Expense',    0],

  ['Gifts / Giving', 'Gifts',        'Expense',   50],
  ['Gifts / Giving', 'Donations',    'Expense',   25],

  ['Financial', 'Savings Transfer',  'Expense',  500],
  ['Financial', 'Investment Contribution','Expense',200],
  ['Financial', 'Other Debt',        'Expense',    0],
  ['Financial', 'Bank Fees',         'Expense',    0],

  ['Uncategorized', 'Uncategorized', 'Expense',    0],
];

// ─── Auto-categorization rules  (keyword → [Category, Subcategory]) ──────────
const CATEGORY_RULES = [
  [/netflix|hulu|disney\+|hbo|paramount|peacock|apple tv|youtube premium|spotify|amazon music|pandora/i, ['Subscriptions','Streaming']],
  [/amazon\.com|amazon prime|amzn/i,                   ['Shopping','Amazon / Online']],
  [/uber eats|doordash|grubhub|postmates|instacart/i,  ['Food','Food Delivery']],
  [/uber|lyft/i,                                       ['Transportation','Uber / Lyft']],
  [/starbucks|dunkin|dutch bros|peet'?s/i,             ['Food','Coffee Shops']],
  [/whole foods|trader joe|kroger|publix|safeway|aldi|wegmans|costco|sam'?s club|walmart|target/i, ['Food','Groceries']],
  [/mcdonald|chick-fil|chipotle|subway|taco bell|pizza hut|domino|wendy|burger king|panera|olive garden|applebee|ihop|denny/i, ['Food','Restaurants']],
  [/cvs|walgreens|rite aid|duane reade/i,              ['Health','Prescriptions']],
  [/planet fitness|24 hour fitness|la fitness|orange theory|equinox|ymca/i, ['Health','Gym / Fitness']],
  [/at&t|verizon|t-mobile|sprint|comcast|xfinity|spectrum|cox|optimum/i, ['Housing','Phone']],
  [/electric|gas bill|water bill|pge|con ed|national grid|duke energy/i, ['Housing','Electric']],
  [/rent|property management|landlord/i,               ['Housing','Rent / Mortgage']],
  [/shell|chevron|exxon|bp|mobil|sunoco|valero|speedway|wawa|pilot|circle k/i, ['Transportation','Gas']],
  [/geico|state farm|progressive|allstate|farmers|liberty mutual/i, ['Transportation','Car Insurance']],
  [/southwest|delta|united|american airlines|jetblue|spirit|frontier/i, ['Travel','Flights']],
  [/marriott|hilton|hyatt|sheraton|holiday inn|airbnb|vrbo/i,           ['Travel','Hotels / Lodging']],
  [/github|adobe|microsoft 365|google workspace|dropbox|notion|slack|zoom/i, ['Subscriptions','Software / Apps']],
  [/new york times|wall street journal|washington post|the atlantic/i,  ['Subscriptions','News / Media']],
  [/paycheck|direct deposit|salary|payroll/i,          ['Income','Salary / Wages']],
  [/venmo|zelle|paypal|cash app/i,                     ['Uncategorized','Uncategorized']],
];

// ─── Color palette ─────────────────────────────────────────────────────────────
const COLOR = {
  HEADER_BG:    '#1a1a2e',
  HEADER_FG:    '#ffffff',
  INCOME_BG:    '#e6f4ea',
  EXPENSE_BG:   '#fce8e6',
  SECTION_BG:   '#f0f4ff',
  OVER_BUDGET:  '#ea4335',
  UNDER_BUDGET: '#34a853',
  NEUTRAL:      '#fbbc04',
  BORDER:       '#dadce0',
  ALT_ROW:      '#f8f9fa',
};
