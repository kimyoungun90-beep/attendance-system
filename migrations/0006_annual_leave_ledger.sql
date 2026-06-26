PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS annual_leave_baseline_uploads (
  route TEXT PRIMARY KEY CHECK (route IN ('homeplus', 'electroland')),
  baseline_date TEXT NOT NULL,
  file_name TEXT,
  employee_count INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS annual_leave_employees (
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  regional_manager TEXT,
  manager TEXT,
  region1 TEXT,
  region2 TEXT,
  store_code TEXT,
  store_name TEXT,
  portal_id TEXT,
  hire_date TEXT,
  basis_hire_date TEXT,
  policy_type TEXT NOT NULL CHECK (policy_type IN ('jan1', 'anniversary')),
  under_one_year INTEGER NOT NULL DEFAULT 0,
  baseline_date TEXT NOT NULL,
  baseline_granted REAL NOT NULL DEFAULT 0,
  baseline_used REAL NOT NULL DEFAULT 0,
  baseline_remaining REAL NOT NULL DEFAULT 0,
  cycle_start TEXT,
  cycle_end TEXT,
  first_promotion_date TEXT,
  second_promotion_date TEXT,
  expiry_date TEXT,
  termination_date TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  PRIMARY KEY (route, employee_id)
);

CREATE TABLE IF NOT EXISTS annual_leave_monthly_uploads (
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  file_name TEXT,
  row_count INTEGER NOT NULL DEFAULT 0,
  approved_days REAL NOT NULL DEFAULT 0,
  rejected_days REAL NOT NULL DEFAULT 0,
  pending_days REAL NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  PRIMARY KEY (route, month)
);

CREATE TABLE IF NOT EXISTS annual_leave_applications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  employee_id TEXT NOT NULL,
  employee_name TEXT,
  leave_date TEXT NOT NULL,
  days REAL NOT NULL DEFAULT 0,
  status TEXT,
  leave_type TEXT,
  application_date TEXT,
  note TEXT,
  source_index INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_annual_leave_employees_route
ON annual_leave_employees(route, under_one_year, employee_id);

CREATE INDEX IF NOT EXISTS idx_annual_leave_applications_route_month
ON annual_leave_applications(route, month, employee_id, leave_date);

CREATE INDEX IF NOT EXISTS idx_annual_leave_applications_employee_date
ON annual_leave_applications(route, employee_id, leave_date);
