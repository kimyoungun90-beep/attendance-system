PRAGMA foreign_keys = ON;

ALTER TABLE attendance_closures ADD COLUMN unexpected_count INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN unexpected_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN dayoff_excess_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN substitute_shortage_people INTEGER NOT NULL DEFAULT 0;

CREATE TABLE IF NOT EXISTS attendance_issue_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  issue_type TEXT NOT NULL CHECK (issue_type IN ('missing_clock_in', 'unexpected_clock_in')),
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  issue_date TEXT NOT NULL,
  weekday TEXT,
  plan_status TEXT,
  actual_in TEXT,
  changed_in TEXT,
  clock_status TEXT,
  result TEXT,
  reason TEXT,
  duplicate_plan_note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, issue_type, employee_id, issue_date)
);

CREATE INDEX IF NOT EXISTS idx_attendance_issue_items_closure
ON attendance_issue_items(closure_id, issue_type, issue_date, employee_id);

CREATE TABLE IF NOT EXISTS attendance_employee_summaries (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  base_allowance REAL NOT NULL DEFAULT 0,
  basic_dayoff_used REAL NOT NULL DEFAULT 0,
  explicit_sub_dayoff_used REAL NOT NULL DEFAULT 0,
  base_excess REAL NOT NULL DEFAULT 0,
  substitute_needed REAL NOT NULL DEFAULT 0,
  available_substitute REAL NOT NULL DEFAULT 0,
  substitute_applied REAL NOT NULL DEFAULT 0,
  remaining_substitute REAL NOT NULL DEFAULT 0,
  shortage REAL NOT NULL DEFAULT 0,
  judgment TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_employee_summaries_month_company
ON attendance_employee_summaries(company, month, employee_id);

CREATE TABLE IF NOT EXISTS substitute_dayoff_grants (
  id TEXT PRIMARY KEY,
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  store TEXT,
  grant_month TEXT NOT NULL,
  granted_days REAL NOT NULL CHECK (granted_days > 0),
  valid_from TEXT NOT NULL,
  valid_to TEXT NOT NULL,
  reason TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_substitute_grants_employee
ON substitute_dayoff_grants(company, employee_id, valid_to, valid_from);

CREATE TABLE IF NOT EXISTS substitute_dayoff_allocations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  grant_id TEXT NOT NULL,
  closure_id TEXT NOT NULL,
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  employee_id TEXT NOT NULL,
  month TEXT NOT NULL,
  used_days REAL NOT NULL CHECK (used_days > 0),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (grant_id) REFERENCES substitute_dayoff_grants(id) ON DELETE RESTRICT,
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(grant_id, closure_id)
);

CREATE INDEX IF NOT EXISTS idx_substitute_allocations_employee
ON substitute_dayoff_allocations(company, employee_id, month, grant_id);
