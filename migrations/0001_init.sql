PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS attendance_closures (
  id TEXT PRIMARY KEY,
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  cutoff_date TEXT NOT NULL,
  check_mode TEXT NOT NULL DEFAULT 'strict',
  plan_file_name TEXT,
  attendance_file_name TEXT,
  plan_people INTEGER NOT NULL DEFAULT 0,
  attendance_people INTEGER NOT NULL DEFAULT 0,
  matched_people INTEGER NOT NULL DEFAULT 0,
  match_rate REAL NOT NULL DEFAULT 0,
  missing_count INTEGER NOT NULL DEFAULT 0,
  missing_people INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_attendance_closures_month_company
ON attendance_closures(month, company, created_at DESC);

CREATE TABLE IF NOT EXISTS attendance_missing_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  company TEXT NOT NULL,
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  missing_date TEXT NOT NULL,
  weekday TEXT,
  plan_status TEXT,
  actual_in TEXT,
  changed_in TEXT,
  clock_status TEXT,
  result TEXT,
  reason TEXT,
  duplicate_plan_note TEXT,
  review_status TEXT NOT NULL DEFAULT '미확인',
  review_note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, employee_id, missing_date)
);

CREATE INDEX IF NOT EXISTS idx_missing_items_closure
ON attendance_missing_items(closure_id, missing_date, employee_id);

-- 2단계부터 사용할 월별 휴무 기준/대체휴무 저장 구조
CREATE TABLE IF NOT EXISTS monthly_dayoff_rules (
  id TEXT PRIMARY KEY,
  company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  base_dayoff_count INTEGER NOT NULL DEFAULT 0,
  substitute_dayoff_count INTEGER NOT NULL DEFAULT 0,
  valid_from TEXT,
  valid_to TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(company, month)
);
