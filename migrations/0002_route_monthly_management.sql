PRAGMA foreign_keys = ON;

-- 현재 운영 중인 v1(attendance_closures)에서 한 번만 실행합니다.
-- 사용자 화면에서는 company를 모두 '경로'로 표시하지만,
-- v1 호환을 위해 attendance_closures.company 컬럼명은 그대로 유지합니다.

ALTER TABLE attendance_closures ADD COLUMN unexpected_count INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN unexpected_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN mismatch_count INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN mismatch_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN dayoff_excess_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN substitute_shortage_people INTEGER NOT NULL DEFAULT 0;
ALTER TABLE attendance_closures ADD COLUMN annual_leave_people INTEGER NOT NULL DEFAULT 0;

-- 과거에 같은 경로·월이 여러 건 저장되어 있다면 가장 최근 행만 남깁니다.
DELETE FROM attendance_closures
WHERE rowid NOT IN (
  SELECT MAX(rowid)
  FROM attendance_closures
  GROUP BY company, month
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_attendance_closures_route_month
ON attendance_closures(company, month);

CREATE TABLE IF NOT EXISTS attendance_issue_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  issue_type TEXT NOT NULL CHECK (issue_type IN ('missing_clock_in', 'unexpected_clock_in')),
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  issue_date TEXT NOT NULL,
  weekday TEXT,
  plan_status TEXT,
  actual_status TEXT,
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

CREATE TABLE IF NOT EXISTS attendance_mismatch_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  issue_date TEXT NOT NULL,
  weekday TEXT,
  plan_status TEXT,
  actual_status TEXT,
  actual_in TEXT,
  changed_in TEXT,
  result TEXT,
  reason TEXT,
  duplicate_plan_note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, employee_id, issue_date)
);

CREATE INDEX IF NOT EXISTS idx_attendance_mismatch_items_closure
ON attendance_mismatch_items(closure_id, issue_date, employee_id);

-- 업로드한 월 마감의 원본 계산값입니다.
-- 과거 월을 교체하면 이 원본값을 기준으로 이후 월 누적값을 전부 다시 계산합니다.
CREATE TABLE IF NOT EXISTS attendance_monthly_employee_facts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  store TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  base_allowance REAL NOT NULL DEFAULT 0,
  basic_dayoff_used REAL NOT NULL DEFAULT 0,
  explicit_sub_dayoff_used REAL NOT NULL DEFAULT 0,
  base_excess REAL NOT NULL DEFAULT 0,
  substitute_needed REAL NOT NULL DEFAULT 0,
  annual_leave_used REAL NOT NULL DEFAULT 0,
  substitute_events_json TEXT NOT NULL DEFAULT '[]',
  annual_leave_events_json TEXT NOT NULL DEFAULT '[]',
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_monthly_employee_facts_route_month
ON attendance_monthly_employee_facts(route, month, employee_id);

-- 경로·발생 월별 공통 부여 설정: 같은 경로·같은 월은 1건만 존재합니다.
CREATE TABLE IF NOT EXISTS route_substitute_grants (
  id TEXT PRIMARY KEY,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  grant_month TEXT NOT NULL,
  granted_days REAL NOT NULL CHECK (granted_days > 0),
  valid_from TEXT NOT NULL,
  valid_to TEXT NOT NULL,
  reason TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(route, grant_month)
);

CREATE INDEX IF NOT EXISTS idx_route_substitute_grants_period
ON route_substitute_grants(route, grant_month, valid_to, valid_from);

-- 재계산으로 생성되는 직원별·월별 대체휴무 차감 내역입니다.
CREATE TABLE IF NOT EXISTS route_substitute_allocations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  grant_id TEXT NOT NULL,
  closure_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  employee_id TEXT NOT NULL,
  month TEXT NOT NULL,
  used_days REAL NOT NULL CHECK (used_days > 0),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (grant_id) REFERENCES route_substitute_grants(id) ON DELETE CASCADE,
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(grant_id, closure_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_route_substitute_allocations_employee
ON route_substitute_allocations(route, employee_id, month, grant_id);

-- 월 마감 원본을 시간순으로 다시 계산한 최종 결과입니다.
CREATE TABLE IF NOT EXISTS attendance_monthly_summaries_v3 (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  closure_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
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
  expired_substitute REAL NOT NULL DEFAULT 0,
  shortage REAL NOT NULL DEFAULT 0,
  current_annual_leave REAL NOT NULL DEFAULT 0,
  cumulative_annual_leave REAL NOT NULL DEFAULT 0,
  judgment TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (closure_id) REFERENCES attendance_closures(id) ON DELETE CASCADE,
  UNIQUE(closure_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_monthly_summaries_v3_route_month
ON attendance_monthly_summaries_v3(route, month, employee_id);
