PRAGMA foreign_keys = ON;

-- v5: 대체휴무·보상휴가 통합 부여와 재계산 구조
-- 실제 배포본은 API 최초 호출 때 누락 구조를 자동 생성하므로 수동 실행은 필수가 아닙니다.
-- 기존 D1 구조 확인 또는 신규 DB 구성 참고용입니다.

ALTER TABLE attendance_closures ADD COLUMN compensation_shortage_people INTEGER NOT NULL DEFAULT 0;

ALTER TABLE attendance_monthly_employee_facts ADD COLUMN compensation_leave_used REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN compensation_needed REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN compensation_events_json TEXT NOT NULL DEFAULT '[]';
ALTER TABLE attendance_monthly_employee_facts ADD COLUMN worked_dates_json TEXT NOT NULL DEFAULT '[]';

ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN compensation_needed REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN available_compensation REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN compensation_applied REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN remaining_compensation REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN expired_compensation REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN compensation_shortage REAL NOT NULL DEFAULT 0;
ALTER TABLE attendance_monthly_summaries_v3 ADD COLUMN compensation_judgment TEXT;

CREATE TABLE IF NOT EXISTS attendance_leave_grants_v5 (
  id TEXT PRIMARY KEY,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  grant_type TEXT NOT NULL CHECK (grant_type IN ('substitute', 'compensation')),
  grant_scope TEXT NOT NULL CHECK (grant_scope IN ('route', 'employee')),
  grant_month TEXT NOT NULL,
  granted_days REAL NOT NULL CHECK (granted_days > 0),
  valid_from TEXT NOT NULL,
  valid_to TEXT NOT NULL,
  eligibility_mode TEXT NOT NULL DEFAULT 'all' CHECK (eligibility_mode IN ('all', 'worked_on_date')),
  criterion_date TEXT,
  employee_id TEXT,
  excluded_employee_ids_json TEXT NOT NULL DEFAULT '[]',
  reason TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS attendance_leave_allocations_v5 (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  grant_id TEXT NOT NULL,
  closure_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  grant_type TEXT NOT NULL CHECK (grant_type IN ('substitute', 'compensation')),
  employee_id TEXT NOT NULL,
  month TEXT NOT NULL,
  used_days REAL NOT NULL CHECK (used_days > 0),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(grant_id, closure_id, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_leave_grants_v5_route_month
ON attendance_leave_grants_v5(route, grant_month, grant_type, grant_scope);

CREATE INDEX IF NOT EXISTS idx_leave_allocations_v5_employee
ON attendance_leave_allocations_v5(route, grant_type, employee_id, month, grant_id);

-- 구버전 경로 공통 대체휴무가 있으면 v5 통합 장부로 이관합니다.
INSERT OR IGNORE INTO attendance_leave_grants_v5
(id, route, grant_type, grant_scope, grant_month, granted_days, valid_from, valid_to,
 eligibility_mode, criterion_date, employee_id, excluded_employee_ids_json,
 reason, note, created_at, updated_at)
SELECT id, route, 'substitute', 'route', grant_month, granted_days, valid_from, valid_to,
       'all', NULL, NULL, '[]', reason, note, created_at, updated_at
FROM route_substitute_grants;
