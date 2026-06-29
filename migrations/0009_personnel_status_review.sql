-- v32 월별 인력현황 ↔ 연차대장 확인 요청 및 수기 상태 관리
CREATE TABLE IF NOT EXISTS attendance_personnel_status_overrides (
  month TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  employee_id TEXT NOT NULL,
  employee_name TEXT,
  issue_type TEXT,
  personnel_status TEXT NOT NULL DEFAULT '확인 요청',
  effective_from TEXT,
  effective_to TEXT,
  destination_route TEXT,
  note TEXT,
  source_type TEXT NOT NULL DEFAULT 'manual' CHECK (source_type IN ('manual', 'workforce', 'evidence')),
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  PRIMARY KEY (month, route, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_personnel_status_month_route
  ON attendance_personnel_status_overrides(month, route, personnel_status, employee_id);
