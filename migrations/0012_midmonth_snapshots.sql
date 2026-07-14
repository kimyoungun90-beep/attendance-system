-- v61: 중간 확인 저장 이력 관리
CREATE TABLE IF NOT EXISTS attendance_midmonth_snapshots (
  id TEXT PRIMARY KEY,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  cutoff_date TEXT NOT NULL,
  snapshot_name TEXT,
  status TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active', 'deleted')),
  employee_count INTEGER NOT NULL DEFAULT 0,
  daily_count INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(route, month, cutoff_date)
);

CREATE TABLE IF NOT EXISTS attendance_midmonth_snapshot_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  snapshot_id TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  employee_id TEXT NOT NULL,
  employee_name TEXT,
  store TEXT,
  status_date TEXT NOT NULL,
  plan_status TEXT,
  has_clock_in INTEGER NOT NULL DEFAULT 0,
  actual_status TEXT,
  actual_in TEXT,
  changed_in TEXT,
  display_status TEXT,
  evidenced INTEGER NOT NULL DEFAULT 0,
  source TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(snapshot_id, employee_id, status_date)
);

CREATE INDEX IF NOT EXISTS idx_midmonth_snapshots_route_month
  ON attendance_midmonth_snapshots(route, month, status, cutoff_date DESC, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_midmonth_snapshot_items_snapshot
  ON attendance_midmonth_snapshot_items(snapshot_id, status_date, employee_id);
