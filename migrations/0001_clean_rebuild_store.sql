CREATE TABLE IF NOT EXISTS attendance_clean_store (
  key TEXT PRIMARY KEY,
  type TEXT NOT NULL,
  route TEXT,
  month TEXT,
  payload_json TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_attendance_clean_store_type_route_month
ON attendance_clean_store (type, route, month);
