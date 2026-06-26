-- v6 월별 인력·매장매칭 및 최종본 양식 지원
CREATE TABLE IF NOT EXISTS attendance_workforce_uploads (
  id TEXT PRIMARY KEY,
  month TEXT NOT NULL UNIQUE,
  file_name TEXT NOT NULL,
  content_type TEXT,
  size_bytes INTEGER NOT NULL DEFAULT 0,
  storage_type TEXT NOT NULL CHECK (storage_type IN ('d1', 'r2')),
  object_key TEXT,
  employee_count INTEGER NOT NULL DEFAULT 0,
  electroland_count INTEGER NOT NULL DEFAULT 0,
  homeplus_count INTEGER NOT NULL DEFAULT 0,
  portal_count INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS attendance_workforce_file_chunks (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  upload_id TEXT NOT NULL,
  chunk_index INTEGER NOT NULL,
  chunk_blob BLOB NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(upload_id, chunk_index)
);

CREATE TABLE IF NOT EXISTS attendance_portal_ids (
  employee_id TEXT PRIMARY KEY,
  portal_id TEXT NOT NULL,
  employee_name TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS attendance_workforce_members (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  upload_id TEXT NOT NULL,
  month TEXT NOT NULL,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  regional_manager TEXT,
  manager TEXT,
  region1 TEXT,
  region2 TEXT,
  store_code TEXT,
  store_name TEXT,
  portal_id TEXT,
  employee_id TEXT NOT NULL,
  employee_name TEXT NOT NULL,
  hire_date TEXT,
  group_hire_date TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(month, route, employee_id, store_code)
);

CREATE INDEX IF NOT EXISTS idx_workforce_uploads_month
  ON attendance_workforce_uploads(month DESC);
CREATE INDEX IF NOT EXISTS idx_workforce_chunks_upload
  ON attendance_workforce_file_chunks(upload_id, chunk_index);
CREATE INDEX IF NOT EXISTS idx_workforce_members_month_route
  ON attendance_workforce_members(month, route, employee_id, store_code);
CREATE INDEX IF NOT EXISTS idx_workforce_members_upload
  ON attendance_workforce_members(upload_id);
