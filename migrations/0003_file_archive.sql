PRAGMA foreign_keys = ON;

-- v4 월 마감 파일 보관함.
-- 실제 배포본은 API 호출 시 누락된 테이블을 자동 생성하므로 수동 실행은 필수가 아닙니다.

CREATE TABLE IF NOT EXISTS attendance_archive_files (
  id TEXT PRIMARY KEY,
  route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
  month TEXT NOT NULL,
  file_kind TEXT NOT NULL CHECK (file_kind IN ('plan', 'attendance', 'result', 'other')),
  file_name TEXT NOT NULL,
  content_type TEXT,
  size_bytes INTEGER NOT NULL DEFAULT 0,
  storage_type TEXT NOT NULL CHECK (storage_type IN ('d1', 'r2')),
  object_key TEXT,
  file_blob BLOB,
  source_type TEXT NOT NULL DEFAULT 'manual' CHECK (source_type IN ('manual', 'closure')),
  closure_id TEXT,
  note TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS attendance_archive_file_chunks (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  file_id TEXT NOT NULL,
  chunk_index INTEGER NOT NULL,
  chunk_blob BLOB NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(file_id, chunk_index)
);

CREATE INDEX IF NOT EXISTS idx_attendance_archive_files_route_month
ON attendance_archive_files(route, month, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_attendance_archive_file_chunks_file
ON attendance_archive_file_chunks(file_id, chunk_index);
