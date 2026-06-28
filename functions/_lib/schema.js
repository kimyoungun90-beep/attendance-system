const CLOSURE_COLUMNS = {
  unexpected_count: "INTEGER NOT NULL DEFAULT 0",
  unexpected_people: "INTEGER NOT NULL DEFAULT 0",
  mismatch_count: "INTEGER NOT NULL DEFAULT 0",
  mismatch_people: "INTEGER NOT NULL DEFAULT 0",
  dayoff_excess_people: "INTEGER NOT NULL DEFAULT 0",
  substitute_shortage_people: "INTEGER NOT NULL DEFAULT 0",
  compensation_shortage_people: "INTEGER NOT NULL DEFAULT 0",
  annual_leave_people: "INTEGER NOT NULL DEFAULT 0",
};

const FACT_COLUMNS = {
  compensation_leave_used: "REAL NOT NULL DEFAULT 0",
  compensation_needed: "REAL NOT NULL DEFAULT 0",
  compensation_events_json: "TEXT NOT NULL DEFAULT '[]'",
  worked_dates_json: "TEXT NOT NULL DEFAULT '[]'",
  occurrence_substitute_dates_json: "TEXT NOT NULL DEFAULT '[]'",
  base_allowance_raw: "REAL NOT NULL DEFAULT 0",
  occurrence_rest_days: "REAL NOT NULL DEFAULT 0",
  occurrence_rest_allowances_json: "TEXT NOT NULL DEFAULT '[]'",
  daily_statuses_json: "TEXT NOT NULL DEFAULT '[]'",
  evidence_dates_json: "TEXT NOT NULL DEFAULT '[]'",
};


const GRANT_COLUMNS = {
  occurrence_date: "TEXT",
};
const SUMMARY_COLUMNS = {
  compensation_needed: "REAL NOT NULL DEFAULT 0",
  available_compensation: "REAL NOT NULL DEFAULT 0",
  compensation_applied: "REAL NOT NULL DEFAULT 0",
  remaining_compensation: "REAL NOT NULL DEFAULT 0",
  expired_compensation: "REAL NOT NULL DEFAULT 0",
  compensation_shortage: "REAL NOT NULL DEFAULT 0",
  compensation_judgment: "TEXT",
};

let initialized = false;
let initializing = null;

export async function ensureSchema(db) {
  if (initialized) return;
  if (initializing) return initializing;
  initializing = initialize(db).then(() => { initialized = true; }).finally(() => { initializing = null; });
  return initializing;
}

async function initialize(db) {
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS attendance_closures (
      id TEXT PRIMARY KEY,
      company TEXT NOT NULL CHECK (company IN ('homeplus', 'electroland')),
      month TEXT NOT NULL,
      cutoff_date TEXT NOT NULL,
      check_mode TEXT NOT NULL DEFAULT 'auto',
      plan_file_name TEXT,
      attendance_file_name TEXT,
      plan_people INTEGER NOT NULL DEFAULT 0,
      attendance_people INTEGER NOT NULL DEFAULT 0,
      matched_people INTEGER NOT NULL DEFAULT 0,
      match_rate REAL NOT NULL DEFAULT 0,
      missing_count INTEGER NOT NULL DEFAULT 0,
      missing_people INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  await ensureColumns(db, "attendance_closures", CLOSURE_COLUMNS);

  await runStaticBatch(db, [
    `CREATE TABLE IF NOT EXISTS attendance_missing_items (
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
      UNIQUE(closure_id, employee_id, missing_date)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_issue_items (
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
      UNIQUE(closure_id, issue_type, employee_id, issue_date)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_mismatch_items (
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
      UNIQUE(closure_id, employee_id, issue_date)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_monthly_employee_facts (
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
      UNIQUE(closure_id, employee_id)
    )`,
    `CREATE TABLE IF NOT EXISTS route_substitute_grants (
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
    )`,
    `CREATE TABLE IF NOT EXISTS route_substitute_allocations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      grant_id TEXT NOT NULL,
      closure_id TEXT NOT NULL,
      route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
      employee_id TEXT NOT NULL,
      month TEXT NOT NULL,
      used_days REAL NOT NULL CHECK (used_days > 0),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(grant_id, closure_id, employee_id)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_monthly_summaries_v3 (
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
      UNIQUE(closure_id, employee_id)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_leave_grants_v5 (
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
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_leave_allocations_v5 (
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
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_archive_files (
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
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_archive_file_chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_id TEXT NOT NULL,
      chunk_index INTEGER NOT NULL,
      chunk_blob BLOB NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(file_id, chunk_index)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_workforce_uploads (
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
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_workforce_file_chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      upload_id TEXT NOT NULL,
      chunk_index INTEGER NOT NULL,
      chunk_blob BLOB NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(upload_id, chunk_index)
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_portal_ids (
      employee_id TEXT PRIMARY KEY,
      portal_id TEXT NOT NULL,
      employee_name TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`,
    `CREATE TABLE IF NOT EXISTS attendance_workforce_members (
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
    )`,
    `CREATE TABLE IF NOT EXISTS annual_leave_baseline_uploads (
      route TEXT PRIMARY KEY CHECK (route IN ('homeplus', 'electroland')),
      baseline_date TEXT NOT NULL,
      file_name TEXT,
      employee_count INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`,
    `CREATE TABLE IF NOT EXISTS annual_leave_employees (
      route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
      employee_id TEXT NOT NULL,
      employee_name TEXT NOT NULL,
      regional_manager TEXT, manager TEXT, region1 TEXT, region2 TEXT,
      store_code TEXT, store_name TEXT, portal_id TEXT,
      hire_date TEXT, basis_hire_date TEXT,
      policy_type TEXT NOT NULL CHECK (policy_type IN ('jan1', 'anniversary')),
      under_one_year INTEGER NOT NULL DEFAULT 0,
      baseline_date TEXT NOT NULL,
      baseline_granted REAL NOT NULL DEFAULT 0,
      baseline_used REAL NOT NULL DEFAULT 0,
      baseline_remaining REAL NOT NULL DEFAULT 0,
      cycle_start TEXT, cycle_end TEXT,
      first_promotion_date TEXT, second_promotion_date TEXT, expiry_date TEXT,
      termination_date TEXT, note TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (route, employee_id)
    )`,
    `CREATE TABLE IF NOT EXISTS annual_leave_monthly_uploads (
      route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
      month TEXT NOT NULL, file_name TEXT,
      row_count INTEGER NOT NULL DEFAULT 0,
      approved_days REAL NOT NULL DEFAULT 0,
      rejected_days REAL NOT NULL DEFAULT 0,
      pending_days REAL NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (route, month)
    )`,
    `CREATE TABLE IF NOT EXISTS annual_leave_applications (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      route TEXT NOT NULL CHECK (route IN ('homeplus', 'electroland')),
      month TEXT NOT NULL, employee_id TEXT NOT NULL, employee_name TEXT,
      leave_date TEXT NOT NULL, days REAL NOT NULL DEFAULT 0,
      status TEXT, leave_type TEXT, application_date TEXT, note TEXT,
      source_index INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )`,
  ]);

  await ensureColumns(db, "attendance_monthly_employee_facts", FACT_COLUMNS);
  await ensureColumns(db, "attendance_monthly_summaries_v3", SUMMARY_COLUMNS);
  await ensureColumns(db, "attendance_leave_grants_v5", GRANT_COLUMNS);

  await runStaticBatch(db, [
    `DELETE FROM attendance_closures
      WHERE rowid NOT IN (SELECT MAX(rowid) FROM attendance_closures GROUP BY company, month)`,
    `CREATE UNIQUE INDEX IF NOT EXISTS ux_attendance_closures_route_month ON attendance_closures(company, month)`,
    `CREATE INDEX IF NOT EXISTS idx_attendance_closures_month_company ON attendance_closures(month, company, created_at DESC)`,
    `CREATE INDEX IF NOT EXISTS idx_missing_items_closure ON attendance_missing_items(closure_id, missing_date, employee_id)`,
    `CREATE INDEX IF NOT EXISTS idx_attendance_issue_items_closure ON attendance_issue_items(closure_id, issue_type, issue_date, employee_id)`,
    `CREATE INDEX IF NOT EXISTS idx_attendance_mismatch_items_closure ON attendance_mismatch_items(closure_id, issue_date, employee_id)`,
    `CREATE INDEX IF NOT EXISTS idx_monthly_employee_facts_route_month ON attendance_monthly_employee_facts(route, month, employee_id)`,
    `CREATE INDEX IF NOT EXISTS idx_route_substitute_grants_period ON route_substitute_grants(route, grant_month, valid_to, valid_from)`,
    `CREATE INDEX IF NOT EXISTS idx_route_substitute_allocations_employee ON route_substitute_allocations(route, employee_id, month, grant_id)`,
    `CREATE INDEX IF NOT EXISTS idx_monthly_summaries_v3_route_month ON attendance_monthly_summaries_v3(route, month, employee_id)`,
    `UPDATE attendance_leave_grants_v5
       SET occurrence_date = COALESCE(NULLIF(criterion_date, ''), grant_month || '-01')
       WHERE occurrence_date IS NULL OR TRIM(occurrence_date) = ''`,
    `CREATE INDEX IF NOT EXISTS idx_leave_grants_v5_route_month ON attendance_leave_grants_v5(route, grant_month, grant_type, grant_scope)`,
    `CREATE INDEX IF NOT EXISTS idx_leave_allocations_v5_employee ON attendance_leave_allocations_v5(route, grant_type, employee_id, month, grant_id)`,
    `CREATE INDEX IF NOT EXISTS idx_attendance_archive_files_route_month ON attendance_archive_files(route, month, created_at DESC)`,
    `CREATE INDEX IF NOT EXISTS idx_attendance_archive_file_chunks_file ON attendance_archive_file_chunks(file_id, chunk_index)`,
    `CREATE INDEX IF NOT EXISTS idx_workforce_uploads_month ON attendance_workforce_uploads(month DESC)`,
    `CREATE INDEX IF NOT EXISTS idx_workforce_chunks_upload ON attendance_workforce_file_chunks(upload_id, chunk_index)`,
    `CREATE INDEX IF NOT EXISTS idx_workforce_members_month_route ON attendance_workforce_members(month, route, employee_id, store_code)`,
    `CREATE INDEX IF NOT EXISTS idx_workforce_members_upload ON attendance_workforce_members(upload_id)`,
    `CREATE INDEX IF NOT EXISTS idx_annual_leave_employees_route ON annual_leave_employees(route, under_one_year, employee_id)`,
    `CREATE INDEX IF NOT EXISTS idx_annual_leave_applications_route_month ON annual_leave_applications(route, month, employee_id, leave_date)`,
    `CREATE INDEX IF NOT EXISTS idx_annual_leave_applications_employee_date ON annual_leave_applications(route, employee_id, leave_date)`,
    `INSERT OR IGNORE INTO attendance_leave_grants_v5
      (id, route, grant_type, grant_scope, grant_month, granted_days, valid_from, valid_to,
       eligibility_mode, criterion_date, employee_id, excluded_employee_ids_json, reason, note, created_at, updated_at)
      SELECT id, route, 'substitute', 'route', grant_month, granted_days, valid_from, valid_to,
             'all', NULL, NULL, '[]', reason, note, created_at, updated_at
      FROM route_substitute_grants`,
  ]);
}

async function ensureColumns(db, tableName, definitions) {
  const columns = await db.prepare(`PRAGMA table_info(${tableName})`).all();
  const existing = new Set((columns.results || []).map((row) => row.name));
  for (const [name, definition] of Object.entries(definitions)) {
    if (existing.has(name)) continue;
    try {
      await db.prepare(`ALTER TABLE ${tableName} ADD COLUMN ${name} ${definition}`).run();
    } catch (error) {
      if (!String(error?.message || error).toLowerCase().includes("duplicate column")) throw error;
    }
  }
}

async function runStaticBatch(db, sqlList) {
  const statements = sqlList.map((sql) => db.prepare(sql));
  for (let index = 0; index < statements.length; index += 40) {
    await db.batch(statements.slice(index, index + 40));
  }
}
