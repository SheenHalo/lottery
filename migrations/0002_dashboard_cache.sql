CREATE TABLE IF NOT EXISTS dashboard_cache (
  cache_key TEXT PRIMARY KEY,
  report_date TEXT,
  payload TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
