CREATE TABLE IF NOT EXISTS lottery_monthly_summary (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date TEXT NOT NULL,
  lottery_group TEXT NOT NULL,
  lottery_type TEXT NOT NULL,
  monthly_sales NUMERIC(18, 6),
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT uq_lottery_monthly_summary
    UNIQUE (report_date, lottery_group, lottery_type)
);

CREATE TABLE IF NOT EXISTS lottery_region_summary (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date TEXT NOT NULL,
  region_name TEXT NOT NULL,
  lottery_group TEXT NOT NULL,
  monthly_sales NUMERIC(18, 6),
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT uq_lottery_region_summary
    UNIQUE (report_date, region_name, lottery_group)
);

CREATE TABLE IF NOT EXISTS report_attachments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date TEXT NOT NULL,
  page_title TEXT NOT NULL,
  source_url TEXT,
  attachment_url TEXT,
  original_name TEXT,
  filename TEXT NOT NULL,
  r2_key TEXT NOT NULL,
  extension TEXT,
  file_size INTEGER,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT uq_report_attachments
    UNIQUE (report_date, r2_key)
);
