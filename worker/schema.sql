CREATE TABLE IF NOT EXISTS leads (
  campaign_id INTEGER NOT NULL,
  person_id TEXT NOT NULL,
  name TEXT,
  email TEXT,
  phone TEXT,
  job_title TEXT,
  company_name TEXT,
  company_domain TEXT,
  linkedin_url TEXT,
  country TEXT,
  why_hot TEXT,
  became_hot_at TEXT,
  status TEXT NOT NULL DEFAULT 'New',
  raw_json TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (campaign_id, person_id)
);

CREATE INDEX IF NOT EXISTS idx_leads_became_hot_at ON leads (became_hot_at DESC);

CREATE TABLE IF NOT EXISTS sync_state (
  key TEXT PRIMARY KEY,
  cursor TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS settings (
  user_id TEXT PRIMARY KEY,
  instruction TEXT NOT NULL DEFAULT '',
  example TEXT NOT NULL DEFAULT '',
  signature TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL
);
