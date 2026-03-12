-- SheetSync D1 Schema
-- Run with: npx wrangler d1 execute sheetsync-licenses --file=schema.sql

-- Individual user licenses
CREATE TABLE IF NOT EXISTS user_licenses (
    email            TEXT PRIMARY KEY,
    domain           TEXT,
    plan             TEXT    NOT NULL DEFAULT 'free',
    status           TEXT    NOT NULL DEFAULT 'none',
    allowed          INTEGER NOT NULL DEFAULT 0,
    customer_id      TEXT,
    subscription_id  TEXT,
    price_id         TEXT,
    amount           REAL,
    renews_at        TEXT,
    last_updated     TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- Domain-wide licenses
CREATE TABLE IF NOT EXISTS domain_licenses (
    domain           TEXT PRIMARY KEY,
    email            TEXT,   -- who purchased
    plan             TEXT    NOT NULL DEFAULT 'free',
    status           TEXT    NOT NULL DEFAULT 'none',
    allowed          INTEGER NOT NULL DEFAULT 0,
    customer_id      TEXT,
    subscription_id  TEXT,
    price_id         TEXT,
    amount           REAL,
    renews_at        TEXT,
    seats            INTEGER,
    last_updated     TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- Login activity (upserted on every license check)
CREATE TABLE IF NOT EXISTS login_events (
    email        TEXT PRIMARY KEY,
    first_seen   TEXT NOT NULL,
    last_seen    TEXT NOT NULL,
    visit_count  INTEGER NOT NULL DEFAULT 1,
    plan         TEXT
);

-- Sales history
CREATE TABLE IF NOT EXISTS sales (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at       TEXT NOT NULL,
    email            TEXT NOT NULL,
    plan             TEXT,
    domain           TEXT,
    customer_id      TEXT,
    subscription_id  TEXT,
    amount           REAL
);

-- Jira proxy circuit-breaker call counts (per domain per day)
CREATE TABLE IF NOT EXISTS domain_stats (
    hostname    TEXT NOT NULL,
    date        TEXT NOT NULL,
    call_count  INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (hostname, date)
);

-- Indexes
CREATE INDEX IF NOT EXISTS idx_user_licenses_domain   ON user_licenses(domain);
CREATE INDEX IF NOT EXISTS idx_login_events_last_seen ON login_events(last_seen);
CREATE INDEX IF NOT EXISTS idx_sales_created_at       ON sales(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_sales_email            ON sales(email);
CREATE INDEX IF NOT EXISTS idx_domain_stats_date      ON domain_stats(date);
