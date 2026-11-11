PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS captures (
    capture_id TEXT PRIMARY KEY,
    digest TEXT NOT NULL,
    task_group TEXT NOT NULL,
    started_at TEXT NOT NULL,
    finished_at TEXT NOT NULL,
    received_at TEXT NOT NULL,
    status TEXT NOT NULL CHECK (status IN ('complete', 'incomplete', 'failed')),
    source_total INTEGER,
    error TEXT NOT NULL,
    payload TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS captures_group_time ON captures(task_group, finished_at);
CREATE TABLE IF NOT EXISTS task_rows (
    capture_id TEXT NOT NULL REFERENCES captures(capture_id),
    row_index INTEGER NOT NULL,
    case_number TEXT NOT NULL,
    task_type TEXT NOT NULL,
    priority TEXT NOT NULL,
    language TEXT NOT NULL,
    scheduled_for TEXT,
    scheduled_raw TEXT NOT NULL,
    assigned_to TEXT NOT NULL,
    PRIMARY KEY (capture_id, row_index)
);
CREATE TABLE IF NOT EXISTS inbox_events (
    id INTEGER PRIMARY KEY,
    occurred_at TEXT NOT NULL,
    filename TEXT NOT NULL,
    status TEXT NOT NULL,
    message TEXT NOT NULL
);
PRAGMA user_version = 1;
