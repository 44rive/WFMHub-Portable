CREATE TABLE IF NOT EXISTS core.pcs_coaching_action (
    coaching_key VARCHAR PRIMARY KEY,
    coaching_status VARCHAR NOT NULL DEFAULT 'PENDING',
    coaching_date DATE,
    coach VARCHAR,
    coaching_comment VARCHAR,
    updated_at TIMESTAMP NOT NULL,
    imported_from VARCHAR
);

CREATE INDEX IF NOT EXISTS idx_pcs_coaching_status
ON core.pcs_coaching_action(coaching_status, coaching_date);

CREATE TRIGGER IF NOT EXISTS trg_pcs_coaching_key_not_null
BEFORE INSERT ON core.pcs_coaching_action WHEN NEW.coaching_key IS NULL
BEGIN SELECT RAISE(ABORT, 'core.pcs_coaching_action.coaching_key cannot be NULL'); END;
