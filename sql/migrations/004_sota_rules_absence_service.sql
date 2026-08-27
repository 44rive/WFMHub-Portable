ALTER TABLE raw.queue_actual ADD COLUMN abandoned_20s DOUBLE;

CREATE TABLE IF NOT EXISTS meta.rule_application (
    run_id VARCHAR PRIMARY KEY,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL,
    rule_file VARCHAR NOT NULL,
    effective_from DATE NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS mart.absence_event (
    event_key VARCHAR PRIMARY KEY,
    agent_day_key VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    market VARCHAR,
    language VARCHAR,
    location VARCHAR,
    activity VARCHAR,
    category VARCHAR NOT NULL,
    event_start TIMESTAMP,
    event_end TIMESTAMP,
    minutes BIGINT NOT NULL,
    hours DOUBLE NOT NULL,
    planned BOOLEAN NOT NULL,
    working BOOLEAN NOT NULL,
    counts_as_absence BOOLEAN NOT NULL,
    counts_as_vacation BOOLEAN NOT NULL,
    counts_as_unpaid BOOLEAN NOT NULL,
    counts_as_shrinkage BOOLEAN NOT NULL,
    mapped BOOLEAN NOT NULL,
    evidence_type VARCHAR NOT NULL,
    source_file VARCHAR,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE TABLE IF NOT EXISTS mart.absence_agent_day (
    agent_day_key VARCHAR PRIMARY KEY,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    market VARCHAR,
    language VARCHAR,
    location VARCHAR,
    scheduled_minutes BIGINT NOT NULL,
    break_minutes BIGINT NOT NULL,
    lunch_minutes BIGINT NOT NULL,
    planned_net_minutes BIGINT NOT NULL,
    production_minutes BIGINT NOT NULL,
    absence_minutes BIGINT NOT NULL,
    vacation_minutes BIGINT NOT NULL,
    unpaid_minutes BIGINT NOT NULL,
    shrinkage_minutes BIGINT NOT NULL,
    late_minutes BIGINT NOT NULL,
    early_leave_minutes BIGINT NOT NULL,
    no_show_minutes BIGINT NOT NULL,
    unmapped_minutes BIGINT NOT NULL,
    absence_rate DOUBLE,
    vacation_rate DOUBLE,
    shrinkage_rate DOUBLE,
    absence_day BOOLEAN NOT NULL,
    absence_spell VARCHAR,
    absence_spells BIGINT NOT NULL,
    absence_days DOUBLE NOT NULL,
    bradford_factor DOUBLE NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE TABLE IF NOT EXISTS mart.service_interval (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    hour_start TIMESTAMP,
    source_system VARCHAR NOT NULL,
    queue VARCHAR,
    business_partner VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    offered DOUBLE,
    answered DOUBLE,
    abandoned DOUBLE,
    short_abandoned DOUBLE,
    answered_within_target DOUBLE,
    handled_seconds DOUBLE,
    sl_gross DOUBLE,
    sl_adjusted DOUBLE,
    sl_profile VARCHAR NOT NULL,
    service_level DOUBLE,
    service_availability DOUBLE,
    abandon_rate DOUBLE,
    aht_seconds DOUBLE,
    source_file VARCHAR,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_absence_event_date_agent
ON mart.absence_event(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_absence_day_date_agent
ON mart.absence_agent_day(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_service_interval_date_scope
ON mart.service_interval(business_date, source_system, lob, language);

CREATE TRIGGER IF NOT EXISTS trg_absence_event_key_not_null
BEFORE INSERT ON mart.absence_event WHEN NEW.event_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart.absence_event.event_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_absence_day_key_not_null
BEFORE INSERT ON mart.absence_agent_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart.absence_agent_day.agent_day_key cannot be NULL'); END;
