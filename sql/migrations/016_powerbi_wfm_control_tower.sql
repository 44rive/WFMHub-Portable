-- WFM-only Power BI control tower facts.
-- Both tables are additive migrations: existing SQLite history and user state
-- remain untouched when a release is installed over an existing Hub.

CREATE TABLE IF NOT EXISTS mart.call_service_15min (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    interval_end TIMESTAMP NOT NULL,
    source_system VARCHAR NOT NULL,
    service_scope VARCHAR NOT NULL,
    comparison_scope VARCHAR NOT NULL,
    queue VARCHAR NOT NULL,
    designation VARCHAR,
    language VARCHAR,
    offered BIGINT NOT NULL,
    answered BIGINT NOT NULL,
    abandoned BIGINT NOT NULL,
    short_abandoned BIGINT NOT NULL,
    abandoned_within_target BIGINT NOT NULL,
    answered_within_target BIGINT NOT NULL,
    talk_seconds DOUBLE NOT NULL,
    hold_seconds DOUBLE NOT NULL,
    wrap_seconds DOUBLE NOT NULL,
    handled_seconds DOUBLE NOT NULL,
    service_level DOUBLE,
    service_availability DOUBLE,
    abandon_rate DOUBLE,
    aht_seconds DOUBLE,
    call_legs BIGINT NOT NULL,
    transferred_legs BIGINT NOT NULL,
    source_files VARCHAR,
    mapping_sha256 VARCHAR NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL,
    PRIMARY KEY (business_date, interval_start, comparison_scope, queue)
);

CREATE INDEX IF NOT EXISTS idx_call_service_15min_scope
ON mart.call_service_15min(business_date, interval_start, comparison_scope, service_scope);

CREATE TABLE IF NOT EXISTS mart.schedule_integrity_agent_day (
    integrity_key VARCHAR PRIMARY KEY,
    agent_day_key VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    scheduled_start TIMESTAMP NOT NULL,
    scheduled_end TIMESTAMP NOT NULL,
    observed_start TIMESTAMP,
    observed_end TIMESTAMP,
    scheduled_minutes BIGINT NOT NULL,
    observed_span_minutes BIGINT,
    start_delta_minutes BIGINT,
    end_delta_minutes BIGINT,
    displaced_minutes BIGINT NOT NULL,
    internal_gap_minutes BIGINT NOT NULL,
    internal_gap_count BIGINT NOT NULL,
    classification VARCHAR NOT NULL,
    pattern_family VARCHAR NOT NULL,
    recurrence_count BIGINT NOT NULL,
    eligible_day_count BIGINT NOT NULL,
    is_recurring BOOLEAN NOT NULL,
    requires_review BOOLEAN NOT NULL,
    evidence_basis VARCHAR NOT NULL,
    confidence VARCHAR NOT NULL,
    evaluation_as_of TIMESTAMP NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_schedule_integrity_date_agent
ON mart.schedule_integrity_agent_day(business_date, agent_id, classification);
