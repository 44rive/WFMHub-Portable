ALTER TABLE meta.source_file
ADD COLUMN source_variant VARCHAR;

CREATE INDEX IF NOT EXISTS idx_source_file_family_variant
ON meta.source_file(source_family, source_variant, active, status, modified_at);

ALTER TABLE mart.agent_pcs_day
ADD COLUMN transferred_legs BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_status_calls BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_participation_responses BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_participation_rate DOUBLE;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_invalid_responses BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_status_blank_responses BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.agent_pcs_day
ADD COLUMN pcs_response_without_status BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN shift_state VARCHAR;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN call_action VARCHAR;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN requires_call BOOLEAN NOT NULL DEFAULT false;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN is_provisional BOOLEAN NOT NULL DEFAULT false;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN evaluation_as_of TIMESTAMP;

ALTER TABLE mart.service_interval
ADD COLUMN sl_target DOUBLE;

ALTER TABLE mart.service_interval
ADD COLUMN sl_state VARCHAR;

CREATE TABLE IF NOT EXISTS mart.staffing_interval (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    interval_end TIMESTAMP NOT NULL,
    lob VARCHAR NOT NULL,
    language VARCHAR NOT NULL,
    scheduled_agents BIGINT NOT NULL,
    observed_agents BIGINT NOT NULL,
    productive_agents BIGINT NOT NULL,
    auxiliary_agents BIGINT NOT NULL,
    scheduled_fte DOUBLE NOT NULL,
    elapsed_scheduled_fte DOUBLE NOT NULL,
    observed_fte DOUBLE NOT NULL,
    productive_fte DOUBLE NOT NULL,
    staffing_variance_fte DOUBLE,
    staffing_gap_fte DOUBLE,
    staffing_state VARCHAR NOT NULL,
    evidence_basis VARCHAR NOT NULL,
    evaluation_as_of TIMESTAMP NOT NULL,
    PRIMARY KEY (business_date, interval_start, lob, language)
);

CREATE TABLE IF NOT EXISTS mart.shift_timeline_segment (
    segment_key VARCHAR PRIMARY KEY,
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
    segment_start TIMESTAMP NOT NULL,
    segment_end TIMESTAMP NOT NULL,
    segment_minutes BIGINT NOT NULL,
    planned_state VARCHAR NOT NULL,
    actual_status VARCHAR,
    actual_category VARCHAR NOT NULL,
    mismatch_type VARCHAR NOT NULL,
    is_gap BOOLEAN NOT NULL,
    observed_source VARCHAR NOT NULL,
    source_file VARCHAR,
    evaluation_as_of TIMESTAMP NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_shift_timeline_date_agent
ON mart.shift_timeline_segment(business_date, agent_id, segment_start);

CREATE TABLE IF NOT EXISTS mart.correction_residual_segment (
    residual_id VARCHAR PRIMARY KEY,
    correction_id VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    residual_start TIMESTAMP NOT NULL,
    residual_end TIMESTAMP NOT NULL,
    residual_minutes BIGINT NOT NULL,
    suggested_activity VARCHAR,
    observed_source VARCHAR,
    source_file VARCHAR,
    verint_reconciliation VARCHAR NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_correction_residual_date_agent
ON mart.correction_residual_segment(business_date, agent_id, residual_start);

CREATE TABLE IF NOT EXISTS mart.verint_final_absence_event (
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
    event_start TIMESTAMP NOT NULL,
    event_end TIMESTAMP NOT NULL,
    minutes BIGINT NOT NULL,
    hours DOUBLE NOT NULL,
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

CREATE INDEX IF NOT EXISTS idx_verint_final_absence_event_date_agent
ON mart.verint_final_absence_event(business_date, agent_id, event_start);

CREATE TABLE IF NOT EXISTS mart.verint_final_absence_agent_day (
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
    planned_net_minutes BIGINT NOT NULL,
    final_absence_minutes BIGINT NOT NULL,
    final_vacation_minutes BIGINT NOT NULL,
    final_unpaid_minutes BIGINT NOT NULL,
    final_shrinkage_minutes BIGINT NOT NULL,
    final_unmapped_minutes BIGINT NOT NULL,
    final_absence_hours DOUBLE NOT NULL,
    final_absence_rate DOUBLE,
    final_absence_day BOOLEAN NOT NULL,
    final_ledger_status VARCHAR NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_verint_final_absence_day_date_agent
ON mart.verint_final_absence_agent_day(business_date, agent_id);
