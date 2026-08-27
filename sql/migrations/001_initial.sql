CREATE SCHEMA IF NOT EXISTS meta;
CREATE SCHEMA IF NOT EXISTS raw;
CREATE SCHEMA IF NOT EXISTS core;
CREATE SCHEMA IF NOT EXISTS mart;

CREATE TABLE IF NOT EXISTS meta.schema_migration (
    version VARCHAR PRIMARY KEY,
    applied_at TIMESTAMP NOT NULL DEFAULT current_timestamp
);

CREATE TABLE IF NOT EXISTS meta.refresh_run (
    run_id VARCHAR PRIMARY KEY,
    started_at TIMESTAMP NOT NULL,
    finished_at TIMESTAMP,
    requested_start DATE,
    requested_end DATE,
    status VARCHAR NOT NULL,
    files_loaded BIGINT DEFAULT 0,
    files_skipped BIGINT DEFAULT 0,
    files_failed BIGINT DEFAULT 0,
    details VARCHAR
);

CREATE TABLE IF NOT EXISTS meta.source_file (
    file_id VARCHAR PRIMARY KEY,
    source_family VARCHAR NOT NULL,
    source_path VARCHAR NOT NULL,
    file_name VARCHAR NOT NULL,
    sha256 VARCHAR NOT NULL,
    size_bytes BIGINT NOT NULL,
    modified_at TIMESTAMP,
    discovered_at TIMESTAMP NOT NULL,
    loaded_at TIMESTAMP,
    active BOOLEAN NOT NULL DEFAULT false,
    status VARCHAR NOT NULL,
    row_count BIGINT NOT NULL DEFAULT 0,
    rejected_count BIGINT NOT NULL DEFAULT 0,
    error_message VARCHAR
);

CREATE INDEX IF NOT EXISTS idx_source_file_family_path
ON meta.source_file(source_family, source_path, active);

CREATE TABLE IF NOT EXISTS meta.quality_issue (
    issue_id VARCHAR PRIMARY KEY,
    run_id VARCHAR,
    detected_at TIMESTAMP NOT NULL,
    source_family VARCHAR,
    source_file VARCHAR,
    business_date DATE,
    agent_id VARCHAR,
    issue_type VARCHAR NOT NULL,
    severity VARCHAR NOT NULL,
    details VARCHAR NOT NULL
);

CREATE TABLE IF NOT EXISTS raw.fte_agent (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    agent_id VARCHAR,
    employment_status VARCHAR,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    market VARCHAR,
    language VARCHAR,
    location VARCHAR,
    city VARCHAR,
    fte DOUBLE,
    end_date DATE
);

CREATE TABLE IF NOT EXISTS raw.schedule_shift (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    schedule_date DATE,
    agent_id_raw VARCHAR,
    agent_id VARCHAR,
    agent_name VARCHAR,
    scheduling_period VARCHAR,
    shift_assignment VARCHAR,
    assignment VARCHAR,
    assignment_type VARCHAR,
    scheduled_start TIMESTAMP,
    scheduled_end TIMESTAMP,
    shift_events VARCHAR,
    parse_ok BOOLEAN
);

CREATE TABLE IF NOT EXISTS raw.schedule_event (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    event_index BIGINT NOT NULL,
    schedule_date DATE,
    agent_id VARCHAR,
    agent_name VARCHAR,
    activity VARCHAR,
    activity_type VARCHAR,
    event_start TIMESTAMP,
    event_end TIMESTAMP,
    parse_ok BOOLEAN
);

CREATE TABLE IF NOT EXISTS raw.lilo (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    extract_date DATE,
    agent_id VARCHAR,
    agent_name VARCHAR,
    first_login TIMESTAMP,
    raw_last_logout TIMESTAMP,
    last_logout TIMESTAMP,
    overnight_adjusted BOOLEAN
);

CREATE TABLE IF NOT EXISTS raw.agent_status (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    serial_number VARCHAR,
    extract_date DATE,
    agent_id VARCHAR,
    agent_name VARCHAR,
    status VARCHAR,
    actual_category VARCHAR,
    status_start TIMESTAMP,
    status_end TIMESTAMP,
    duration_seconds BIGINT,
    queue VARCHAR
);

CREATE TABLE IF NOT EXISTS raw.forecast_interval (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    queue_name VARCHAR,
    business_date DATE,
    interval_time TIME,
    interval_minutes BIGINT,
    interval_start TIMESTAMP,
    volume_forecast DOUBLE,
    abandons_forecast DOUBLE,
    sl_forecast DOUBLE,
    sl_required DOUBLE,
    aht_forecast_seconds DOUBLE,
    headcount_forecast DOUBLE,
    net_staffing_forecast DOUBLE,
    fte_forecast DOUBLE,
    fte_required DOUBLE
);

CREATE TABLE IF NOT EXISTS raw.queue_actual (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    source_system VARCHAR NOT NULL,
    business_date DATE,
    interval_time TIME,
    interval_start TIMESTAMP,
    hour_start TIMESTAMP,
    language VARCHAR,
    queue_id VARCHAR,
    queue VARCHAR,
    business_partner VARCHAR,
    lob VARCHAR,
    offered DOUBLE,
    answered DOUBLE,
    abandoned DOUBLE,
    short_calls DOUBLE,
    answered_15s DOUBLE,
    answered_20s DOUBLE,
    answered_30s DOUBLE,
    asa_seconds DOUBLE,
    aht_seconds DOUBLE
);

CREATE TABLE IF NOT EXISTS core.correction_action (
    correction_id VARCHAR PRIMARY KEY,
    confirmed_activity VARCHAR,
    validation_status VARCHAR,
    owner VARCHAR,
    comment VARCHAR,
    injected_date DATE,
    updated_at TIMESTAMP NOT NULL,
    imported_from VARCHAR
);

CREATE TABLE IF NOT EXISTS core.dim_agent (
    agent_id VARCHAR PRIMARY KEY,
    canonical_name VARCHAR,
    employment_status VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    market VARCHAR,
    language VARCHAR,
    location VARCHAR,
    city VARCHAR,
    fte DOUBLE,
    match_method VARCHAR
);

CREATE TABLE IF NOT EXISTS mart.attendance_agent_day (
    agent_day_key VARCHAR PRIMARY KEY,
    business_date DATE,
    agent_id VARCHAR,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    market VARCHAR,
    language VARCHAR,
    location VARCHAR,
    scheduled_start TIMESTAMP,
    scheduled_end TIMESTAMP,
    scheduled_minutes BIGINT,
    assignment VARCHAR,
    assignment_type VARCHAR,
    planned_absence_minutes BIGINT,
    first_login TIMESTAMP,
    last_logout TIMESTAMP,
    source_loaded BOOLEAN,
    lilo_row_present BOOLEAN,
    seen_in_lilo BOOLEAN,
    raw_late_minutes BIGINT,
    raw_early_leave_minutes BIGINT,
    uncoded_late_minutes BIGINT,
    uncoded_early_leave_minutes BIGINT,
    no_show_minutes BIGINT,
    worked_span_minutes BIGINT,
    attendance_result VARCHAR,
    attendance_percent DOUBLE,
    schedule_source VARCHAR,
    lilo_source VARCHAR
);

CREATE TABLE IF NOT EXISTS mart.conformance_agent_day (
    agent_day_key VARCHAR PRIMARY KEY,
    business_date DATE,
    agent_id VARCHAR,
    scheduled_minutes BIGINT,
    scheduled_net_minutes BIGINT,
    planned_absence_minutes BIGINT,
    planned_lunch_minutes BIGINT,
    planned_break_minutes BIGINT,
    productive_minutes BIGINT,
    auxiliary_minutes BIGINT,
    break_minutes BIGINT,
    lunch_minutes BIGINT,
    unavailable_minutes BIGINT,
    logged_off_minutes BIGINT,
    status_covered_minutes BIGINT,
    status_coverage_percent DOUBLE,
    login_span_minutes BIGINT,
    measurement_basis VARCHAR,
    worked_minutes BIGINT,
    conformance_percent DOUBLE,
    break_overrun_minutes BIGINT,
    lunch_overrun_minutes BIGINT,
    unexplained_minutes BIGINT
);

CREATE TABLE IF NOT EXISTS mart.correction_candidate (
    correction_id VARCHAR PRIMARY KEY,
    business_date DATE,
    agent_id VARCHAR,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    scheduled_start TIMESTAMP,
    scheduled_end TIMESTAMP,
    first_login TIMESTAMP,
    last_logout TIMESTAMP,
    priority BIGINT,
    detected_issue VARCHAR,
    gap_start TIMESTAMP,
    gap_end TIMESTAMP,
    gap_minutes BIGINT,
    confidence VARCHAR,
    suggested_activity VARCHAR,
    source_file VARCHAR,
    confirmed_activity VARCHAR,
    validation_status VARCHAR,
    owner VARCHAR,
    comment VARCHAR,
    injected_date DATE
);

CREATE TABLE IF NOT EXISTS mart.rta_snapshot (
    snapshot_at TIMESTAMP,
    agent_id VARCHAR,
    agent_name VARCHAR,
    team_leader VARCHAR,
    lob VARCHAR,
    scheduled_start TIMESTAMP,
    scheduled_end TIMESTAMP,
    planned_activity VARCHAR,
    actual_status VARCHAR,
    actual_category VARCHAR,
    status_start TIMESTAMP,
    minutes_in_status BIGINT,
    rta_result VARCHAR,
    severity VARCHAR,
    freshness VARCHAR,
    source_file VARCHAR
);

CREATE TABLE IF NOT EXISTS mart.forecast_hour (
    business_date DATE,
    hour_start TIMESTAMP,
    queue_name VARCHAR,
    volume_forecast DOUBLE,
    fte_forecast DOUBLE,
    fte_required DOUBLE,
    sl_forecast DOUBLE,
    sl_required DOUBLE,
    aht_forecast_seconds DOUBLE,
    source_file VARCHAR
);

CREATE TABLE IF NOT EXISTS mart.intraday_queue_interval (
    business_date DATE,
    interval_start TIMESTAMP,
    hour_start TIMESTAMP,
    source_system VARCHAR,
    queue VARCHAR,
    business_partner VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    offered DOUBLE,
    answered DOUBLE,
    abandoned DOUBLE,
    short_calls DOUBLE,
    answered_20s DOUBLE,
    service_level_20s DOUBLE,
    abandon_rate DOUBLE,
    asa_seconds DOUBLE,
    aht_seconds DOUBLE,
    source_file VARCHAR
);

CREATE TABLE IF NOT EXISTS mart.source_health (
    source_family VARCHAR,
    expected_path VARCHAR,
    newest_file VARCHAR,
    newest_business_date DATE,
    modified_at TIMESTAMP,
    loaded_at TIMESTAMP,
    row_count BIGINT,
    rejected_count BIGINT,
    status VARCHAR,
    details VARCHAR
);
