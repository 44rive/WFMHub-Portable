ALTER TABLE mart.attendance_agent_day ADD COLUMN actual_first_seen TIMESTAMP;
ALTER TABLE mart.attendance_agent_day ADD COLUMN actual_last_seen TIMESTAMP;
ALTER TABLE mart.attendance_agent_day ADD COLUMN actual_evidence VARCHAR;
ALTER TABLE mart.attendance_agent_day ADD COLUMN status_covered_minutes BIGINT;
ALTER TABLE mart.attendance_agent_day ADD COLUMN status_source VARCHAR;

ALTER TABLE mart.correction_candidate ADD COLUMN observed_source VARCHAR;
ALTER TABLE mart.correction_candidate ADD COLUMN verint_reconciliation VARCHAR;
ALTER TABLE mart.correction_candidate ADD COLUMN verint_activity VARCHAR;
ALTER TABLE mart.correction_candidate ADD COLUMN verint_category VARCHAR;
ALTER TABLE mart.correction_candidate ADD COLUMN verint_overlap_minutes BIGINT;
ALTER TABLE mart.correction_candidate ADD COLUMN verint_source_file VARCHAR;

ALTER TABLE mart.absence_event ADD COLUMN reconciliation_status VARCHAR;
ALTER TABLE mart.absence_event ADD COLUMN verint_activity VARCHAR;
ALTER TABLE mart.absence_event ADD COLUMN verint_category VARCHAR;
ALTER TABLE mart.absence_event ADD COLUMN verint_overlap_minutes BIGINT;
ALTER TABLE mart.absence_event ADD COLUMN verint_source_file VARCHAR;

ALTER TABLE mart.absence_agent_day ADD COLUMN unverified_minutes BIGINT NOT NULL DEFAULT 0;
ALTER TABLE mart.absence_agent_day ADD COLUMN corrected_minutes BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.forecast_hour ADD COLUMN service_scope VARCHAR;
ALTER TABLE mart.forecast_hour ADD COLUMN comparison_scope VARCHAR;
ALTER TABLE mart.forecast_hour ADD COLUMN mapping_status VARCHAR;
ALTER TABLE mart.forecast_hour ADD COLUMN mapping_sha256 VARCHAR;

DROP INDEX IF EXISTS idx_forecast_hour_grain;
CREATE UNIQUE INDEX idx_forecast_hour_grain_v050
ON mart.forecast_hour(business_date, hour_start, service_scope, source_file, queue_name);

ALTER TABLE mart.intraday_queue_interval ADD COLUMN service_scope VARCHAR;
ALTER TABLE mart.intraday_queue_interval ADD COLUMN comparison_scope VARCHAR;
ALTER TABLE mart.intraday_queue_interval ADD COLUMN designation VARCHAR;
ALTER TABLE mart.intraday_queue_interval ADD COLUMN mapping_status VARCHAR;
ALTER TABLE mart.intraday_queue_interval ADD COLUMN mapping_sha256 VARCHAR;

ALTER TABLE mart.service_interval ADD COLUMN service_scope VARCHAR;
ALTER TABLE mart.service_interval ADD COLUMN comparison_scope VARCHAR;
ALTER TABLE mart.service_interval ADD COLUMN designation VARCHAR;
ALTER TABLE mart.service_interval ADD COLUMN mapping_status VARCHAR;
ALTER TABLE mart.service_interval ADD COLUMN mapping_sha256 VARCHAR;

CREATE TABLE IF NOT EXISTS mart.verint_final_exception (
    exception_key VARCHAR PRIMARY KEY,
    agent_day_key VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    activity VARCHAR,
    category VARCHAR,
    event_start TIMESTAMP,
    event_end TIMESTAMP,
    minutes BIGINT NOT NULL,
    exception_type VARCHAR NOT NULL,
    source_file VARCHAR,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE TABLE IF NOT EXISTS meta.mapping_application (
    run_id VARCHAR PRIMARY KEY,
    mapping_sha256 VARCHAR NOT NULL,
    mapping_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_verint_final_exception_date_agent
ON mart.verint_final_exception(business_date, agent_id);
