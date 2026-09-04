CREATE TABLE IF NOT EXISTS raw.fte_time_off (
    source_file_id VARCHAR NOT NULL,
    source_sheet VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    source_kind VARCHAR NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    start_date DATE NOT NULL,
    end_date DATE,
    day_coverage VARCHAR,
    start_time TIME,
    end_time TIME,
    absence_type VARCHAR NOT NULL,
    record_status VARCHAR NOT NULL,
    comment VARCHAR
);

CREATE INDEX IF NOT EXISTS idx_fte_time_off_agent_dates
ON raw.fte_time_off(agent_id, start_date, end_date, source_file_id);

CREATE TABLE IF NOT EXISTS mart.planned_time_off_segment (
    segment_key VARCHAR PRIMARY KEY,
    agent_day_key VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    team_leader VARCHAR,
    ops_manager VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    source_kind VARCHAR NOT NULL,
    absence_type VARCHAR NOT NULL,
    record_status VARCHAR NOT NULL,
    segment_start TIMESTAMP NOT NULL,
    segment_end TIMESTAMP NOT NULL,
    planned_minutes BIGINT NOT NULL,
    source_file VARCHAR NOT NULL,
    source_sheet VARCHAR NOT NULL,
    source_row BIGINT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_planned_time_off_date_agent
ON mart.planned_time_off_segment(business_date, agent_id, segment_start);

ALTER TABLE mart.attendance_agent_day
ADD COLUMN planned_work_minutes BIGINT;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN planning_overlay VARCHAR;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN planning_overlay_minutes BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.attendance_agent_day
ADD COLUMN planning_overlay_source VARCHAR;

ALTER TABLE mart.staffing_interval
ADD COLUMN gross_scheduled_fte DOUBLE NOT NULL DEFAULT 0;

ALTER TABLE mart.staffing_interval
ADD COLUMN planned_time_off_fte DOUBLE NOT NULL DEFAULT 0;
