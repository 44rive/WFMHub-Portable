ALTER TABLE meta.source_file ADD COLUMN scope_fingerprint VARCHAR;
ALTER TABLE meta.source_file ADD COLUMN scoped_out_count BIGINT NOT NULL DEFAULT 0;
ALTER TABLE mart.source_health ADD COLUMN scoped_out_count BIGINT NOT NULL DEFAULT 0;

CREATE TRIGGER IF NOT EXISTS trg_refresh_run_key_not_null
BEFORE INSERT ON meta.refresh_run WHEN NEW.run_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta.refresh_run.run_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_source_file_key_not_null
BEFORE INSERT ON meta.source_file WHEN NEW.file_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta.source_file.file_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_quality_issue_key_not_null
BEFORE INSERT ON meta.quality_issue WHEN NEW.issue_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta.quality_issue.issue_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_correction_action_key_not_null
BEFORE INSERT ON core.correction_action WHEN NEW.correction_id IS NULL
BEGIN SELECT RAISE(ABORT, 'core.correction_action.correction_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_dim_agent_key_not_null
BEFORE INSERT ON core.dim_agent WHEN NEW.agent_id IS NULL
BEGIN SELECT RAISE(ABORT, 'core.dim_agent.agent_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_attendance_key_not_null
BEFORE INSERT ON mart.attendance_agent_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart.attendance_agent_day.agent_day_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_conformance_key_not_null
BEFORE INSERT ON mart.conformance_agent_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart.conformance_agent_day.agent_day_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_correction_candidate_key_not_null
BEFORE INSERT ON mart.correction_candidate WHEN NEW.correction_id IS NULL
BEGIN SELECT RAISE(ABORT, 'mart.correction_candidate.correction_id cannot be NULL'); END;

CREATE INDEX IF NOT EXISTS idx_source_file_dedupe
ON meta.source_file(source_family, source_path, sha256, scope_fingerprint, status);

CREATE UNIQUE INDEX IF NOT EXISTS idx_source_file_one_active_path
ON meta.source_file(source_family, source_path) WHERE active=true;

CREATE INDEX IF NOT EXISTS idx_source_file_active
ON meta.source_file(source_family, active, status, modified_at);

CREATE INDEX IF NOT EXISTS idx_fte_agent_id
ON raw.fte_agent(agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_schedule_date_agent
ON raw.schedule_shift(schedule_date, agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_schedule_event_time_agent
ON raw.schedule_event(event_start, event_end, source_file_id, agent_id);

CREATE INDEX IF NOT EXISTS idx_lilo_date_agent
ON raw.lilo(extract_date, agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_status_agent_time
ON raw.agent_status(agent_id, status_start, status_end, source_file_id);

CREATE INDEX IF NOT EXISTS idx_status_date
ON raw.agent_status(extract_date, source_file_id);

CREATE INDEX IF NOT EXISTS idx_status_serial_source
ON raw.agent_status(serial_number, source_file_id);

CREATE INDEX IF NOT EXISTS idx_forecast_date_queue
ON raw.forecast_interval(business_date, queue_name, interval_start, source_file_id);

CREATE INDEX IF NOT EXISTS idx_queue_actual_date_source
ON raw.queue_actual(business_date, source_system, interval_start, source_file_id);

CREATE INDEX IF NOT EXISTS idx_attendance_date_result
ON mart.attendance_agent_day(business_date, attendance_result, agent_id);

CREATE INDEX IF NOT EXISTS idx_conformance_date_agent
ON mart.conformance_agent_day(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_correction_status_date
ON mart.correction_candidate(validation_status, business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_quality_severity_date
ON meta.quality_issue(severity, business_date, issue_type);

CREATE UNIQUE INDEX IF NOT EXISTS idx_forecast_hour_grain
ON mart.forecast_hour(business_date, hour_start, queue_name);

CREATE UNIQUE INDEX IF NOT EXISTS idx_intraday_interval_grain
ON mart.intraday_queue_interval(
    source_system, business_date, interval_start,
    ifnull(queue, ''), ifnull(business_partner, ''), ifnull(lob, ''), ifnull(language, '')
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_source_health_family
ON mart.source_health(source_family);
