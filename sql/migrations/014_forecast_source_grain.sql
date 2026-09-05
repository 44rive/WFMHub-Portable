CREATE TABLE IF NOT EXISTS mart.forecast_interval (
    business_date DATE,
    interval_start TIMESTAMP,
    interval_end TIMESTAMP,
    interval_minutes BIGINT,
    queue_name VARCHAR,
    volume_forecast DOUBLE,
    abandons_forecast DOUBLE,
    fte_forecast DOUBLE,
    fte_required DOUBLE,
    headcount_forecast DOUBLE,
    net_staffing_forecast DOUBLE,
    sl_forecast DOUBLE,
    sl_required DOUBLE,
    aht_forecast_seconds DOUBLE,
    source_file VARCHAR,
    service_scope VARCHAR,
    comparison_scope VARCHAR,
    mapping_status VARCHAR,
    mapping_sha256 VARCHAR
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_forecast_interval_grain_v0171
ON mart.forecast_interval(
    business_date, interval_start, service_scope, source_file, queue_name
);

ALTER TABLE mart.forecast_hour ADD COLUMN source_interval_minutes BIGINT;
ALTER TABLE mart.forecast_hour ADD COLUMN source_interval_count BIGINT;
