CREATE TABLE IF NOT EXISTS mart.call_service_hour (
    business_date DATE NOT NULL,
    hour_start TIMESTAMP NOT NULL,
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
    PRIMARY KEY (business_date, hour_start, comparison_scope, queue)
);

CREATE INDEX IF NOT EXISTS idx_call_service_hour_scope
ON mart.call_service_hour(business_date, comparison_scope, service_scope);
