CREATE TABLE IF NOT EXISTS meta.metric_application (
    run_id VARCHAR PRIMARY KEY,
    catalog_version VARCHAR NOT NULL,
    catalog_sha256 VARCHAR NOT NULL,
    catalog_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS meta.analytics_application (
    run_id VARCHAR PRIMARY KEY,
    analytics_version VARCHAR NOT NULL,
    analytics_sha256 VARCHAR NOT NULL,
    analytics_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS mart.metric_value (
    metric_key VARCHAR PRIMARY KEY,
    run_id VARCHAR NOT NULL,
    business_date DATE NOT NULL,
    interval_start TIMESTAMP,
    source_model VARCHAR NOT NULL,
    grain VARCHAR NOT NULL,
    entity_key VARCHAR NOT NULL,
    metric_id VARCHAR NOT NULL,
    method_id VARCHAR NOT NULL,
    method_effective_from DATE NOT NULL,
    domain VARCHAR NOT NULL,
    unit VARCHAR NOT NULL,
    aggregation VARCHAR NOT NULL,
    source_system VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    team_leader VARCHAR,
    agent_id VARCHAR,
    numerator DOUBLE,
    denominator DOUBLE,
    sample_size DOUBLE,
    metric_value DOUBLE,
    target_value DOUBLE,
    metric_state VARCHAR NOT NULL,
    catalog_version VARCHAR NOT NULL,
    catalog_sha256 VARCHAR NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_metric_value_period_metric
ON mart.metric_value(business_date, metric_id, method_id);

CREATE INDEX IF NOT EXISTS idx_metric_value_scope
ON mart.metric_value(metric_id, source_system, lob, language, team_leader, business_date);

CREATE TABLE IF NOT EXISTS mart.analysis_finding (
    finding_id VARCHAR PRIMARY KEY,
    run_id VARCHAR NOT NULL,
    finding_rank BIGINT NOT NULL,
    finding_type VARCHAR NOT NULL,
    severity VARCHAR NOT NULL,
    domain VARCHAR NOT NULL,
    metric_id VARCHAR,
    method_id VARCHAR,
    period_start DATE NOT NULL,
    period_end DATE NOT NULL,
    source_system VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    team_leader VARCHAR,
    agent_id VARCHAR,
    title VARCHAR NOT NULL,
    summary VARCHAR NOT NULL,
    current_value DOUBLE,
    reference_value DOUBLE,
    target_value DOUBLE,
    delta_value DOUBLE,
    unit VARCHAR,
    evidence_dataset VARCHAR NOT NULL,
    evidence_filter VARCHAR NOT NULL,
    catalog_version VARCHAR NOT NULL,
    catalog_sha256 VARCHAR NOT NULL,
    analytics_version VARCHAR NOT NULL,
    analytics_sha256 VARCHAR NOT NULL,
    created_at TIMESTAMP NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_analysis_finding_period_domain
ON mart.analysis_finding(period_end, domain, severity, finding_rank);
