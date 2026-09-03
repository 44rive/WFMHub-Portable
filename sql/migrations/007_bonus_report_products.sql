CREATE TABLE IF NOT EXISTS raw.bonus_import (
    import_id VARCHAR PRIMARY KEY,
    source_path VARCHAR NOT NULL,
    file_name VARCHAR NOT NULL,
    source_sha256 VARCHAR NOT NULL,
    imported_at TIMESTAMP NOT NULL,
    period VARCHAR NOT NULL,
    active BOOLEAN NOT NULL,
    agent_rows BIGINT NOT NULL,
    rule_rows BIGINT NOT NULL,
    policy_rows BIGINT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_bonus_import_period
ON raw.bonus_import(period, active, imported_at);

CREATE TABLE IF NOT EXISTS raw.bonus_agent_month (
    import_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    period VARCHAR NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    population VARCHAR,
    aht DOUBLE,
    productivity DOUBLE,
    pcs_score DOUBLE,
    pcs_participation DOUBLE,
    qm DOUBLE,
    absence_rate DOUBLE,
    voc_detractors BIGINT,
    currency VARCHAR,
    monthly_fixed_salary DOUBLE,
    target_bonus_rate DOUBLE,
    reference_bonus_override DOUBLE,
    eligible_days DOUBLE,
    scheduled_days DOUBLE,
    employment_status VARCHAR,
    data_status VARCHAR,
    notes VARCHAR,
    PRIMARY KEY (import_id, source_row)
);

CREATE INDEX IF NOT EXISTS idx_bonus_agent_period
ON raw.bonus_agent_month(period, agent_id, import_id);

CREATE TABLE IF NOT EXISTS raw.bonus_kpi_rule (
    import_id VARCHAR NOT NULL,
    population VARCHAR NOT NULL,
    kpi VARCHAR NOT NULL,
    direction VARCHAR NOT NULL,
    tier1_bonus DOUBLE NOT NULL,
    tier1_target DOUBLE NOT NULL,
    tier2_bonus DOUBLE NOT NULL,
    tier2_target DOUBLE NOT NULL,
    PRIMARY KEY (import_id, population, kpi)
);

CREATE TABLE IF NOT EXISTS raw.bonus_policy (
    import_id VARCHAR NOT NULL,
    policy VARCHAR NOT NULL,
    selected_value VARCHAR,
    allowed_values VARCHAR,
    formula_impact VARCHAR,
    owner VARCHAR,
    status VARCHAR,
    comments VARCHAR,
    PRIMARY KEY (import_id, policy)
);

CREATE TABLE IF NOT EXISTS mart.bonus_agent_month (
    period VARCHAR NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    population VARCHAR,
    core_ready BOOLEAN NOT NULL,
    eligibility VARCHAR,
    aht_earned DOUBLE,
    productivity_earned DOUBLE,
    pcs_earned DOUBLE,
    participation_earned DOUBLE,
    qm_earned DOUBLE,
    absence_earned DOUBLE,
    extra_pcs_earned DOUBLE,
    gross_achievement DOUBLE,
    voc_malus DOUBLE,
    final_achievement DOUBLE,
    reference_bonus DOUBLE,
    proration DOUBLE,
    scenario_payout DOUBLE,
    released_payout DOUBLE,
    release_status VARCHAR NOT NULL,
    data_issue VARCHAR,
    import_id VARCHAR NOT NULL,
    source_sha256 VARCHAR NOT NULL,
    PRIMARY KEY (period, agent_id)
);

CREATE INDEX IF NOT EXISTS idx_bonus_mart_period_population
ON mart.bonus_agent_month(period, population, release_status);

CREATE TABLE IF NOT EXISTS mart.bonus_kpi_result (
    period VARCHAR NOT NULL,
    agent_id VARCHAR NOT NULL,
    agent_name VARCHAR,
    population VARCHAR,
    kpi VARCHAR NOT NULL,
    actual_value DOUBLE,
    earned_weight DOUBLE,
    direction VARCHAR,
    tier1_target DOUBLE,
    tier2_target DOUBLE,
    import_id VARCHAR NOT NULL,
    PRIMARY KEY (period, agent_id, kpi)
);
