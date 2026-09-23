-- WFMHub current SQLite schema. Logical names such as mart.attendance_agent_day
-- are stored with portable prefixes such as mart_attendance_agent_day.

CREATE TABLE IF NOT EXISTS core_dim_agent (
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

CREATE TABLE IF NOT EXISTS mart_absence_agent_day (
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
    break_minutes BIGINT NOT NULL,
    lunch_minutes BIGINT NOT NULL,
    planned_net_minutes BIGINT NOT NULL,
    production_minutes BIGINT NOT NULL,
    absence_minutes BIGINT NOT NULL,
    vacation_minutes BIGINT NOT NULL,
    unpaid_minutes BIGINT NOT NULL,
    shrinkage_minutes BIGINT NOT NULL,
    late_minutes BIGINT NOT NULL,
    early_leave_minutes BIGINT NOT NULL,
    no_show_minutes BIGINT NOT NULL,
    unmapped_minutes BIGINT NOT NULL,
    absence_rate DOUBLE,
    vacation_rate DOUBLE,
    shrinkage_rate DOUBLE,
    absence_day BOOLEAN NOT NULL,
    absence_spell VARCHAR,
    absence_spells BIGINT NOT NULL,
    absence_days DOUBLE NOT NULL,
    bradford_factor DOUBLE NOT NULL,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
, unverified_minutes BIGINT NOT NULL DEFAULT 0, corrected_minutes BIGINT NOT NULL DEFAULT 0);

CREATE TABLE IF NOT EXISTS mart_absence_event (
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
    event_start TIMESTAMP,
    event_end TIMESTAMP,
    minutes BIGINT NOT NULL,
    hours DOUBLE NOT NULL,
    planned BOOLEAN NOT NULL,
    working BOOLEAN NOT NULL,
    counts_as_absence BOOLEAN NOT NULL,
    counts_as_vacation BOOLEAN NOT NULL,
    counts_as_unpaid BOOLEAN NOT NULL,
    counts_as_shrinkage BOOLEAN NOT NULL,
    mapped BOOLEAN NOT NULL,
    evidence_type VARCHAR NOT NULL,
    source_file VARCHAR,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
, reconciliation_status VARCHAR, verint_activity VARCHAR, verint_category VARCHAR, verint_overlap_minutes BIGINT, verint_source_file VARCHAR);

CREATE TABLE IF NOT EXISTS mart_agent_pcs_day (
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
    call_legs BIGINT,
    handled_calls BIGINT,
    inbound_calls BIGINT,
    outbound_calls BIGINT,
    talk_seconds BIGINT,
    hold_seconds BIGINT,
    wrap_seconds BIGINT,
    handle_seconds BIGINT,
    average_talk_seconds DOUBLE,
    average_hold_seconds DOUBLE,
    average_wrap_seconds DOUBLE,
    average_handle_seconds DOUBLE,
    pcs_enabled_calls BIGINT,
    survey_responses BIGINT,
    response_rate DOUBLE,
    q1_response_count BIGINT,
    q1_score_sum DOUBLE,
    q1_average DOUBLE,
    q2_response_count BIGINT,
    q2_score_sum DOUBLE,
    q2_average DOUBLE,
    pcs_score_count BIGINT,
    pcs_score_sum DOUBLE,
    pcs_average DOUBLE,
    top_box_responses BIGINT,
    low_score_responses BIGINT,
    top_box_percent DOUBLE,
    low_score_percent DOUBLE,
    comments_count BIGINT
, transferred_legs BIGINT NOT NULL DEFAULT 0, pcs_status_calls BIGINT NOT NULL DEFAULT 0, pcs_participation_responses BIGINT NOT NULL DEFAULT 0, pcs_participation_rate DOUBLE, pcs_invalid_responses BIGINT NOT NULL DEFAULT 0, pcs_status_blank_responses BIGINT NOT NULL DEFAULT 0, pcs_response_without_status BIGINT NOT NULL DEFAULT 0);

CREATE TABLE IF NOT EXISTS mart_analysis_finding (
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

CREATE TABLE IF NOT EXISTS mart_attendance_agent_day (
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
, actual_first_seen TIMESTAMP, actual_last_seen TIMESTAMP, actual_evidence VARCHAR, status_covered_minutes BIGINT, status_source VARCHAR, shift_state VARCHAR, call_action VARCHAR, requires_call BOOLEAN NOT NULL DEFAULT false, is_provisional BOOLEAN NOT NULL DEFAULT false, evaluation_as_of TIMESTAMP, planned_work_minutes BIGINT, planning_overlay VARCHAR, planning_overlay_minutes BIGINT NOT NULL DEFAULT 0, planning_overlay_source VARCHAR);

CREATE TABLE IF NOT EXISTS mart_bonus_agent_month (
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
    source_sha256 VARCHAR NOT NULL, team_leader VARCHAR, ops_manager VARCHAR,
    PRIMARY KEY (period, agent_id)
);

CREATE TABLE IF NOT EXISTS mart_bonus_kpi_result (
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

CREATE TABLE IF NOT EXISTS mart_call_service_15min (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    interval_end TIMESTAMP NOT NULL,
    source_system VARCHAR NOT NULL,
    service_scope VARCHAR NOT NULL,
    comparison_scope VARCHAR NOT NULL,
    queue VARCHAR NOT NULL,
    designation VARCHAR,
    language VARCHAR NOT NULL,
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
    PRIMARY KEY (
        business_date, interval_start, source_system, service_scope,
        comparison_scope, queue, language
    )
);

CREATE TABLE IF NOT EXISTS mart_call_service_hour (
    business_date DATE NOT NULL,
    hour_start TIMESTAMP NOT NULL,
    source_system VARCHAR NOT NULL,
    service_scope VARCHAR NOT NULL,
    comparison_scope VARCHAR NOT NULL,
    queue VARCHAR NOT NULL,
    designation VARCHAR,
    language VARCHAR NOT NULL,
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
    abandoned_within_target BIGINT NOT NULL DEFAULT 0,
    PRIMARY KEY (
        business_date, hour_start, source_system, service_scope,
        comparison_scope, queue, language
    )
);

CREATE TABLE IF NOT EXISTS mart_correction_candidate (
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
, observed_source VARCHAR, verint_reconciliation VARCHAR, verint_activity VARCHAR, verint_category VARCHAR, verint_overlap_minutes BIGINT, verint_source_file VARCHAR);

CREATE TABLE IF NOT EXISTS mart_correction_residual_segment (
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

CREATE TABLE IF NOT EXISTS mart_forecast_hour (
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
, service_scope VARCHAR, comparison_scope VARCHAR, mapping_status VARCHAR, mapping_sha256 VARCHAR, source_interval_minutes BIGINT, source_interval_count BIGINT);

CREATE TABLE IF NOT EXISTS mart_forecast_interval (
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

CREATE TABLE IF NOT EXISTS mart_metric_value (
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

CREATE TABLE IF NOT EXISTS mart_planned_time_off_segment (
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

CREATE TABLE IF NOT EXISTS mart_schedule_integrity_agent_day (
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

CREATE TABLE IF NOT EXISTS mart_service_interval (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    hour_start TIMESTAMP,
    source_system VARCHAR NOT NULL,
    queue VARCHAR,
    business_partner VARCHAR,
    lob VARCHAR,
    language VARCHAR,
    offered DOUBLE,
    answered DOUBLE,
    abandoned DOUBLE,
    short_abandoned DOUBLE,
    answered_within_target DOUBLE,
    handled_seconds DOUBLE,
    sl_gross DOUBLE,
    sl_adjusted DOUBLE,
    sl_profile VARCHAR NOT NULL,
    service_level DOUBLE,
    service_availability DOUBLE,
    abandon_rate DOUBLE,
    aht_seconds DOUBLE,
    source_file VARCHAR,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL
, service_scope VARCHAR, comparison_scope VARCHAR, designation VARCHAR, mapping_status VARCHAR, mapping_sha256 VARCHAR, sl_target DOUBLE, sl_state VARCHAR, abandoned_within_target DOUBLE NOT NULL DEFAULT 0);

CREATE TABLE IF NOT EXISTS mart_shift_timeline_segment (
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

CREATE TABLE IF NOT EXISTS mart_source_health (
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
, scoped_out_count BIGINT NOT NULL DEFAULT 0);

CREATE TABLE IF NOT EXISTS mart_staffing_interval (
    business_date DATE NOT NULL,
    interval_start TIMESTAMP NOT NULL,
    interval_end TIMESTAMP NOT NULL,
    lob VARCHAR NOT NULL,
    language VARCHAR NOT NULL,
    planning_group VARCHAR NOT NULL,
    staff_type VARCHAR NOT NULL,
    capacity_mapping_status VARCHAR NOT NULL,
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
    gross_scheduled_fte DOUBLE NOT NULL DEFAULT 0,
    planned_time_off_fte DOUBLE NOT NULL DEFAULT 0,
    PRIMARY KEY (
        business_date, interval_start, lob, language, planning_group, staff_type
    )
);

CREATE TABLE IF NOT EXISTS mart_verint_final_absence_agent_day (
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

CREATE TABLE IF NOT EXISTS mart_verint_final_absence_event (
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

CREATE TABLE IF NOT EXISTS mart_verint_final_exception (
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

CREATE TABLE IF NOT EXISTS meta_analytics_application (
    run_id VARCHAR PRIMARY KEY,
    analytics_version VARCHAR NOT NULL,
    analytics_sha256 VARCHAR NOT NULL,
    analytics_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS meta_mapping_application (
    run_id VARCHAR PRIMARY KEY,
    mapping_sha256 VARCHAR NOT NULL,
    mapping_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS meta_metric_application (
    run_id VARCHAR PRIMARY KEY,
    catalog_version VARCHAR NOT NULL,
    catalog_sha256 VARCHAR NOT NULL,
    catalog_file VARCHAR NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS meta_quality_issue (
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

CREATE TABLE IF NOT EXISTS meta_refresh_run (
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

CREATE TABLE IF NOT EXISTS meta_rule_application (
    run_id VARCHAR PRIMARY KEY,
    rule_version VARCHAR NOT NULL,
    rule_sha256 VARCHAR NOT NULL,
    rule_file VARCHAR NOT NULL,
    effective_from DATE NOT NULL,
    applied_at TIMESTAMP NOT NULL
);

CREATE TABLE IF NOT EXISTS meta_schema_migration (
    version VARCHAR PRIMARY KEY,
    applied_at TIMESTAMP NOT NULL DEFAULT current_timestamp
);

CREATE TABLE IF NOT EXISTS meta_source_file (
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
, scope_fingerprint VARCHAR, scoped_out_count BIGINT NOT NULL DEFAULT 0, source_variant VARCHAR);

CREATE TABLE IF NOT EXISTS raw_agent_status (
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

CREATE TABLE IF NOT EXISTS raw_bonus_agent_month (
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
    notes VARCHAR, team_leader VARCHAR, ops_manager VARCHAR,
    PRIMARY KEY (import_id, source_row)
);

CREATE TABLE IF NOT EXISTS raw_bonus_import (
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
, source_cached_total DOUBLE, source_cached_rows BIGINT, import_version VARCHAR);

CREATE TABLE IF NOT EXISTS raw_bonus_kpi_rule (
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

CREATE TABLE IF NOT EXISTS raw_bonus_policy (
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

CREATE TABLE IF NOT EXISTS raw_call_leg (
    source_file_id VARCHAR NOT NULL,
    source_row BIGINT NOT NULL,
    call_key VARCHAR NOT NULL,
    interaction_key VARCHAR NOT NULL,
    business_date DATE,
    call_start TIMESTAMP,
    call_end TIMESTAMP,
    communication_type VARCHAR,
    call_direction VARCHAR,
    originating_address VARCHAR,
    business_partner_id VARCHAR,
    lob VARCHAR,
    destination_address VARCHAR,
    service VARCHAR,
    call_reference_number VARCHAR,
    call_id VARCHAR,
    call_progress VARCHAR,
    queue_wait_seconds BIGINT,
    queue_id VARCHAR,
    queue VARCHAR,
    call_treatment_id VARCHAR,
    call_treatment VARCHAR,
    agent_group_id VARCHAR,
    agent_group VARCHAR,
    called_user_group VARCHAR,
    agent_id VARCHAR,
    agent_name VARCHAR,
    clearing_party VARCHAR,
    talk_seconds BIGINT,
    hold_seconds BIGINT,
    wrap_seconds BIGINT,
    completion_code VARCHAR,
    transferred BOOLEAN,
    conference_at TIMESTAMP,
    conference_duration_seconds BIGINT,
    shared_call_reference VARCHAR,
    ringing_seconds BIGINT,
    internal BOOLEAN,
    direct BOOLEAN,
    menu_progress VARCHAR,
    recording_mode VARCHAR,
    recording_consent VARCHAR,
    language_ivr VARCHAR,
    language VARCHAR,
    voicebot_id VARCHAR,
    voicebot_destination VARCHAR,
    voicebot_percentage_split DOUBLE,
    voicebot_name VARCHAR,
    routing_intent VARCHAR,
    agent_intent VARCHAR,
    licence_plate VARCHAR,
    ani VARCHAR,
    post_call_survey_mode VARCHAR,
    pcs_status VARCHAR,
    question_1 VARCHAR,
    question_2 VARCHAR,
    question_3 VARCHAR,
    question_4 VARCHAR,
    question_5 VARCHAR,
    question_6 VARCHAR,
    question_7 VARCHAR,
    question_8 VARCHAR,
    question_9 VARCHAR,
    question_10 VARCHAR,
    question_1_score DOUBLE,
    question_2_score DOUBLE,
    question_3_score DOUBLE,
    question_4_score DOUBLE,
    question_5_score DOUBLE,
    question_6_score DOUBLE,
    question_7_score DOUBLE,
    question_8_score DOUBLE,
    question_9_score DOUBLE,
    question_10_score DOUBLE,
    session_identifier VARCHAR,
    callback_id VARCHAR,
    callback_at TIMESTAMP,
    callback_status VARCHAR,
    callback_offered BOOLEAN
);

CREATE TABLE IF NOT EXISTS raw_forecast_interval (
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

CREATE TABLE IF NOT EXISTS raw_fte_agent (
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

CREATE TABLE IF NOT EXISTS raw_fte_time_off (
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

CREATE TABLE IF NOT EXISTS raw_lilo (
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

CREATE TABLE IF NOT EXISTS raw_schedule_event (
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

CREATE TABLE IF NOT EXISTS raw_schedule_shift (
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

CREATE INDEX IF NOT EXISTS idx_absence_day_date_agent
ON mart_absence_agent_day(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_absence_event_date_agent
ON mart_absence_event(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_agent_pcs_day_date_agent
ON mart_agent_pcs_day(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_analysis_finding_period_domain
ON mart_analysis_finding(period_end, domain, severity, finding_rank);

CREATE INDEX IF NOT EXISTS idx_attendance_date_result
ON mart_attendance_agent_day(business_date, attendance_result, agent_id);

CREATE INDEX IF NOT EXISTS idx_bonus_agent_period
ON raw_bonus_agent_month(period, agent_id, import_id);

CREATE INDEX IF NOT EXISTS idx_bonus_import_period
ON raw_bonus_import(period, active, imported_at);

CREATE INDEX IF NOT EXISTS idx_bonus_mart_period_population
ON mart_bonus_agent_month(period, population, release_status);

CREATE INDEX IF NOT EXISTS idx_call_leg_date_agent
ON raw_call_leg(business_date, agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_call_leg_key
ON raw_call_leg(call_key, source_file_id);

CREATE INDEX IF NOT EXISTS idx_call_service_15min_scope
ON mart_call_service_15min(
    business_date, interval_start, comparison_scope, service_scope
);

CREATE INDEX IF NOT EXISTS idx_call_service_hour_scope
ON mart_call_service_hour(business_date, comparison_scope, service_scope);

CREATE INDEX IF NOT EXISTS idx_correction_residual_date_agent
ON mart_correction_residual_segment(business_date, agent_id, residual_start);

CREATE INDEX IF NOT EXISTS idx_correction_status_date
ON mart_correction_candidate(validation_status, business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_forecast_date_queue
ON raw_forecast_interval(business_date, queue_name, interval_start, source_file_id);

CREATE UNIQUE INDEX IF NOT EXISTS idx_forecast_hour_grain
ON mart_forecast_hour(business_date, hour_start, service_scope, source_file, queue_name);

CREATE UNIQUE INDEX IF NOT EXISTS idx_forecast_interval_grain
ON mart_forecast_interval(
    business_date, interval_start, service_scope, source_file, queue_name
);

CREATE INDEX IF NOT EXISTS idx_fte_agent_id
ON raw_fte_agent(agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_fte_time_off_agent_dates
ON raw_fte_time_off(agent_id, start_date, end_date, source_file_id);

CREATE INDEX IF NOT EXISTS idx_lilo_date_agent
ON raw_lilo(extract_date, agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_metric_value_period_metric
ON mart_metric_value(business_date, metric_id, method_id);

CREATE INDEX IF NOT EXISTS idx_metric_value_scope
ON mart_metric_value(metric_id, source_system, lob, language, team_leader, business_date);

CREATE INDEX IF NOT EXISTS idx_planned_time_off_date_agent
ON mart_planned_time_off_segment(business_date, agent_id, segment_start);

CREATE INDEX IF NOT EXISTS idx_quality_severity_date
ON meta_quality_issue(severity, business_date, issue_type);

CREATE INDEX IF NOT EXISTS idx_schedule_date_agent
ON raw_schedule_shift(schedule_date, agent_id, source_file_id);

CREATE INDEX IF NOT EXISTS idx_schedule_event_time_agent
ON raw_schedule_event(event_start, event_end, source_file_id, agent_id);

CREATE INDEX IF NOT EXISTS idx_schedule_integrity_date_agent
ON mart_schedule_integrity_agent_day(business_date, agent_id, classification);

CREATE INDEX IF NOT EXISTS idx_service_interval_date_scope
ON mart_service_interval(business_date, source_system, lob, language);

CREATE INDEX IF NOT EXISTS idx_shift_timeline_date_agent
ON mart_shift_timeline_segment(business_date, agent_id, segment_start);

CREATE INDEX IF NOT EXISTS idx_source_file_active
ON meta_source_file(source_family, active, status, modified_at);

CREATE INDEX IF NOT EXISTS idx_source_file_dedupe
ON meta_source_file(source_family, source_path, sha256, scope_fingerprint, status);

CREATE INDEX IF NOT EXISTS idx_source_file_family_path
ON meta_source_file(source_family, source_path, active);

CREATE INDEX IF NOT EXISTS idx_source_file_family_variant
ON meta_source_file(source_family, source_variant, active, status, modified_at);

CREATE UNIQUE INDEX IF NOT EXISTS idx_source_file_one_active_path
ON meta_source_file(source_family, source_path) WHERE active=true;

CREATE UNIQUE INDEX IF NOT EXISTS idx_source_health_family
ON mart_source_health(source_family);

CREATE INDEX IF NOT EXISTS idx_staffing_capacity_grain
ON mart_staffing_interval(
    business_date, interval_start, planning_group, staff_type
);

CREATE INDEX IF NOT EXISTS idx_status_agent_time
ON raw_agent_status(agent_id, status_start, status_end, source_file_id);

CREATE INDEX IF NOT EXISTS idx_status_date
ON raw_agent_status(extract_date, source_file_id);

CREATE INDEX IF NOT EXISTS idx_status_serial_source
ON raw_agent_status(serial_number, source_file_id);

CREATE INDEX IF NOT EXISTS idx_verint_final_absence_day_date_agent
ON mart_verint_final_absence_agent_day(business_date, agent_id);

CREATE INDEX IF NOT EXISTS idx_verint_final_absence_event_date_agent
ON mart_verint_final_absence_event(business_date, agent_id, event_start);

CREATE INDEX IF NOT EXISTS idx_verint_final_exception_date_agent
ON mart_verint_final_exception(business_date, agent_id);

CREATE TRIGGER IF NOT EXISTS trg_absence_day_key_not_null
BEFORE INSERT ON mart_absence_agent_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart_absence_agent_day.agent_day_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_absence_event_key_not_null
BEFORE INSERT ON mart_absence_event WHEN NEW.event_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart_absence_event.event_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_agent_pcs_day_key_not_null
BEFORE INSERT ON mart_agent_pcs_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart_agent_pcs_day.agent_day_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_attendance_key_not_null
BEFORE INSERT ON mart_attendance_agent_day WHEN NEW.agent_day_key IS NULL
BEGIN SELECT RAISE(ABORT, 'mart_attendance_agent_day.agent_day_key cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_correction_candidate_key_not_null
BEFORE INSERT ON mart_correction_candidate WHEN NEW.correction_id IS NULL
BEGIN SELECT RAISE(ABORT, 'mart_correction_candidate.correction_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_dim_agent_key_not_null
BEFORE INSERT ON core_dim_agent WHEN NEW.agent_id IS NULL
BEGIN SELECT RAISE(ABORT, 'core_dim_agent.agent_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_quality_issue_key_not_null
BEFORE INSERT ON meta_quality_issue WHEN NEW.issue_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta_quality_issue.issue_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_refresh_run_key_not_null
BEFORE INSERT ON meta_refresh_run WHEN NEW.run_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta_refresh_run.run_id cannot be NULL'); END;

CREATE TRIGGER IF NOT EXISTS trg_source_file_key_not_null
BEFORE INSERT ON meta_source_file WHEN NEW.file_id IS NULL
BEGIN SELECT RAISE(ABORT, 'meta_source_file.file_id cannot be NULL'); END;

CREATE VIEW IF NOT EXISTS core_clean_call_leg AS
SELECT source_file_id, source_row, call_key, interaction_key, business_date,
       call_start, call_end, communication_type, call_direction,
       originating_address, business_partner_id, lob, destination_address,
       service, call_reference_number, call_id, call_progress,
       queue_wait_seconds, queue_id, queue, call_treatment_id, call_treatment,
       agent_group_id, agent_group, called_user_group, agent_id, agent_name,
       clearing_party, talk_seconds, hold_seconds, wrap_seconds,
       completion_code, transferred, conference_at,
       conference_duration_seconds, shared_call_reference, ringing_seconds,
       internal, direct, menu_progress, recording_mode, recording_consent,
       language_ivr, language, voicebot_id, voicebot_destination,
       voicebot_percentage_split, voicebot_name, routing_intent, agent_intent,
       licence_plate, ani, post_call_survey_mode, pcs_status,
       question_1, question_2, question_3, question_4, question_5,
       question_6, question_7, question_8, question_9, question_10,
       question_1_score, question_2_score, question_3_score,
       question_4_score, question_5_score, question_6_score,
       question_7_score, question_8_score, question_9_score,
       question_10_score, session_identifier, callback_id, callback_at,
       callback_status, callback_offered, source_file
FROM (
    SELECT r.*, f.file_name AS source_file,
           row_number() OVER (
               PARTITION BY r.call_key
               ORDER BY f.modified_at DESC, f.loaded_at DESC,
                        f.file_name DESC, r.source_row DESC
           ) AS row_rank
    FROM raw_call_leg r
    JOIN meta_source_file f ON f.file_id=r.source_file_id
    WHERE f.active=true AND f.status='SUCCESS'
) ranked
WHERE row_rank=1;
