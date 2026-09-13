-- Separate workforce capacity grain from Storm service queue grain.
-- Existing rows are retained for safety and rebuilt from unchanged sources on
-- the next model refresh.

ALTER TABLE mart.staffing_interval RENAME TO mart_staffing_interval_pre_capacity;

CREATE TABLE mart.staffing_interval (
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

INSERT INTO mart.staffing_interval (
    business_date, interval_start, interval_end, lob, language,
    planning_group, staff_type, capacity_mapping_status,
    scheduled_agents, observed_agents, productive_agents, auxiliary_agents,
    scheduled_fte, elapsed_scheduled_fte, observed_fte, productive_fte,
    staffing_variance_fte, staffing_gap_fte, staffing_state, evidence_basis,
    evaluation_as_of, gross_scheduled_fte, planned_time_off_fte
)
SELECT business_date, interval_start, interval_end, lob, language,
       lob, lob, 'LEGACY_REBUILD_REQUIRED',
       scheduled_agents, observed_agents, productive_agents, auxiliary_agents,
       scheduled_fte, elapsed_scheduled_fte, observed_fte, productive_fte,
       staffing_variance_fte, staffing_gap_fte, staffing_state, evidence_basis,
       evaluation_as_of, gross_scheduled_fte, planned_time_off_fte
FROM mart_staffing_interval_pre_capacity;

DROP TABLE mart_staffing_interval_pre_capacity;

CREATE INDEX idx_staffing_capacity_grain
ON mart.staffing_interval(
    business_date, interval_start, planning_group, staff_type
);
