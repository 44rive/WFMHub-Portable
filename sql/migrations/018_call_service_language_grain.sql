-- The Call-by-Call model can legitimately contain the same mapped queue and
-- time bucket in more than one language.  Preserve that governed dimension in
-- the physical key instead of colliding at the former queue/time-only grain.

ALTER TABLE mart.call_service_hour
RENAME TO mart_call_service_hour_pre_language_grain;

CREATE TABLE mart.call_service_hour (
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

INSERT INTO mart.call_service_hour (
    business_date, hour_start, source_system, service_scope,
    comparison_scope, queue, designation, language, offered, answered,
    abandoned, short_abandoned, answered_within_target, talk_seconds,
    hold_seconds, wrap_seconds, handled_seconds, service_level,
    service_availability, abandon_rate, aht_seconds, call_legs,
    transferred_legs, source_files, mapping_sha256, rule_version,
    rule_sha256, abandoned_within_target
)
SELECT
    business_date, hour_start, source_system, service_scope,
    comparison_scope, queue, designation,
    coalesce(nullif(upper(trim(language)), ''), '(blank)'),
    offered, answered, abandoned, short_abandoned, answered_within_target,
    talk_seconds, hold_seconds, wrap_seconds, handled_seconds, service_level,
    service_availability, abandon_rate, aht_seconds, call_legs,
    transferred_legs, source_files, mapping_sha256, rule_version,
    rule_sha256, abandoned_within_target
FROM mart_call_service_hour_pre_language_grain;

DROP TABLE mart_call_service_hour_pre_language_grain;

CREATE INDEX idx_call_service_hour_scope
ON mart.call_service_hour(business_date, comparison_scope, service_scope);

ALTER TABLE mart.call_service_15min
RENAME TO mart_call_service_15min_pre_language_grain;

CREATE TABLE mart.call_service_15min (
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

INSERT INTO mart.call_service_15min (
    business_date, interval_start, interval_end, source_system,
    service_scope, comparison_scope, queue, designation, language,
    offered, answered, abandoned, short_abandoned,
    abandoned_within_target, answered_within_target, talk_seconds,
    hold_seconds, wrap_seconds, handled_seconds, service_level,
    service_availability, abandon_rate, aht_seconds, call_legs,
    transferred_legs, source_files, mapping_sha256, rule_version, rule_sha256
)
SELECT
    business_date, interval_start, interval_end, source_system,
    service_scope, comparison_scope, queue, designation,
    coalesce(nullif(upper(trim(language)), ''), '(blank)'),
    offered, answered, abandoned, short_abandoned,
    abandoned_within_target, answered_within_target, talk_seconds,
    hold_seconds, wrap_seconds, handled_seconds, service_level,
    service_availability, abandon_rate, aht_seconds, call_legs,
    transferred_legs, source_files, mapping_sha256, rule_version, rule_sha256
FROM mart_call_service_15min_pre_language_grain;

DROP TABLE mart_call_service_15min_pre_language_grain;

CREATE INDEX idx_call_service_15min_scope
ON mart.call_service_15min(
    business_date, interval_start, comparison_scope, service_scope
);
