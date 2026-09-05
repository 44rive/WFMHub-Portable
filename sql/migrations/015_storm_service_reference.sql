-- Keep Storm's two abandoned-call exclusions separate:
--   short_abandoned            elapsed response time below 5 seconds
--   abandoned_within_target    non-short abandon inside the 20-second target
-- This makes the governed denominator reproducible from Call-by-Call.
ALTER TABLE mart.call_service_hour
ADD COLUMN abandoned_within_target BIGINT NOT NULL DEFAULT 0;

ALTER TABLE mart.service_interval
ADD COLUMN abandoned_within_target DOUBLE NOT NULL DEFAULT 0;
