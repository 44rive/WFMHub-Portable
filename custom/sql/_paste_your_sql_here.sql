-- Copy this file, remove the leading underscore, then edit it.
-- Only one read-only SELECT or WITH query is allowed.
SELECT business_date,
       sum(handled_calls) AS handled_calls,
       sum(survey_responses) AS survey_responses
FROM mart.agent_pcs_day
WHERE business_date BETWEEN :start AND :end
GROUP BY business_date
ORDER BY business_date;
