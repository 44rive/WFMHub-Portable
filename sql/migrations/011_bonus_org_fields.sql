ALTER TABLE raw.bonus_agent_month ADD COLUMN team_leader VARCHAR;
ALTER TABLE raw.bonus_agent_month ADD COLUMN ops_manager VARCHAR;

ALTER TABLE mart.bonus_agent_month ADD COLUMN team_leader VARCHAR;
ALTER TABLE mart.bonus_agent_month ADD COLUMN ops_manager VARCHAR;
