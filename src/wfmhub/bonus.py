"""Governed Bonus Matrix import, calculation and standard report product."""

from __future__ import annotations

import hashlib
from calendar import monthrange
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import Config
from .database import DatabaseConnection
from .decision_products import _atomic_book, _audit_rows, _finish, _ratio
from .bonus_analysis_report import _cached_result_total
from .shared_reports import (
    _bonus_period,
    _bonus_value,
    _calculate_bonus,
    _clean,
    _number,
    _policy_map,
    _read_bonus_source,
)
from .template_reports import KpiCard


@dataclass(frozen=True)
class BonusImportResult:
    import_id: str
    period: str
    agents: int
    rules: int
    policies: int
    unchanged: bool


def _text(value: Any) -> str | None:
    cleaned = _clean(value)
    return cleaned or None


def import_bonus_matrix(conn: DatabaseConnection, source: Path) -> BonusImportResult:
    """Read Bonus Matrix v1.2 without editing it and activate one period version."""

    source = source.resolve()
    if not source.exists():
        raise FileNotFoundError(source)
    content = source.read_bytes()
    source_sha = hashlib.sha256(content).hexdigest()
    import_id = source_sha
    existing = conn.execute(
        "SELECT period, agent_rows, rule_rows, policy_rows, active, source_cached_total FROM raw.bonus_import WHERE import_id=?",
        [import_id],
    ).fetchone()
    if existing and bool(existing[4]):
        raw_count = conn.execute(
            "SELECT count(*) FROM raw.bonus_agent_month WHERE import_id=?", [import_id],
        ).fetchone()[0]
        rule_count = conn.execute(
            "SELECT count(*) FROM raw.bonus_kpi_rule WHERE import_id=?", [import_id],
        ).fetchone()[0]
        policy_count = conn.execute(
            "SELECT count(*) FROM raw.bonus_policy WHERE import_id=?", [import_id],
        ).fetchone()[0]
        mart_count = conn.execute(
            "SELECT count(*) FROM mart.bonus_agent_month WHERE import_id=?", [import_id],
        ).fetchone()[0]
        if (
            raw_count == existing[1]
            and mart_count == existing[1]
            and rule_count == existing[2]
            and policy_count == existing[3]
        ):
            return BonusImportResult(import_id, existing[0], existing[1], existing[2], existing[3], True)

    records, rules, policies = _read_bonus_source(source)
    period = _bonus_period(records)
    if not records:
        raise ValueError("Bonus source contains no populated Raw_Data agent rows")
    policy_values = _policy_map(policies)
    policies_ready = bool(policies) and all(
        len(row) > 5 and _clean(row[5]).casefold() == "validated" for row in policies
    )
    results = [_calculate_bonus(record, rules, policy_values, policies_ready) for record in records]
    source_cached_total, source_cached_rows = _cached_result_total(source)

    agent_sql = """INSERT INTO raw.bonus_agent_month VALUES
        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
    rule_map = {(rule.population, rule.kpi): rule for rule in rules}
    actual_names = (
        ("AHT", ("AHT",)),
        ("Productivity", ("Productivity",)),
        ("PCS Score", ("PCS Score",)),
        ("PCS % (Participation)", ("PCS % Participation", "PCS % (Participation)")),
        ("QM", ("QM",)),
        ("Abs%", ("Abs%",)),
        ("Extra Bonus (PCS Score)", ("PCS Score",)),
    )

    conn.execute("SAVEPOINT import_bonus_matrix")
    try:
        conn.execute("UPDATE raw.bonus_import SET active=false WHERE period=?", [period])
        conn.execute(
            """INSERT OR REPLACE INTO raw.bonus_import
               (import_id, source_path, file_name, source_sha256, imported_at, period,
                active, agent_rows, rule_rows, policy_rows, source_cached_total, source_cached_rows)
               VALUES (?, ?, ?, ?, ?, ?, true, ?, ?, ?, ?, ?)""",
            [import_id, str(source), source.name, source_sha, datetime.now(), period,
             len(records), len(rules), len(policies), source_cached_total, source_cached_rows],
        )
        conn.execute("DELETE FROM raw.bonus_agent_month WHERE import_id=?", [import_id])
        conn.execute("DELETE FROM raw.bonus_kpi_rule WHERE import_id=?", [import_id])
        conn.execute("DELETE FROM raw.bonus_policy WHERE import_id=?", [import_id])

        for row_number, record in enumerate(records, 1):
            conn.execute(agent_sql, [
                import_id, row_number, _text(_bonus_value(record, "Period")) or period,
                _text(_bonus_value(record, "Agent ID")), _text(_bonus_value(record, "Agent Name")),
                _text(_bonus_value(record, "Population")), _number(_bonus_value(record, "AHT")),
                _number(_bonus_value(record, "Productivity")), _number(_bonus_value(record, "PCS Score")),
                _number(_bonus_value(record, "PCS % Participation", "PCS % (Participation)")),
                _number(_bonus_value(record, "QM")), _number(_bonus_value(record, "Abs%")),
                int(_number(_bonus_value(record, "VOC Detractor Count")) or 0),
                _text(_bonus_value(record, "Currency")) or "MAD",
                _number(_bonus_value(record, "Monthly Fixed Salary")),
                _number(_bonus_value(record, "Target Bonus Rate")),
                _number(_bonus_value(record, "Reference Bonus Override")),
                _number(_bonus_value(record, "Eligible Days")),
                _number(_bonus_value(record, "Scheduled Days")),
                _text(_bonus_value(record, "Employment Status")),
                _text(_bonus_value(record, "Data Status")),
                _text(_bonus_value(record, "Notes")),
            ])
        for rule in rules:
            conn.execute(
                "INSERT INTO raw.bonus_kpi_rule VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                [import_id, rule.population, rule.kpi, rule.direction, rule.tier1_bonus,
                 rule.tier1_target, rule.tier2_bonus, rule.tier2_target],
            )
        for policy in policies:
            values = list(policy[:7]) + [None] * (7 - len(policy[:7]))
            conn.execute(
                "INSERT INTO raw.bonus_policy VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                [import_id, *[_text(value) for value in values[:7]]],
            )

        conn.execute("DELETE FROM mart.bonus_agent_month WHERE period=?", [period])
        conn.execute("DELETE FROM mart.bonus_kpi_result WHERE period=?", [period])
        for record, result in zip(records, results):
            agent_id = _text(_bonus_value(record, "Agent ID"))
            agent_name = _text(_bonus_value(record, "Agent Name"))
            population = _text(_bonus_value(record, "Population")) or ""
            earned = list(result.earned) + [None] * (7 - len(result.earned))
            conn.execute(
                """INSERT INTO mart.bonus_agent_month VALUES
                   (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                [period, agent_id, agent_name, population, result.core_ready, result.eligibility,
                 *earned[:7], result.gross, result.malus, result.final, result.reference,
                 result.proration, result.scenario, result.release, result.status,
                 result.issue, import_id, source_sha],
            )
            for index, (kpi, aliases) in enumerate(actual_names):
                rule = rule_map.get((population, kpi))
                if rule is None:
                    continue
                conn.execute(
                    "INSERT INTO mart.bonus_kpi_result VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [period, agent_id, agent_name, population, kpi,
                     _number(_bonus_value(record, *aliases)), earned[index], rule.direction,
                     rule.tier1_target, rule.tier2_target, import_id],
                )
        conn.execute("RELEASE SAVEPOINT import_bonus_matrix")
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT import_bonus_matrix")
        conn.execute("RELEASE SAVEPOINT import_bonus_matrix")
        raise
    return BonusImportResult(import_id, period, len(records), len(rules), len(policies), False)


def _selected_period(conn: DatabaseConnection, end: date) -> str:
    wanted = end.strftime("%Y-%m")
    exists = conn.execute("SELECT 1 FROM mart.bonus_agent_month WHERE period=? LIMIT 1", [wanted]).fetchone()
    if exists:
        return wanted
    latest = conn.execute("SELECT max(period) FROM mart.bonus_agent_month").fetchone()[0]
    return latest or wanted


def build_bonus_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    period = _selected_period(conn, end)
    period_start = date.fromisoformat(f"{period}-01")
    period_end = period_start.replace(day=monthrange(period_start.year, period_start.month)[1])
    book, partial, target = _atomic_book(
        config, "bonus", "BONUS PERFORMANCE", period_start, period_end, output,
    )
    totals = conn.execute(
        """SELECT count(*), coalesce(sum(scenario_payout),0), coalesce(sum(released_payout),0),
                  coalesce(sum(CASE WHEN release_status='READY' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN release_status<>'READY' THEN 1 ELSE 0 END),0),
                  avg(final_achievement), avg(proration)
           FROM mart.bonus_agent_month WHERE period=?""",
        [period],
    ).fetchone()
    agents, scenario, released, ready, blocked, achievement, proration = totals
    import_row = conn.execute(
        """SELECT import_id, file_name, source_sha256, imported_at,
                  source_cached_total, source_cached_rows
           FROM raw.bonus_import WHERE period=? AND active=true ORDER BY imported_at DESC LIMIT 1""",
        [period],
    ).fetchone()
    policy_both = False
    if import_row:
        value = conn.execute(
            "SELECT selected_value FROM raw.bonus_policy WHERE import_id=? AND lower(policy)='absence treatment'",
            [import_row[0]],
        ).fetchone()
        policy_both = bool(value and str(value[0]).casefold() == "both")
    status = "FINAL" if agents and not blocked and released else "INCOMPLETE"
    status_text = (
        "Absence is configured as both KPI and eligibility/proration; this double-treatment must be resolved"
        if policy_both else
        f"{blocked:,} agent(s) remain blocked by input, policy, data or eligibility gates"
        if blocked else
        "All imported policy and data gates are ready"
        if agents else
        "No Bonus Matrix has been imported; use Import Bonus Matrix v1.2"
    )
    source_total = import_row[4] if import_row else None
    controlled_difference = scenario - source_total if source_total is not None else None
    kpi_rows = conn.execute(
        """SELECT kpi, count(*), avg(actual_value), avg(earned_weight),
                  sum(CASE WHEN coalesce(earned_weight,0)>0 THEN 1 ELSE 0 END)
           FROM mart.bonus_kpi_result WHERE period=? GROUP BY kpi ORDER BY kpi""",
        [period],
    ).fetchall()
    book.dashboard(
        [
            KpiCard("Scenario payout", scenario, "money", "Management scenario"),
            KpiCard("Source workbook payout", source_total, "money", "Cached Results sheet"),
            KpiCard("Control adjustment", controlled_difference, "money", "Incomplete rows held"),
            KpiCard("Released payout", released, "money", "Payroll-gated amount"),
            KpiCard("Ready agents", ready, "integer", f"{agents:,} imported for {period}"),
            KpiCard("Blocked agents", blocked, "integer"),
            KpiCard("Average achievement", achievement, "percent"),
            KpiCard("Average proration", proration, "percent"),
        ],
        status,
        status_text,
        ["KPI", "Configured Agents", "Average Actual", "Average Earned Weight", "Agents Earning"],
        kpi_rows,
        [
            "Bonus Matrix v1.2 is imported read-only. Python/SQLite recalculates the result and keeps the source hash.",
            "Scenario payout is visible for discussion; Released payout stays blank until every required gate passes.",
            "One absence must have one financial consequence: KPI, eligibility or proration—never a duplicate penalty.",
            "Threshold changes must be effective-dated and simulated before they are approved for payroll.",
        ],
        (("Agents earning", 4),),
        "column",
    )
    detail_headers = [
        "period", "agent_id", "agent_name", "population", "core_ready", "eligibility",
        "aht_earned", "productivity_earned", "pcs_earned", "participation_earned",
        "qm_earned", "absence_earned", "extra_pcs_earned", "gross_achievement",
        "voc_malus", "final_achievement", "reference_bonus", "proration",
        "scenario_payout", "released_payout", "release_status", "data_issue",
    ]
    cursor = conn.execute(
        f"SELECT {', '.join(detail_headers)} FROM mart.bonus_agent_month WHERE period=? ORDER BY population, release_status, agent_name",
        [period],
    )
    detail_rows = cursor.fetchall()
    book.table("AGENT_DETAIL", "Bonus agent results", "Governed monthly calculation and release gate per agent.", detail_headers, detail_rows)
    kpi_headers = ["period", "agent_id", "agent_name", "population", "kpi", "actual_value", "earned_weight", "direction", "tier1_target", "tier2_target"]
    kpi_detail = conn.execute(
        f"SELECT {', '.join(kpi_headers)} FROM mart.bonus_kpi_result WHERE period=? ORDER BY kpi, population, agent_name",
        [period],
    ).fetchall()
    book.table("KPI_ANALYSIS", "Bonus KPI analysis", "One agent/KPI result for target distribution and scenario analysis.", kpi_headers, kpi_detail)
    actions = [row for row in detail_rows if row[detail_headers.index("release_status")] != "READY"]
    book.table("ACTIONS", "Bonus release exceptions", "Resolve every blocked row before payroll release.", detail_headers, actions)
    book.definitions([
        ("KPI award", "Tier 1 tested before Tier 2 using configured direction", "Achievement contribution", "Population-specific rule"),
        ("Final achievement", "Gross achievement after VOC malus and achievement cap", "Payout multiplier", "Policy controlled"),
        ("Scenario payout", "Reference bonus x proration x final achievement", "Management simulation", "Not payroll authority"),
        ("Released payout", "Scenario only when policy, eligibility and data gates pass", "Payroll handoff", "Blank while blocked"),
        ("Absence", "KPI, eligibility or proration according to approved policy", "Attendance consequence", "Never double-count the same absence"),
    ])
    extra = []
    if import_row:
        extra = [
            ("Source workbook", import_row[1], import_row[3]),
            ("Source SHA-256", import_row[2], "Proves source was not changed"),
            ("Source cached payout", import_row[4], f"{import_row[5]} cached Results row(s)"),
        ]
    book.audit(_audit_rows(conn, config, "bonus", period_start, period_end, extra))
    return _finish(book, partial, target)
