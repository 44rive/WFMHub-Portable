"""Focused WFM/Operations report products using one workbook design contract."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import report_pack, report_pack_folder
from .reports import COLORS, _query
from .rules import load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook, KpiCard, ModelTable


@dataclass(frozen=True)
class NamedPeriod:
    label: str
    start: date
    end: date


def _month_start(value: date) -> date:
    return value.replace(day=1)


def _previous_month(value: date) -> tuple[date, date]:
    end = value.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def _comparison_periods(start: date, end: date) -> list[NamedPeriod]:
    previous_start, previous_end = _previous_month(end)
    prior_mtd_end = min(previous_end, previous_start + timedelta(days=end.day - 1))
    duration = (end - start).days + 1
    prior_equal_end = start - timedelta(days=1)
    prior_equal_start = prior_equal_end - timedelta(days=duration - 1)
    periods = [
        NamedPeriod("Latest day", end, end),
        NamedPeriod("Selected period", start, end),
        NamedPeriod("Current MTD", _month_start(end), end),
        NamedPeriod("Prior-month same days", previous_start, prior_mtd_end),
        NamedPeriod("Previous full month", previous_start, previous_end),
    ]
    if start != _month_start(end):
        periods.insert(2, NamedPeriod("Previous equal period", prior_equal_start, prior_equal_end))
    return periods


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    return float(numerator) / float(denominator) if denominator else None


def _delta(value: float | None, reference: float | None, percent: bool = False) -> str:
    if value is None or reference is None:
        return "Comparison not available"
    difference = value - reference
    return f"{difference:+.1%} vs prior" if percent else f"{difference:+.2f} vs prior"


def _output_path(config: Config, key: str, start: date, end: date, generated: datetime, output: Path | None) -> Path:
    pack = report_pack(key)
    return (
        output
        or report_pack_folder(config, key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()


def _source_state(conn: DatabaseConnection, families: Sequence[str], through: date, final: bool = False) -> tuple[str, str]:
    if not families:
        return ("FINAL" if final else "LIVE"), "All required governed datasets are available"
    placeholders = ",".join("?" for _ in families)
    rows = conn.execute(
        f"""SELECT source_family, newest_business_date, status
            FROM mart.source_health WHERE source_family IN ({placeholders})""",
        list(families),
    ).fetchall()
    found = {str(row[0]).lower(): row for row in rows}
    problems = []
    for family in families:
        if family.casefold() in {"start_end", "activities"}:
            variant = "START_END" if family.casefold() == "start_end" else "ACTIVITIES"
            variant_row = conn.execute(
                """SELECT max(r.schedule_date), count(*)
                   FROM raw.schedule_shift r
                   JOIN meta.source_file f ON f.file_id=r.source_file_id
                   WHERE f.active=true AND f.status='SUCCESS' AND f.source_variant=?""",
                [variant],
            ).fetchone()
            variant_date, variant_count = variant_row
            if not variant_count:
                problems.append(f"{family}: no successful {variant} extract")
            elif variant_date is None or str(variant_date)[:10] < through.isoformat():
                problems.append(f"{family}: latest {variant_date or 'unknown'}")
            continue
        row = found.get(family.lower())
        if row is None:
            problems.append(f"{family}: no health record")
        elif str(row[2]).upper() != "SUCCESS":
            problems.append(f"{family}: {row[2]}")
        # FTE is a point-in-time scope roster rather than a dated fact source.
        # Its freshness is represented by load status and source hash, so a
        # NULL business date is expected and must not make every report red.
        elif family.casefold() != "fte" and (row[1] is None or str(row[1])[:10] < through.isoformat()):
            problems.append(f"{family}: latest {row[1] or 'unknown'}")
    if problems:
        return "INCOMPLETE", "; ".join(problems)
    if final:
        return "FINAL", f"Required sources loaded through {through}"
    if through >= date.today():
        return "PROVISIONAL", "Current-day values can still change before shifts and queues close"
    return "LIVE", f"Required sources loaded through {through}"


def _audit_rows(
    conn: DatabaseConnection,
    config: Config,
    report_key: str,
    start: date,
    end: date,
    extra: Iterable[Sequence[Any]] = (),
) -> list[Sequence[Any]]:
    latest = conn.execute(
        "SELECT run_id, finished_at, details FROM meta.refresh_run WHERE status='SUCCESS' ORDER BY finished_at DESC LIMIT 1"
    ).fetchone()
    return [
        ("Report product", report_key, "One workbook = one operational decision"),
        ("Selected period", f"{start} to {end}", "Explicit report boundary"),
        ("Generated", datetime.now(), "Local work-machine time"),
        ("Refresh run", latest[0] if latest else None, latest[2] if latest else "No successful refresh metadata"),
        ("Calculation authority", "Python + SQLite", "Excel contains presentation only"),
        ("Template model folder", str(config.output / "model_data" / report_key), "Power Query: connection only + Add to Data Model"),
        *list(extra),
    ]


def _atomic_book(
    config: Config,
    key: str,
    title: str,
    start: date,
    end: date,
    output: Path | None,
) -> tuple[DecisionWorkbook, Path, Path]:
    generated = datetime.now()
    target = _output_path(config, key, start, end, generated, output)
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    return DecisionWorkbook(partial, config, key, title, start, end, generated), partial, target


def _finish(book: DecisionWorkbook, partial: Path, target: Path) -> Path:
    try:
        book.close()
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def _pcs_aggregate(conn: DatabaseConnection, period: NamedPeriod) -> tuple[Any, ...]:
    row = conn.execute(
        """SELECT coalesce(sum(pcs_score_sum),0), coalesce(sum(survey_responses),0),
                  coalesce(sum(pcs_participation_responses),0), coalesce(sum(pcs_status_calls),0),
                  coalesce(sum(low_score_responses),0), coalesce(sum(top_box_responses),0),
                  coalesce(sum(inbound_calls),0)
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?""",
        [period.start, period.end],
    ).fetchone()
    score_sum, valid, participants, eligible, low, high, inbound = row
    return (
        period.label, period.start, period.end, _ratio(score_sum, valid),
        _ratio(participants, eligible), valid, eligible, low, high, inbound,
    )


def build_pcs_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the daily/MTD/monthly PCS decision product."""

    book, partial, target = _atomic_book(config, "pcs", "PCS PERFORMANCE", start, end, output)
    periods = _comparison_periods(start, end)
    comparison = [_pcs_aggregate(conn, period) for period in periods]
    by_label = {row[0]: row for row in comparison}
    latest = by_label["Latest day"]
    mtd = by_label["Current MTD"]
    prior = by_label["Prior-month same days"]
    status, status_text = _source_state(conn, ("fte", "calls"), end)
    book.dashboard(
        [
            KpiCard("Latest-day PCS", latest[3], "decimal", str(end)),
            KpiCard("Latest-day participation", latest[4], "percent", f"{latest[5]:,} valid response(s)"),
            KpiCard("MTD PCS", mtd[3], "decimal", _delta(mtd[3], prior[3])),
            KpiCard("MTD participation", mtd[4], "percent", _delta(mtd[4], prior[4], True)),
            KpiCard("MTD valid responses", mtd[5], "integer", f"{mtd[9]:,} inbound call leg(s)"),
            KpiCard("MTD PCSStatus=1", mtd[6], "integer", "Participation denominator"),
            KpiCard("MTD score <= 3", mtd[7], "integer", "Follow-up population"),
            KpiCard("MTD score > 3", mtd[8], "integer", "Positive population"),
        ],
        status,
        status_text,
        ["Period", "Start", "End", "PCS Average", "Participation %", "Valid Responses", "PCSStatus=1", "<=3", ">3", "Inbound Legs"],
        comparison,
        [
            "The report can be generated every three hours, but the KPI scope remains daily and monthly.",
            "Current MTD is compared with the same number of calendar days in the previous month; a partial month is never compared only with a full month.",
            "PCS average and participation are ratios of summed counters, not averages of agent percentages.",
            "Use the Excel template model files for PivotTables and LOB / Team Leader / Agent slicers.",
        ],
        (("PCS Average", 3), ("Participation", 4)),
    )

    headers, rows = _query(
        conn,
        """SELECT business_date, sum(pcs_score_sum) AS pcs_score_sum,
                  sum(survey_responses) AS valid_responses,
                  sum(pcs_participation_responses) AS participating_responses,
                  sum(pcs_status_calls) AS eligible_calls,
                  CASE WHEN sum(survey_responses)>0 THEN sum(pcs_score_sum)*1.0/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0 THEN sum(pcs_participation_responses)*1.0/sum(pcs_status_calls) END AS participation_rate,
                  sum(low_score_responses) AS score_le_3,
                  sum(top_box_responses) AS score_gt_3
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date ORDER BY business_date""",
        [min(period.start for period in periods), max(period.end for period in periods)],
    )
    book.table("TREND", "PCS daily trend", "Daily governed counters for current and comparison periods.", headers, rows)

    headers, rows = _query(
        conn,
        """SELECT agent_id, max(agent_name) AS agent_name, max(team_leader) AS team_leader,
                  max(ops_manager) AS ops_manager, max(lob) AS lob, max(language) AS language,
                  sum(inbound_calls) AS inbound_call_legs,
                  sum(pcs_status_calls) AS pcs_status_1,
                  sum(pcs_participation_responses) AS q1_nonblank,
                  sum(survey_responses) AS valid_q1,
                  sum(pcs_score_sum) AS q1_score_sum,
                  CASE WHEN sum(survey_responses)>0 THEN sum(pcs_score_sum)*1.0/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0 THEN sum(pcs_participation_responses)*1.0/sum(pcs_status_calls) END AS participation_rate,
                  sum(low_score_responses) AS score_le_3,
                  sum(top_box_responses) AS score_gt_3,
                  sum(pcs_invalid_responses) AS invalid_q1,
                  CASE WHEN sum(survey_responses)<20 THEN 'LOW_SAMPLE' ELSE 'OK' END AS sample_state
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY agent_id ORDER BY max(lob), max(team_leader), max(agent_name)""",
        [start, end],
    )
    ws = book.table("AGENT_DETAIL", "PCS agent detail", "Selected-period agent results; filter by LOB, Team Leader or Agent.", headers, rows)
    if rows:
        state_col = headers.index("sample_state")
        ws.conditional_format(4, state_col, 3 + len(rows), state_col, {"type": "text", "criteria": "containing", "value": "LOW_SAMPLE", "format": book.report.error})
    actions = [row for row in rows if row[headers.index("score_le_3")] or row[headers.index("sample_state")] == "LOW_SAMPLE"]
    book.table("ACTIONS", "PCS follow-up queue", "Agents with a low score or insufficient response sample. This is coaching evidence, not a penalty list.", headers, actions)
    book.definitions([
        ("PCS Average", "Sum of valid inbound Q1 scores / valid inbound Q1 responses", "Customer experience result", "Only configured discrete Q1 scores are valid"),
        ("PCS Participation", "Inbound raw Q1 nonblank / inbound PCSStatus=1", "Survey participation opportunity", "Invalid nonblank Q1 remains in the numerator"),
        ("Score <= 3", "Count of valid Q1 responses at or below 3", "Follow-up volume", "A count, not a percentage"),
        ("Low sample", "Fewer than 20 valid responses in the selected period", "Interpretation warning", "Does not change the calculated score"),
    ])
    book.audit(_audit_rows(conn, config, "pcs", start, end))
    return _finish(book, partial, target)


def _service_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    scope_marks = ",".join("?" for _ in profile.service_scopes)
    system_marks = ",".join("?" for _ in profile.source_systems)
    cursor = conn.execute(
        f"""SELECT business_date, interval_start, hour_start, source_system, queue,
                   service_scope, comparison_scope, designation, mapping_status,
                   offered, answered, abandoned, short_abandoned,
                   answered_within_target, handled_seconds, source_file
            FROM mart.service_interval
            WHERE business_date BETWEEN ? AND ?
              AND service_scope IN ({scope_marks})
              AND source_system IN ({system_marks})
            ORDER BY business_date, interval_start, source_system, queue""",
        [start, end, *profile.service_scopes, *profile.source_systems],
    )
    headers = [item[0] for item in cursor.description]
    return [dict(zip(headers, row)) for row in cursor.fetchall()]


def _profile_metric(catalog: MetricCatalog, profile: ServiceProfile, metric_id: str, on_date: date):
    methods = {}
    for service_scope in profile.service_scopes:
        for source_system in profile.source_systems:
            method = catalog.method_for(metric_id, on_date, {
                "lob": service_scope,
                "source_system": source_system,
            })
            if method is None:
                raise ValueError(
                    f"Service profile {profile.profile_id!r} selects metric {metric_id!r}, "
                    f"but no method applies to {service_scope}/{source_system} on {on_date}"
                )
            methods[(method.method_id, method.effective_from, method.priority)] = method
    if len(methods) != 1:
        names = ", ".join(key[0] for key in methods)
        raise ValueError(
            f"Service profile {profile.profile_id!r} crosses incompatible {metric_id} methods: {names}. "
            "Split it into separate service profiles."
        )
    return next(iter(methods.values()))


def _service_aggregate(
    rows: Iterable[dict[str, Any]],
    profile: ServiceProfile,
    catalog: MetricCatalog,
    on_date: date,
) -> dict[str, float | str | None]:
    values = list(rows)
    offered = sum(float(row.get("offered") or 0) for row in values)
    answered = sum(float(row.get("answered") or 0) for row in values)
    abandoned = sum(float(row.get("abandoned") or 0) for row in values)
    short = sum(float(row.get("short_abandoned") or 0) for row in values)
    within = sum(float(row.get("answered_within_target") or 0) for row in values)
    handled = sum(float(row.get("handled_seconds") or 0) for row in values)
    components = {
        "offered": offered,
        "answered": answered,
        "abandoned": abandoned,
        "short_abandoned": short,
        "answered_within_target": within,
        "handled_seconds": handled,
    }
    service_level = evaluate_metric(
        _profile_metric(catalog, profile, profile.service_level_metric, on_date), components,
    )
    availability = evaluate_metric(
        _profile_metric(catalog, profile, profile.availability_metric, on_date), components,
    )
    aht = evaluate_metric(
        _profile_metric(catalog, profile, profile.aht_metric, on_date), components,
    )
    return {
        "offered": offered,
        "answered": answered,
        "short_abandoned": short,
        "within_target": within,
        "service_level": service_level.value,
        "service_target": service_level.method.target,
        "service_state": service_level.state,
        "service_method": service_level.method.method_id,
        "availability": availability.value,
        "aht_seconds": aht.value,
    }


def build_service_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Build LOB service performance, beginning with the Ford OEM profile."""

    catalog = load_service_profiles(config.home, config.service_profiles)
    profile = catalog.select(profile_id, end)
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    book, partial, target = _atomic_book(config, "service", f"SERVICE PERFORMANCE  /  {profile.label.upper()}", start, end, output)
    periods = _comparison_periods(start, end)
    all_start, all_end = min(p.start for p in periods), max(p.end for p in periods)
    raw_rows = _service_rows(conn, profile, all_start, all_end)
    period_rows = []
    for period in periods:
        scoped = [row for row in raw_rows if period.start.isoformat() <= str(row["business_date"])[:10] <= period.end.isoformat()]
        metrics = _service_aggregate(scoped, profile, metric_catalog, period.end)
        period_rows.append((period.label, period.start, period.end, metrics["offered"], metrics["answered"], metrics["service_level"], metrics["availability"], metrics["aht_seconds"]))
    by_label = {row[0]: row for row in period_rows}
    latest, mtd, prior = by_label["Latest day"], by_label["Current MTD"], by_label["Prior-month same days"]
    latest_method = _profile_metric(metric_catalog, profile, profile.service_level_metric, end)
    forecast_marks = ",".join("?" for _ in profile.service_scopes)
    forecast_total = conn.execute(
        f"""SELECT coalesce(sum(volume_forecast),0) FROM mart.forecast_hour
            WHERE business_date=? AND service_scope IN ({forecast_marks})""",
        [end, *profile.service_scopes],
    ).fetchone()[0]
    forecast_method = metric_catalog.method_for(
        "forecast_attainment", end, {"lob": profile.service_scopes[0]},
    )
    forecast_evaluation = evaluate_metric(
        forecast_method, {"actual_volume": latest[3], "forecast_volume": forecast_total},
    ) if forecast_method else None
    families = tuple(system.lower() for system in profile.source_systems) + ("forecast",)
    status, status_text = _source_state(conn, families, end)
    book.dashboard(
        [
            KpiCard("Latest-day service level", latest[5], "percent", f"Target {latest_method.target:.0%}" if latest_method.target is not None else "No target"),
            KpiCard("Latest-day availability", latest[6], "percent", "Answered / offered"),
            KpiCard("Latest-day offered", latest[3], "integer", f"Forecast {forecast_total:,.0f}"),
            KpiCard("Forecast attainment", forecast_evaluation.value if forecast_evaluation else None, "percent", "Actual offered / forecast"),
            KpiCard("Latest-day answered", latest[4], "integer", "All mapped queues"),
            KpiCard("Latest-day AHT seconds", latest[7], "decimal", "Weighted by answered contacts"),
            KpiCard("MTD service level", mtd[5], "percent", _delta(mtd[5], prior[5], True)),
            KpiCard("MTD availability", mtd[6], "percent", _delta(mtd[6], prior[6], True)),
        ],
        status,
        status_text,
        ["Period", "Start", "End", "Offered", "Answered", "Service Level %", "Availability %", "AHT Seconds"],
        period_rows,
        [
            "This report includes only queues mapped into the selected service profile; queue membership is controlled in queue_mapping.csv.",
            f"{profile.label}: the extract's within-{rulebook.target_seconds}s counter is evaluated by {profile.service_level_metric}.{latest_method.method_id} from metric_catalog.toml.",
            "Service availability is answered / offered. It is never agent availability or adherence.",
            "Forecast contributes forecast only. APBE/APFR/APDE contribute actual performance only.",
        ],
        (("Service Level", 5), ("Availability", 6)),
    )

    hourly: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in raw_rows:
        if start.isoformat() <= str(row["business_date"])[:10] <= end.isoformat():
            group = profile.group_for(row.get("queue"))
            hourly[(str(row["business_date"])[:10], str(row["hour_start"]), group)].append(row)
    hourly_rows = []
    for (business_date, hour_start, group), rows in sorted(hourly.items()):
        row_date = date.fromisoformat(business_date)
        metrics = _service_aggregate(rows, profile, metric_catalog, row_date)
        hourly_rows.append((business_date, hour_start, group, metrics["offered"], metrics["answered"], metrics["within_target"], metrics["short_abandoned"], metrics["service_level"], metrics["availability"], metrics["aht_seconds"], metrics["service_method"], metrics["service_state"]))
    headers = ["business_date", "hour_start", "service_group", "offered", "answered", "answered_within_target", "short_abandoned", "service_level", "service_availability", "aht_seconds", "service_method", "service_state"]
    ws = book.table("INTRADAY", f"{profile.label} hourly service", "Ford / Toyota / Chery are separated by configured queue-name groups; totals remain additive.", headers, hourly_rows)
    if hourly_rows:
        state_col = headers.index("service_state")
        ws.conditional_format(4, state_col, 3 + len(hourly_rows), state_col, {"type": "text", "criteria": "containing", "value": "BELOW_TARGET", "format": book.report.error})
    detail_headers = ["business_date", "interval_start", "source_system", "queue", "service_scope", "designation", "mapping_status", "offered", "answered", "abandoned", "short_abandoned", "answered_within_target", "handled_seconds", "source_file"]
    detail_rows = [tuple(row.get(header) for header in detail_headers) for row in raw_rows if start.isoformat() <= str(row["business_date"])[:10] <= end.isoformat()]
    book.table("QUEUE_DETAIL", "Mapped queue evidence", "Compact governed queue intervals for reconciliation; no original extract rows are modified.", detail_headers, detail_rows)
    action_rows = [row for row in hourly_rows if row[-1] == "BELOW_TARGET"]
    book.table("ACTIONS", "Service intervals below target", "Prioritize intervals with high offered volume and validate staffing/forecast before escalation.", headers, action_rows)
    book.definitions([
        ("Service Level", f"({latest_method.numerator}) / ({latest_method.denominator})", "Contract performance", f"Governed method {latest_method.method_id}; ratio of summed counters"),
        ("Service Availability", "Answered / offered", "Ability of the service to answer demand", "Not an agent metric"),
        ("Forecast Attainment", "Actual offered / forecast", "Demand realization", "Forecast and actual joined only by mapped comparison scope"),
        ("AHT", "Total handled seconds / answered contacts", "Workload driver", "Weighted, never average of interval AHT"),
    ])
    book.audit(_audit_rows(conn, config, "service", start, end, [
        ("Service profile", profile.profile_id, f"{catalog.version} / {catalog.sha256}"),
        ("Included scopes", ", ".join(profile.service_scopes), ", ".join(profile.source_systems)),
        ("Service metric", profile.service_level_metric, f"{metric_catalog.version} / {metric_catalog.sha256}"),
    ]))
    return _finish(book, partial, target)


def build_attendance_today_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build only the live daily attendance calling product."""

    report_day = end
    book, partial, target = _atomic_book(config, "attendance", "ATTENDANCE TODAY", report_day, report_day, output)
    totals = conn.execute(
        """SELECT count(*),
                  coalesce(sum(CASE WHEN requires_call THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_NO_SHOW' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_LATE' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_NOT_SEEN_NOW' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN source_loaded=false THEN 1 ELSE 0 END),0)
           FROM mart.attendance_agent_day
           WHERE business_date=? AND assignment_type NOT IN ('Off','Planned absence')""",
        [report_day],
    ).fetchone()
    scheduled, call_now, no_show, late, not_seen, missing = totals
    status, status_text = _source_state(conn, ("fte", "start_end", "lilo", "agent_status"), report_day)
    if missing:
        status, status_text = "INCOMPLETE", f"{missing:,} scheduled row(s) do not have complete attendance evidence"
    book.dashboard(
        [
            KpiCard("Scheduled working", scheduled, "integer", "Agent-day rows"),
            KpiCard("Call now", call_now, "integer", "Current action queue"),
            KpiCard("Confirmed no-show", no_show, "integer", "Only after completed shift"),
            KpiCard("Late", late, "integer", "Beyond configured tolerance"),
            KpiCard("Not seen yet", not_seen, "integer", "Provisional current-day state"),
            KpiCard("Missing evidence", missing, "integer", "Unknown, never no-show"),
        ],
        status,
        status_text,
        ["Result", "Agents"],
        [("Scheduled working", scheduled), ("Call now", call_now), ("Confirmed no-show", no_show), ("Late", late), ("Not seen yet", not_seen), ("Missing evidence", missing)],
        [
            "This workbook is the live calling list; it is not an adherence report.",
            "An unfinished shift can be late or not seen, but can never be marked as early leave.",
            "No-show requires a completed working shift, a loaded blank LILO row, and no active Agent Status evidence.",
            "Use Attendance Corrections only after the operating day is complete.",
        ],
        (("Agents", 1),),
        "column",
    )
    headers, rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, scheduled_start, scheduled_end, shift_state,
                  call_action, attendance_result, actual_first_seen, actual_last_seen,
                  uncoded_late_minutes, no_show_minutes, actual_evidence,
                  source_loaded, is_provisional, evaluation_as_of
           FROM mart.attendance_agent_day
           WHERE business_date=? AND requires_call=true
           ORDER BY CASE call_action WHEN 'CALL_NO_SHOW' THEN 1 WHEN 'CALL_LATE' THEN 2 ELSE 3 END,
                    scheduled_start, lob, team_leader, agent_name""",
        [report_day],
    )
    ws = book.table("ACTIONS", "People to contact now", "Operational queue ordered by severity and scheduled start.", headers, rows)
    if rows:
        col = headers.index("call_action")
        ws.conditional_format(4, col, 3 + len(rows), col, {"type": "text", "criteria": "containing", "value": "CALL_NO_SHOW", "format": book.report.error})
    trend_headers, trend_rows = _query(
        conn,
        """SELECT business_date, count(*) AS scheduled_working,
                  sum(CASE WHEN requires_call THEN 1 ELSE 0 END) AS call_actions,
                  sum(CASE WHEN call_action='CALL_NO_SHOW' THEN 1 ELSE 0 END) AS no_shows,
                  sum(CASE WHEN call_action='CALL_LATE' THEN 1 ELSE 0 END) AS late,
                  sum(CASE WHEN source_loaded=false THEN 1 ELSE 0 END) AS missing_evidence
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND assignment_type NOT IN ('Off','Planned absence')
           GROUP BY business_date ORDER BY business_date""",
        [start, end],
    )
    book.table("TREND", "Attendance action trend", "Daily counts for the requested range; today remains provisional until shifts close.", trend_headers, trend_rows)
    book.definitions([
        ("Call no-show", "Completed shift + loaded blank LILO + no active status", "Call and validate absence", "Missing source is never a no-show"),
        ("Call late", "First observed evidence after scheduled start plus tolerance", "Contact / record explanation", "Current shift may still be running"),
        ("Not seen now", "Shift started, no observed evidence yet", "Immediate operational check", "Always provisional"),
        ("Early leave", "Last observed evidence before completed scheduled end", "Historical correction only", "Never evaluated before shift end"),
    ])
    book.audit(_audit_rows(conn, config, "attendance", report_day, report_day))
    return _finish(book, partial, target)


def build_staffing_coverage_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the standalone LOB/language staffing decision product."""

    report_day = end
    book, partial, target = _atomic_book(config, "staffing", "STAFFING & COVERAGE", report_day, report_day, output)
    totals = conn.execute(
        """SELECT max(staffing_gap_fte),
                  coalesce(sum(CASE WHEN staffing_state IN ('GAP','PARTIAL_GAP') THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN staffing_state='DATA_MISSING' THEN 1 ELSE 0 END),0),
                  avg(scheduled_fte), avg(observed_fte), avg(productive_fte)
           FROM mart.staffing_interval WHERE business_date=?""",
        [report_day],
    ).fetchone()
    peak_gap, gap_intervals, missing_intervals, avg_scheduled, avg_observed, avg_productive = totals
    status, status_text = _source_state(conn, ("fte", "start_end", "lilo", "agent_status"), report_day)
    if missing_intervals:
        status, status_text = "INCOMPLETE", f"{missing_intervals:,} interval(s) have missing attendance evidence"
    book.dashboard(
        [
            KpiCard("Peak gap FTE", peak_gap, "decimal", "Largest 15-minute deficit"),
            KpiCard("Gap intervals", gap_intervals, "integer", "LOB/language intervals"),
            KpiCard("Average scheduled FTE", avg_scheduled, "decimal"),
            KpiCard("Average observed FTE", avg_observed, "decimal"),
            KpiCard("Average productive FTE", avg_productive, "decimal"),
            KpiCard("Missing intervals", missing_intervals, "integer", "Unknown, not zero"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [("Peak gap FTE", peak_gap), ("Gap intervals", gap_intervals), ("Average scheduled FTE", avg_scheduled), ("Average observed FTE", avg_observed), ("Average productive FTE", avg_productive)],
        [
            "Staffing is grouped by roster LOB and language; it is not mixed with service or attendance calling.",
            "Observed FTE is calculated from agent-seconds inside each 15-minute interval.",
            "Future and missing-evidence intervals stay unknown rather than becoming false staffing deficits.",
        ],
    )
    headers, rows = _query(
        conn,
        """SELECT business_date, interval_start, interval_end, lob, language,
                  scheduled_agents, observed_agents, productive_agents,
                  scheduled_fte, elapsed_scheduled_fte, observed_fte,
                  productive_fte, staffing_variance_fte, staffing_gap_fte,
                  staffing_state, evidence_basis, evaluation_as_of
           FROM mart.staffing_interval WHERE business_date=?
           ORDER BY interval_start, lob, language""",
        [report_day],
    )
    book.table("INTRADAY", "15-minute staffing coverage", "Complete daily evidence at LOB/language grain.", headers, rows)
    actions = [row for row in rows if row[headers.index("staffing_state")] in {"GAP", "PARTIAL_GAP", "DATA_MISSING"}]
    book.table("ACTIONS", "Staffing exceptions", "Intervals requiring redeployment, investigation or evidence repair.", headers, actions)
    book.definitions([
        ("Scheduled FTE", "Scheduled agent-seconds / 900", "Expected interval capacity", "Roster LOB/language"),
        ("Observed FTE", "Observed agent-seconds / 900", "Actual presence", "LILO + Agent Status evidence"),
        ("Productive FTE", "Productive-status seconds / 900", "Available handling capacity", "Not adherence"),
        ("Staffing gap", "MAX(0, elapsed scheduled FTE - observed FTE)", "Immediate staffing deficit", "Blank for future/missing evidence"),
    ])
    book.audit(_audit_rows(conn, config, "staffing", report_day, report_day))
    return _finish(book, partial, target)


def build_final_absence_product_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the final corrected Verint absence and shrinkage product."""

    book, partial, target = _atomic_book(config, "absence", "FINAL ABSENCE & SHRINKAGE", start, end, output)
    totals = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0), coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0), coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0), coalesce(sum(final_unmapped_minutes),0),
                  coalesce(sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END),0)
           FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?""",
        [start, end],
    ).fetchone()
    planned, absence, vacation, unpaid, shrinkage, unmapped, absence_days = totals
    status, status_text = _source_state(conn, ("fte", "start_end", "activities"), end, final=True)
    if unmapped:
        status, status_text = "INCOMPLETE", f"{unmapped / 60:,.2f} hour(s) of Verint Activities are unmapped"
    period_rows = []
    for period in _comparison_periods(start, end):
        row = conn.execute(
            """SELECT coalesce(sum(planned_net_minutes),0), coalesce(sum(final_absence_minutes),0),
                      coalesce(sum(final_vacation_minutes),0), coalesce(sum(final_shrinkage_minutes),0)
               FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?""",
            [period.start, period.end],
        ).fetchone()
        period_rows.append((period.label, period.start, period.end, row[0] / 60, row[1] / 60, _ratio(row[1], row[0]), row[2] / 60, _ratio(row[3], row[0])))
    book.dashboard(
        [
            KpiCard("Planned net hours", planned / 60, "decimal"),
            KpiCard("Final absence hours", absence / 60, "decimal"),
            KpiCard("Final absence rate", _ratio(absence, planned), "percent"),
            KpiCard("Absence agent-days", absence_days, "integer"),
            KpiCard("Vacation hours", vacation / 60, "decimal"),
            KpiCard("Unpaid hours", unpaid / 60, "decimal"),
            KpiCard("Shrinkage rate", _ratio(shrinkage, planned), "percent"),
            KpiCard("Unmapped hours", unmapped / 60, "decimal", "Must be resolved before payroll"),
        ],
        status,
        status_text,
        ["Period", "Start", "End", "Planned Hours", "Absence Hours", "Absence Rate %", "Vacation Hours", "Shrinkage Rate %"],
        period_rows,
        [
            "This is the final corrected ledger: Verint Activities only, clipped to StartEndTimes schedule boundaries.",
            "LILO and Agent Status detect operational gaps but do not create final payroll categories.",
            "Overlapping Activities are unioned before totals; never sum event evidence directly.",
        ],
        (("Absence Rate", 5), ("Shrinkage Rate", 7)),
    )
    trend_headers, trend_rows = _query(
        conn,
        """SELECT business_date, sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0 THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0 THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate
           FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date ORDER BY business_date""",
        [start, end],
    )
    book.table("TREND", "Final absence daily trend", "Daily overlap-safe final counters.", trend_headers, trend_rows)
    detail_headers, detail_rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS final_absence_hours,
                  final_vacation_minutes/60.0 AS final_vacation_hours,
                  final_unpaid_minutes/60.0 AS final_unpaid_hours,
                  final_shrinkage_minutes/60.0 AS final_shrinkage_hours,
                  final_unmapped_minutes/60.0 AS final_unmapped_hours,
                  final_absence_rate, final_absence_day, final_ledger_status
           FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name""",
        [start, end],
    )
    book.table("AGENT_DETAIL", "Final absence agent ledger", "Authoritative agent/day grain for payroll and Operations.", detail_headers, detail_rows)
    exceptions = [row for row in detail_rows if row[detail_headers.index("final_unmapped_hours")] or row[detail_headers.index("final_ledger_status")] == "UNMAPPED_REVIEW"]
    book.table("EXCEPTIONS", "Final-ledger exceptions", "Unmapped classifications must be resolved before payroll use.", detail_headers, exceptions)
    book.definitions([
        ("Final absence rate", "Unioned classified absence minutes / planned net minutes", "Payroll absence", "Activities-only final ledger"),
        ("Final shrinkage rate", "Unioned configured shrinkage minutes / planned net minutes", "Capacity loss", "Includes configured nonproductive categories"),
        ("Planned net", "Scheduled span capped by configured standard day", "Common denominator", "StartEndTimes only"),
        ("Unmapped", "Verint Activity without an approved classification", "Rulebook action", "Blocks final status"),
    ])
    book.audit(_audit_rows(conn, config, "absence", start, end))
    return _finish(book, partial, target)


def build_attendance_corrections_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build a completed-day correction queue with the Verint-style timeline."""

    from .governed_workbooks import _add_shift_view, _completed_timeline_day

    report_day = _completed_timeline_day(conn, end)
    book, partial, target = _atomic_book(
        config, "corrections", "ATTENDANCE CORRECTIONS", report_day, report_day, output,
    )
    gap_count, gap_minutes, agents = conn.execute(
        """SELECT count(*), coalesce(sum(residual_minutes),0), count(DISTINCT agent_id)
           FROM mart.correction_residual_segment WHERE business_date=?""",
        [report_day],
    ).fetchone()
    missing = conn.execute(
        """SELECT count(*) FROM mart.attendance_agent_day
           WHERE business_date=? AND assignment_type NOT IN ('Off','Planned absence')
             AND source_loaded=false""",
        [report_day],
    ).fetchone()[0]
    status, status_text = _source_state(
        conn, ("fte", "start_end", "lilo", "agent_status", "activities"), report_day, final=True,
    )
    if missing:
        status, status_text = "INCOMPLETE", f"{missing:,} scheduled row(s) lack complete observed evidence"
    book.dashboard(
        [
            KpiCard("Residual segments", gap_count, "integer"),
            KpiCard("Residual gap hours", gap_minutes / 60 if gap_minutes else 0, "decimal"),
            KpiCard("Agents to review", agents, "integer"),
            KpiCard("Missing evidence", missing, "integer"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [("Residual segments", gap_count), ("Residual gap hours", gap_minutes / 60 if gap_minutes else 0), ("Agents to review", agents), ("Missing evidence", missing)],
        [
            "Only completed-day residual gaps appear here; an unfinished current-day shift cannot become an early-leave correction.",
            "GAPS is the editable action list. Pale-blue columns are the only human decisions accepted for import.",
            "SHIFT_VIEW visualizes the full planned-versus-observed shift; the exact segment data is exported to the template model package.",
            "Verint Activities verify whether an observed gap is corrected; they never create the original gap.",
        ],
    )
    headers, rows = _query(
        conn,
        """SELECT r.residual_id, r.correction_id, r.business_date,
                  r.agent_id, c.agent_name, c.team_leader, c.ops_manager,
                  c.lob, d.language, c.scheduled_start, c.scheduled_end,
                  c.detected_issue, r.residual_start AS gap_start,
                  r.residual_end AS gap_end, r.residual_minutes AS gap_minutes,
                  c.confidence, r.suggested_activity, r.observed_source,
                  r.verint_reconciliation, c.verint_activity,
                  c.verint_overlap_minutes, c.validation_status,
                  c.confirmed_activity, c.owner, c.comment, c.injected_date,
                  r.source_file
           FROM mart.correction_residual_segment r
           JOIN mart.correction_candidate c ON c.correction_id=r.correction_id
           LEFT JOIN core.dim_agent d ON d.agent_id=r.agent_id
           WHERE r.business_date=?
             AND coalesce(c.validation_status,'Open') NOT IN ('Injected','Rejected')
           ORDER BY c.priority, r.residual_minutes DESC, r.agent_id, r.residual_start""",
        [report_day],
    )
    ws = book.table(
        "GAPS", "Residual gaps ready for Verint correction",
        "Edit only Confirmed Activity, Validation Status, Owner, Comment and Injected Date; then import this workbook.",
        headers, rows,
        editable_headers={"Confirmed Activity", "Validation Status", "Owner", "Comment", "Injected Date"},
    )
    if rows:
        display = [value.replace("_", " ").title().replace("Id", "ID") for value in headers]
        status_col = display.index("Validation Status")
        ws.data_validation(4, status_col, 3 + len(rows), status_col, {
            "validate": "list", "source": ["Open", "Validated", "Injected", "Rejected"],
        })
        suggestions = sorted({str(row[headers.index("suggested_activity")]) for row in rows if row[headers.index("suggested_activity")]})
        activities = list(dict.fromkeys([*suggestions, "Absent", "Late", "Early Leave", "Unpaid Leave", "Vacation"]))[:20]
        if activities:
            activity_col = display.index("Confirmed Activity")
            ws.data_validation(4, activity_col, 3 + len(rows), activity_col, {"validate": "list", "source": activities})

    timeline_headers, timeline_rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader,
                  ops_manager, lob, language, scheduled_start, scheduled_end,
                  segment_start, segment_end, segment_minutes, planned_state,
                  actual_status, actual_category, mismatch_type, is_gap,
                  observed_source, source_file, evaluation_as_of
           FROM mart.shift_timeline_segment t WHERE business_date=?
             AND EXISTS (
                 SELECT 1 FROM mart.correction_residual_segment r
                 WHERE r.business_date=t.business_date AND r.agent_id=t.agent_id
             )
           ORDER BY agent_id, segment_start""",
        [report_day],
    )
    timeline_dicts = [dict(zip(timeline_headers, row)) for row in timeline_rows]
    _add_shift_view(book.report, timeline_dicts, report_day)
    book.tables.append(ModelTable("TIMELINE", timeline_headers, timeline_rows))
    book.definitions([
        ("Observed gap", "Scheduled time minus unioned LILO/Agent Status evidence", "Correction candidate", "Activities cannot create the gap"),
        ("Residual gap", "Observed gap minus unioned corrected Verint Activities", "Minutes still requiring review", "Overlap-safe"),
        ("Current-day tail", "Future portion of an unfinished shift", "No action", "Never early leave"),
        ("Correction ID", "Stable action key", "Import decisions", "One correction may have several residual segments"),
    ])
    book.audit(_audit_rows(conn, config, "corrections", report_day, report_day))
    return _finish(book, partial, target)
