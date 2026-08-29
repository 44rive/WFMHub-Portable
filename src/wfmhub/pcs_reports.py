"""Build the standalone Agent call-performance and PCS workbook."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from .config import Config
from .database import DatabaseConnection
from .report_packs import report_pack, report_pack_folder
from .reports import ExcelReport, _query


def _start(report: ExcelReport, config: Config, start: date, end: date, generated: datetime) -> None:
    ws = report.workbook.add_worksheet("START_HERE")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "WFMHub Agent PCS", report.title)
    ws.merge_range(
        "A2:H2",
        f"Generated {generated:%Y-%m-%d %H:%M}; period {start:%Y-%m-%d} to {end:%Y-%m-%d}",
        report.subtitle,
    )
    lines = [
        "Call-by-call rows are FTE-scoped, parsed, typed and deduplicated in SQLite; raw calls are not copied into this workbook.",
        f"PCS scored questions: {', '.join('Q'+str(value) for value in config.pcs.scored_questions)}; comments: {', '.join('Q'+str(value) for value in config.pcs.comment_questions) or 'none'}.",
        "PCS Average: each response averages its valid configured scores, then every response has equal weight in the agent average.",
        f"Valid scale {config.pcs.minimum_score:g}-{config.pcs.maximum_score:g}; top box >= {config.pcs.top_box_minimum:g}; low score <= {config.pcs.low_score_maximum:g}.",
        "A blank average means no valid response. It is not a zero score.",
        "Use SOURCE_HEALTH and DATA_QUALITY before sending the report.",
    ]
    for row, line in enumerate(lines, 4):
        ws.write(row - 1, 0, line, report.body)
    ws.set_column("A:A", 120)
    ws.set_column("B:H", 3)


def _summary(report: ExcelReport, conn: DatabaseConnection, start: date, end: date) -> None:
    ws = report.workbook.add_worksheet("SUMMARY")
    ws.hide_gridlines(2)
    ws.merge_range("A1:F1", "Agent PCS summary", report.title)
    ws.merge_range("A2:F2", "Weighted period totals from agent/day PCS marts.", report.subtitle)
    metrics = [
        ("Agent-days with calls", "SELECT count(*) FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Handled calls", "SELECT coalesce(sum(handled_calls),0) FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("PCS-enabled calls", "SELECT coalesce(sum(pcs_enabled_calls),0) FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Survey responses", "SELECT coalesce(sum(survey_responses),0) FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Response rate", "SELECT CASE WHEN sum(pcs_enabled_calls)>0 THEN 1.0*sum(survey_responses)/sum(pcs_enabled_calls) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", True),
        ("PCS average", "SELECT CASE WHEN sum(pcs_score_count)>0 THEN 1.0*sum(pcs_score_sum)/sum(pcs_score_count) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Q1 average", "SELECT CASE WHEN sum(q1_response_count)>0 THEN 1.0*sum(q1_score_sum)/sum(q1_response_count) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Q2 average", "SELECT CASE WHEN sum(q2_response_count)>0 THEN 1.0*sum(q2_score_sum)/sum(q2_response_count) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", False),
        ("Top-box %", "SELECT CASE WHEN sum(survey_responses)>0 THEN 1.0*sum(top_box_responses)/sum(survey_responses) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", True),
        ("Low-score %", "SELECT CASE WHEN sum(survey_responses)>0 THEN 1.0*sum(low_score_responses)/sum(survey_responses) END FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?", True),
    ]
    ws.write("A4", "KPI", report.header)
    ws.write("B4", "Value", report.header)
    for index, (label, sql, percentage) in enumerate(metrics, 4):
        value = conn.execute(sql, [start, end]).fetchone()[0]
        ws.write(index, 0, label, report.body)
        fmt = report.percent if percentage else report.decimal if "average" in label.lower() else report.integer
        ws.write(index, 1, value, fmt)
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 18)


def _survey_score_sql(config: Config) -> tuple[str, str]:
    minimum, maximum = config.pcs.minimum_score, config.pcs.maximum_score
    valid = [
        f"CASE WHEN question_{number}_score BETWEEN {minimum:g} AND {maximum:g} THEN 1 ELSE 0 END"
        for number in sorted(set(config.pcs.scored_questions))
    ]
    values = [
        f"CASE WHEN question_{number}_score BETWEEN {minimum:g} AND {maximum:g} THEN question_{number}_score ELSE 0 END"
        for number in sorted(set(config.pcs.scored_questions))
    ]
    return " + ".join(valid), " + ".join(values)


def _python_recipes(report: ExcelReport) -> None:
    ws = report.workbook.add_worksheet("PYTHON_RECIPES")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "Python in Excel recipes", report.title)
    ws.merge_range("A2:H2", "Copy a recipe into a Python in Excel cell when that feature is enabled by Microsoft 365 policy.", report.subtitle)
    recipes = [
        ("Load agent summary", 'df = xl("tblAgentPcs[#All]", headers=True)\ndf'),
        ("Team-leader PCS", 'df = xl("tblAgentPcs[#All]", headers=True)\ndf.groupby("Team Leader", dropna=False)["PCS Average"].mean().sort_values()'),
        ("PCS distribution", 'df = xl("tblAgentPcs[#All]", headers=True)\ndf["PCS Average"].dropna().plot(kind="hist", bins=10, title="Agent PCS distribution")'),
        ("Daily trend", 'daily = xl("tblDailyTrend[#All]", headers=True)\ndaily.groupby("Date")["PCS Average"].mean().plot(title="Daily PCS")'),
    ]
    ws.write("A4", "Recipe", report.header)
    ws.write("B4", "Python", report.header)
    wrap = report.workbook.add_format({"font_name": "Consolas", "font_size": 9, "text_wrap": True, "valign": "top"})
    for row, (name, code) in enumerate(recipes, 4):
        ws.write(row, 0, name, report.body)
        ws.write(row, 1, code, wrap)
        ws.set_row(row, 55)
    ws.set_column("A:A", 24)
    ws.set_column("B:B", 110)


def build_pcs_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    generated = datetime.now()
    pack = report_pack("quality_pcs")
    output = (
        output
        or report_pack_folder(config, pack.key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _start(report, config, start, end, generated)
        _summary(report, conn, start, end)
        headers, rows = _query(
            conn,
            """SELECT agent_id, agent_name, team_leader, ops_manager, lob,
                      market, language, location, sum(call_legs) call_legs,
                      sum(handled_calls) handled_calls, sum(inbound_calls) inbound_calls,
                      sum(outbound_calls) outbound_calls,
                      CASE WHEN sum(handled_calls)>0 THEN 1.0*sum(talk_seconds)/sum(handled_calls) END average_talk_seconds,
                      CASE WHEN sum(handled_calls)>0 THEN 1.0*sum(hold_seconds)/sum(handled_calls) END average_hold_seconds,
                      CASE WHEN sum(handled_calls)>0 THEN 1.0*sum(wrap_seconds)/sum(handled_calls) END average_wrap_seconds,
                      CASE WHEN sum(handled_calls)>0 THEN 1.0*sum(handle_seconds)/sum(handled_calls) END average_handle_seconds,
                      sum(pcs_enabled_calls) pcs_enabled_calls,
                      sum(survey_responses) survey_responses,
                      CASE WHEN sum(pcs_status_calls)>0 THEN 1.0*sum(pcs_participation_responses)/sum(pcs_status_calls) END response_rate,
                      CASE WHEN sum(q1_response_count)>0 THEN 1.0*sum(q1_score_sum)/sum(q1_response_count) END q1_average,
                      CASE WHEN sum(q2_response_count)>0 THEN 1.0*sum(q2_score_sum)/sum(q2_response_count) END q2_average,
                      CASE WHEN sum(pcs_score_count)>0 THEN 1.0*sum(pcs_score_sum)/sum(pcs_score_count) END pcs_average,
                      CASE WHEN sum(survey_responses)>0 THEN 1.0*sum(top_box_responses)/sum(survey_responses) END top_box_percent,
                      CASE WHEN sum(survey_responses)>0 THEN 1.0*sum(low_score_responses)/sum(survey_responses) END low_score_percent,
                      sum(comments_count) comments_count
               FROM mart.agent_pcs_day
               WHERE business_date BETWEEN ? AND ?
               GROUP BY agent_id, agent_name, team_leader, ops_manager, lob,
                        market, language, location
               ORDER BY CASE WHEN sum(pcs_score_count)=0 THEN 1 ELSE 0 END,
                        pcs_average, agent_name
               LIMIT ?""",
            [start, end, config.report_limits.get("max_pcs_agent_rows", 100000)],
        )
        report.add_table_sheet("AGENT_PCS", "Agent PCS period summary", "One row per in-scope agent with calls in the selected period.", headers, rows)
        headers, rows = _query(
            conn,
            """SELECT business_date, agent_id, agent_name, team_leader,
                      ops_manager, lob, handled_calls, inbound_calls,
                      outbound_calls, average_handle_seconds, pcs_enabled_calls,
                      survey_responses, response_rate, q1_average, q2_average,
                      pcs_average, top_box_percent, low_score_percent, comments_count
               FROM mart.agent_pcs_day
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_name
               LIMIT ?""",
            [start, end, config.report_limits.get("max_pcs_agent_rows", 100000)],
        )
        report.add_table_sheet("DAILY_TREND", "Daily Agent PCS", "One row per agent/day after call deduplication.", headers, rows)

        score_count, score_sum = _survey_score_sql(config)
        comment_test = " OR ".join(
            f"coalesce(trim(question_{number}), '') <> ''" for number in config.pcs.comment_questions
        ) or "0"
        headers, rows = _query(
            conn,
            f"""WITH responses AS (
                    SELECT c.business_date, c.call_start, c.call_reference_number,
                           c.call_id, c.agent_id, coalesce(d.canonical_name,c.agent_name) agent_name,
                           d.team_leader, coalesce(d.lob,c.lob) lob, c.queue, c.service,
                           c.language, c.question_1, c.question_2, c.question_3,
                           CASE WHEN ({score_count})>0 THEN 1.0*({score_sum})/({score_count}) END pcs_score,
                           CASE WHEN {comment_test} THEN 1 ELSE 0 END has_comment,
                           c.source_file
                    FROM core.clean_call_leg c LEFT JOIN core.dim_agent d USING(agent_id)
                    WHERE c.business_date BETWEEN ? AND ?
                      AND upper(coalesce(c.call_direction,''))='I'
                      AND coalesce(c.post_call_survey_mode,'')=?
                )
                SELECT business_date, call_start, call_reference_number, call_id,
                       agent_id, agent_name, team_leader, lob, queue, service,
                       language, question_1, question_2, question_3, pcs_score,
                       source_file
                FROM responses
                WHERE pcs_score IS NOT NULL OR has_comment=1
                ORDER BY CASE WHEN pcs_score IS NULL THEN 2 ELSE 1 END,
                         pcs_score, business_date, call_start
                LIMIT ?""",
            [start, end, config.pcs.survey_mode, config.report_limits.get("max_pcs_exception_rows", 100000)],
        )
        report.add_table_sheet("SURVEY_RESPONSES", "PCS responses and comments", "Bounded response detail; use clean-data export for the full call dataset.", headers, rows)
        headers, rows = _query(
            conn,
            """SELECT detected_at, source_family, source_file, business_date,
                      agent_id, issue_type, severity, details
               FROM meta.quality_issue
               WHERE (business_date IS NULL OR business_date BETWEEN ? AND ?)
                 AND source_family IN ('calls','pcs')
               ORDER BY CASE severity WHEN 'ERROR' THEN 1 ELSE 2 END, business_date""",
            [start, end],
        )
        report.add_table_sheet("DATA_QUALITY", "PCS data quality", "Scope, parsing and response checks for this pack.", headers, rows, exception_column="Severity")
        headers, rows = _query(
            conn,
            """SELECT source_family, expected_path, newest_file, newest_business_date,
                      modified_at, loaded_at, row_count, rejected_count,
                      scoped_out_count, status, details
               FROM mart.source_health WHERE source_family IN ('fte','calls')
               ORDER BY source_family""",
        )
        report.add_table_sheet("SOURCE_HEALTH", "PCS source health", "FTE scope and Call-by-Call coverage.", headers, rows, exception_column="Status")
        _python_recipes(report)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output
