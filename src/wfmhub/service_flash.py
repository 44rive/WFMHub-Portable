"""Configuration-driven daily service flashes reconstructed from Book1."""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .mapping import QueueMapping, load_queue_mapping
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import publish_report, report_current_path
from .reports import COLORS
from .rules import Rulebook, load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook


def _marks(values: Sequence[str]) -> str:
    return ",".join("?" for _ in values)


def _as_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return None


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    if numerator is None or denominator is None or float(denominator) == 0:
        return None
    return float(numerator) / float(denominator)


def _profile_comparison_scopes(
    profile: ServiceProfile,
    mapping: QueueMapping,
) -> tuple[str, ...]:
    return mapping.comparison_scopes_for(profile.service_scopes)


def _profile_method(
    catalog: MetricCatalog,
    profile: ServiceProfile,
    metric_id: str,
    on_date: date,
):
    methods = {}
    for service_scope in profile.service_scopes:
        for source_system in profile.flash_source_systems:
            method = catalog.method_for(
                metric_id, on_date,
                {"lob": service_scope, "source_system": source_system},
            )
            if method is None:
                raise ValueError(
                    f"Flash profile {profile.profile_id!r} has no {metric_id!r} "
                    f"method for {service_scope}/{source_system} on {on_date}"
                )
            methods[(method.method_id, method.effective_from, method.priority)] = method
    if len(methods) != 1:
        names = ", ".join(key[0] for key in methods)
        raise ValueError(
            f"Flash profile {profile.profile_id!r} crosses incompatible "
            f"{metric_id} methods: {names}"
        )
    return next(iter(methods.values()))


def _aggregate(
    rows: Iterable[dict[str, Any]],
    profile: ServiceProfile,
    catalog: MetricCatalog,
    on_date: date,
) -> dict[str, Any] | None:
    values = list(rows)
    if not values:
        return None
    components = {
        name: sum(float(row.get(name) or 0) for row in values)
        for name in (
            "offered", "answered", "abandoned", "short_abandoned",
            "answered_within_target", "handled_seconds",
        )
    }
    service = evaluate_metric(
        _profile_method(catalog, profile, profile.service_level_metric, on_date),
        components,
    )
    availability = evaluate_metric(
        _profile_method(catalog, profile, profile.availability_metric, on_date),
        components,
    )
    abandon = evaluate_metric(
        _profile_method(catalog, profile, "abandon_rate", on_date), components,
    )
    aht = evaluate_metric(
        _profile_method(catalog, profile, profile.aht_metric, on_date), components,
    )
    return {
        **components,
        "service_level": service.value,
        "service_target": service.method.target,
        "service_method": service.method.method_id,
        "service_state": service.state,
        "availability": availability.value,
        "abandon_rate": abandon.value,
        "aht_seconds": aht.value,
    }


def _profile_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    cursor = conn.execute(
        f"""SELECT business_date, hour_start, source_system, service_scope,
                   comparison_scope, queue, designation, language, offered,
                   answered, abandoned, short_abandoned,
                   answered_within_target, talk_seconds, hold_seconds,
                   wrap_seconds, handled_seconds, service_level,
                   service_availability, abandon_rate, aht_seconds, call_legs,
                   transferred_legs, source_files
            FROM mart.call_service_hour
            WHERE business_date BETWEEN ? AND ?
              AND service_scope IN ({_marks(profile.service_scopes)})
              AND source_system IN ({_marks(profile.flash_source_systems)})
            ORDER BY business_date, hour_start, queue""",
        [start, end, *profile.service_scopes, *profile.flash_source_systems],
    )
    headers = [item[0] for item in cursor.description]
    return [dict(zip(headers, row)) for row in cursor.fetchall()]


def _included_in_flash_total(
    profile: ServiceProfile,
    row: dict[str, Any],
) -> bool:
    return (
        not profile.flash_total_groups
        or profile.group_for(row.get("queue")) in profile.flash_total_groups
    )


def _forecast_by_hour(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    mapping: QueueMapping,
    report_day: date,
) -> dict[int, float]:
    scopes = _profile_comparison_scopes(profile, mapping)
    rows = conn.execute(
        f"""SELECT hour_start, sum(volume_forecast)
            FROM mart.forecast_hour
            WHERE business_date=? AND comparison_scope IN ({_marks(scopes)})
            GROUP BY hour_start ORDER BY hour_start""",
        [report_day, *scopes],
    ).fetchall()
    result: dict[int, float] = {}
    for hour_start, value in rows:
        parsed = _as_datetime(hour_start)
        if parsed is not None and value is not None:
            result[parsed.hour] = float(value)
    return result


def _workforce_by_hour(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    report_day: date,
) -> dict[int, dict[str, float]]:
    lob_marks = _marks(profile.staffing_lobs)
    attendance = conn.execute(
        f"""SELECT agent_id, scheduled_start, scheduled_end, assignment_type,
                   planned_work_minutes
            FROM mart.attendance_agent_day
            WHERE business_date=? AND lob IN ({lob_marks})""",
        [report_day, *profile.staffing_lobs],
    ).fetchall()
    absence = conn.execute(
        f"""SELECT agent_id, category, event_start, event_end, counts_as_absence
            FROM mart.absence_event
            WHERE business_date=? AND lob IN ({lob_marks})""",
        [report_day, *profile.staffing_lobs],
    ).fetchall()
    output: dict[int, dict[str, float]] = {}
    for hour in range(profile.operating_start_hour, profile.operating_end_hour + 1):
        left = datetime.combine(report_day, datetime.min.time()) + timedelta(hours=hour)
        right = left + timedelta(hours=1)
        planned = {
            str(agent_id)
            for agent_id, raw_start, raw_end, assignment_type, planned_work in attendance
            if str(assignment_type or "").casefold() != "off"
            and float(planned_work or 0) > 0
            and (_as_datetime(raw_start) or right) < right
            and (_as_datetime(raw_end) or left) > left
        }
        categories: dict[str, set[str]] = defaultdict(set)
        absent: set[str] = set()
        for agent_id, category, raw_start, raw_end, counts_as_absence in absence:
            event_start, event_end = _as_datetime(raw_start), _as_datetime(raw_end)
            if event_start is None or event_end is None or event_start >= right or event_end <= left:
                continue
            agent_key = str(agent_id)
            category_key = str(category or "").upper()
            categories[category_key].add(agent_key)
            if counts_as_absence:
                absent.add(agent_key)
        short = categories.get("SICKNESS_SHORT", set())
        long = categories.get("SICKNESS_LONG", set())
        late_early = categories.get("LATE", set()) | categories.get("EARLY_LEAVE", set())
        output[hour] = {
            "planned_hc": float(len(planned)),
            "short_sickness_hc": float(len(short)),
            "long_sickness_hc": float(len(long)),
            "late_early_hc": float(len(late_early)),
            "absence_hc": float(len(absent)),
            "absence_rate": _ratio(len(absent), len(planned)),
        }
    return output


def _hourly_model(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    mapping: QueueMapping,
    metrics: MetricCatalog,
    report_day: date,
    all_rows: Sequence[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any] | None, dict[str, dict[str, Any] | None], int | None]:
    rows = [
        row for row in all_rows
        if str(row["business_date"])[:10] == report_day.isoformat()
        and _included_in_flash_total(profile, row)
    ]
    by_hour: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        hour = _as_datetime(row.get("hour_start"))
        if (
            hour is not None
            and profile.operating_start_hour <= hour.hour <= profile.operating_end_hour
        ):
            by_hour[hour.hour].append(row)
    forecast = _forecast_by_hour(conn, profile, mapping, report_day)
    workforce = _workforce_by_hour(conn, profile, report_day) if profile.flash_layout == "workforce" else {}
    cutoff = max(by_hour) if by_hour else None
    hourly: list[dict[str, Any]] = []
    for hour in range(profile.operating_start_hour, profile.operating_end_hour + 1):
        source = by_hour.get(hour, [])
        aggregate = _aggregate(source, profile, metrics, report_day)
        group_values = {
            group.label: _aggregate(
                [row for row in source if profile.group_for(row.get("queue")) == group.label],
                profile, metrics, report_day,
            )
            for group in profile.groups
        }
        forecast_value = forecast.get(hour)
        actual_value = aggregate["offered"] if aggregate is not None else None
        state = (
            ("FUTURE" if report_day >= date.today() else "AFTER CUTOFF")
            if cutoff is not None and hour > cutoff
            else "NO MAPPED CALLS" if aggregate is None
            else "FORECAST MISSING" if forecast_value is None
            else "READY"
        )
        hourly.append({
            "profile_id": profile.profile_id, "flash": profile.label,
            "business_date": report_day, "hour": hour,
            "hour_label": f"{hour:02d}:00", "forecast": forecast_value,
            "forecast_attainment": _ratio(actual_value, forecast_value),
            "data_state": state, **(aggregate or {
                "offered": None, "answered": None, "abandoned": None,
                "short_abandoned": None, "answered_within_target": None,
                "handled_seconds": None, "service_level": None,
                "service_target": _profile_method(
                    metrics, profile, profile.service_level_metric, report_day,
                ).target,
                "service_method": _profile_method(
                    metrics, profile, profile.service_level_metric, report_day,
                ).method_id,
                "service_state": "NO_SAMPLE", "availability": None,
                "abandon_rate": None, "aht_seconds": None,
            }),
            "groups": group_values,
            **workforce.get(hour, {}),
        })
    through_hours = [row for row in hourly if cutoff is not None and row["hour"] <= cutoff]
    total_source = [row for hour in by_hour if cutoff is not None and hour <= cutoff for row in by_hour[hour]]
    total = _aggregate(total_source, profile, metrics, report_day)
    if total is not None:
        total["forecast"] = sum(
            float(row["forecast"] or 0) for row in through_hours if row["forecast"] is not None
        ) if any(row["forecast"] is not None for row in through_hours) else None
        total["forecast_attainment"] = _ratio(total["offered"], total["forecast"])
    group_totals = {
        group.label: _aggregate(
            [row for row in total_source if profile.group_for(row.get("queue")) == group.label],
            profile, metrics, report_day,
        )
        for group in profile.groups
    }
    return hourly, total, group_totals, cutoff


def _formats(book: DecisionWorkbook) -> dict[str, Any]:
    add = book.report.workbook.add_format
    return {
        "card_label": add({
            "font_name": "Aptos", "font_size": 9, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["teal"],
            "align": "center", "valign": "vcenter", "border": 1,
            "border_color": COLORS["thin"],
        }),
        "card_value": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "border": 1,
            "border_color": COLORS["thin"],
        }),
        "card_integer": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0",
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_percent": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0.0%",
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_seconds": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": '0 "s"',
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_note": add({
            "font_name": "Aptos", "font_size": 8, "font_color": COLORS["muted"],
            "bg_color": COLORS["canvas"], "align": "center", "valign": "vcenter",
            "text_wrap": True, "border": 1, "border_color": COLORS["thin"],
        }),
        "total": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "border": 1, "border_color": COLORS["thin"],
        }),
        "total_integer": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": "#,##0", "border": 1, "border_color": COLORS["thin"],
        }),
        "total_percent": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": "0.0%", "border": 1, "border_color": COLORS["thin"],
        }),
        "total_seconds": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": '0 "s"', "border": 1, "border_color": COLORS["thin"],
        }),
    }


def _write_card(
    ws,
    formats: dict[str, Any],
    column: int,
    label: str,
    value: Any,
    kind: str,
    note: str,
) -> None:
    ws.merge_range(4, column, 4, column + 1, label.upper(), formats["card_label"])
    value_format = formats.get(f"card_{kind}", formats["card_value"])
    ws.merge_range(5, column, 6, column + 1, value if value is not None else "—", value_format)
    ws.merge_range(7, column, 7, column + 1, note, formats["card_note"])


def _table_value_format(book: DecisionWorkbook, header: str):
    lowered = header.casefold()
    if (
        lowered.startswith("sl ")
        or any(token in lowered for token in (
            "availability", "deviation", "service level", "tsl", "absence rate",
        ))
    ):
        return book.report.percent
    if "aht" in lowered:
        return book.report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"],
            "num_format": '0 "s"', "bottom": 1, "bottom_color": COLORS["thin"],
        })
    if any(token in lowered for token in ("forecast", "actual", "handled", "sickness", "planned", "absence hc", "late/early")):
        return book.report.integer
    return book.report.body


def _flash_columns(
    profile: ServiceProfile,
    hourly: Sequence[dict[str, Any]],
) -> tuple[list[str], list[list[Any]], int, int, int]:
    if profile.flash_layout == "oem_split":
        headers = [
            "Hour", "Volume Forecasted", "Volume Ford", "Volume Toyota",
            "SL Ford", "Availability Ford", "Availability Toyota", "AHT",
        ]
        rows = []
        for row in hourly:
            ford = row["groups"].get("Ford") or {}
            toyota = row["groups"].get("Toyota") or {}
            rows.append([
                row["hour_label"], row["forecast"], ford.get("offered"),
                toyota.get("offered"), ford.get("service_level"),
                ford.get("availability"), toyota.get("availability"),
                row["aht_seconds"],
            ])
        return headers, rows, 1, 2, 4
    if profile.flash_layout == "workforce":
        headers = [
            "Hour", "Volume Forecasted", "Volume Actual", "Volume Handled",
            "Volume Handled in SL", "Deviation", "Availability", "TSL",
            "AHT", "Planned HC", "Short Sickness", "Long Sickness",
            "Late/Early Leave", "Absence HC", "Absence Rate", "Data State",
        ]
        rows = [[
            row["hour_label"], row["forecast"], row["offered"], row["answered"],
            row["answered_within_target"], row["forecast_attainment"],
            row["availability"], row["service_level"], row["aht_seconds"],
            row.get("planned_hc"), row.get("short_sickness_hc"),
            row.get("long_sickness_hc"), row.get("late_early_hc"),
            row.get("absence_hc"), row.get("absence_rate"), row["data_state"],
        ] for row in hourly]
        return headers, rows, 1, 2, 7
    headers = [
        "Hour", "Volume Forecasted", "Volume Actual", "Volume Handled",
        "Volume Handled in SL", "Deviation", "Availability", "TSL", "AHT",
        "Data State",
    ]
    rows = [[
        row["hour_label"], row["forecast"], row["offered"], row["answered"],
        row["answered_within_target"], row["forecast_attainment"],
        row["availability"], row["service_level"], row["aht_seconds"],
        row["data_state"],
    ] for row in hourly]
    return headers, rows, 1, 2, 7


def _flash_cards(
    profile: ServiceProfile,
    total: dict[str, Any] | None,
    groups: dict[str, dict[str, Any] | None],
    hourly: Sequence[dict[str, Any]],
) -> list[tuple[str, Any, str, str]]:
    value = total or {}
    if profile.flash_layout == "oem_split":
        ford, toyota = groups.get("Ford") or {}, groups.get("Toyota") or {}
        return [
            ("Availability OEM", value.get("availability"), "percent", "Handled / actual"),
            ("Availability Ford", ford.get("availability"), "percent", "Mapped Ford queues"),
            ("Availability Toyota", toyota.get("availability"), "percent", "Toyota and Lexus"),
            ("Deviation", value.get("forecast_attainment"), "percent", "Actual / forecast through cutoff"),
            ("TSL OEM", value.get("service_level"), "percent", value.get("service_method") or "Configured method"),
            ("TSL Ford", ford.get("service_level"), "percent", "Mapped Ford queues"),
            ("TSL Toyota", toyota.get("service_level"), "percent", "Toyota and Lexus"),
        ]
    if profile.flash_layout == "workforce":
        planned = max((row.get("planned_hc") or 0 for row in hourly), default=0)
        absent = max((row.get("absence_hc") or 0 for row in hourly), default=0)
        return [
            ("Dispatch", "N/C", "value", "Source not configured"),
            ("Follow-up", "N/C", "value", "Source not configured"),
            ("Mailbox BNL", "N/C", "value", "Source not configured"),
            ("Absence Rate HC", _ratio(absent, planned), "percent", "Peak absent HC / peak planned HC"),
            ("Deviation", value.get("forecast_attainment"), "percent", "Actual / forecast through cutoff"),
            ("Availability", value.get("availability"), "percent", "Handled / actual"),
            ("TSL", value.get("service_level"), "percent", value.get("service_method") or "Configured method"),
            ("AHT", value.get("aht_seconds"), "seconds", "Weighted handled seconds"),
        ]
    return [
        ("Forecast", value.get("forecast"), "integer", "Through latest actual hour"),
        ("Actual", value.get("offered"), "integer", "Unique mapped interactions"),
        ("Handled", value.get("answered"), "integer", "Answered interactions"),
        ("Handled in SL", value.get("answered_within_target"), "integer", "Answered inside threshold"),
        ("Deviation", value.get("forecast_attainment"), "percent", "Actual / forecast through cutoff"),
        ("Availability", value.get("availability"), "percent", "Handled / actual"),
        ("TSL", value.get("service_level"), "percent", value.get("service_method") or "Configured method"),
        ("AHT", value.get("aht_seconds"), "seconds", "Weighted handled seconds"),
    ]


def _add_flash_sheet(
    book: DecisionWorkbook,
    profile: ServiceProfile,
    report_day: date,
    hourly: Sequence[dict[str, Any]],
    total: dict[str, Any] | None,
    group_totals: dict[str, dict[str, Any] | None],
    cutoff: int | None,
) -> None:
    ws = book.report.workbook.add_worksheet(profile.flash_sheet)
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.merge_range("A1:Q1", f"FLASH  /  {profile.label.upper()}", book.report.title)
    cutoff_text = f"through {cutoff:02d}:59" if cutoff is not None else "no mapped call interactions"
    ws.merge_range(
        "A2:Q2",
        f"{report_day:%Y-%m-%d}  |  Call-by-Call actuals {cutoff_text}  |  generated {book.generated:%Y-%m-%d %H:%M}",
        book.report.subtitle,
    )
    ws.set_row(0, 34)
    ws.set_row(1, 21)
    formats = _formats(book)
    for index, card in enumerate(_flash_cards(profile, total, group_totals, hourly)):
        _write_card(ws, formats, index * 2, *card)
    headers, display_rows, forecast_col, actual_col, sl_col = _flash_columns(profile, hourly)
    table_row = 10
    for column, header in enumerate(headers):
        ws.write(table_row, column, header, book.report.header)
    for row_index, values in enumerate(display_rows, table_row + 1):
        for column, value in enumerate(values):
            fmt = _table_value_format(book, headers[column])
            if value is None:
                ws.write_blank(row_index, column, None, fmt)
            else:
                ws.write(row_index, column, value, fmt)
    if display_rows:
        table_name = "tbl" + re.sub(r"[^A-Za-z0-9]", "", profile.flash_sheet)
        ws.add_table(table_row, 0, table_row + len(display_rows), len(headers) - 1, {
            "name": table_name,
            "style": "Table Style Light 9",
            "columns": [{"header": header, "header_format": book.report.header} for header in headers],
        })
    total_row = table_row + 1 + len(display_rows)
    label = f"Day through {cutoff:02d}:59" if cutoff is not None else "Day total unavailable"
    ws.write(total_row, 0, label, formats["total"])
    total_values = total or {}
    if profile.flash_layout == "oem_split":
        ford, toyota = (
            group_totals.get("Ford") or {}, group_totals.get("Toyota") or {},
        )
        values = [
            total_values.get("forecast"), ford.get("offered"), toyota.get("offered"),
            ford.get("service_level"), ford.get("availability"),
            toyota.get("availability"), total_values.get("aht_seconds"),
        ]
    elif profile.flash_layout == "workforce":
        values = [
            total_values.get("forecast"), total_values.get("offered"),
            total_values.get("answered"), total_values.get("answered_within_target"),
            total_values.get("forecast_attainment"), total_values.get("availability"),
            total_values.get("service_level"), total_values.get("aht_seconds"),
            max((row.get("planned_hc") or 0 for row in hourly), default=0),
            max((row.get("short_sickness_hc") or 0 for row in hourly), default=0),
            max((row.get("long_sickness_hc") or 0 for row in hourly), default=0),
            max((row.get("late_early_hc") or 0 for row in hourly), default=0),
            max((row.get("absence_hc") or 0 for row in hourly), default=0),
            max((row.get("absence_rate") or 0 for row in hourly), default=0),
            "READY" if total else "INCOMPLETE",
        ]
    else:
        values = [
            total_values.get("forecast"), total_values.get("offered"),
            total_values.get("answered"), total_values.get("answered_within_target"),
            total_values.get("forecast_attainment"), total_values.get("availability"),
            total_values.get("service_level"), total_values.get("aht_seconds"),
            "READY" if total else "INCOMPLETE",
        ]
    for column, value in enumerate(values, 1):
        header = headers[column]
        lowered = header.casefold()
        fmt = (
            formats["total_percent"] if lowered.startswith("sl ") or any(
                token in lowered
                for token in ("availability", "deviation", "tsl", "absence rate")
            )
            else formats["total_seconds"] if "aht" in lowered
            else formats["total_integer"] if isinstance(value, (int, float)) and value is not None
            else formats["total"]
        )
        if value is None:
            ws.write_blank(total_row, column, None, fmt)
        else:
            ws.write(total_row, column, value, fmt)

    data_last_row = table_row + len(display_rows)
    if display_rows:
        chart = book.report.workbook.add_chart({"type": "column"})
        volume_series = (
            (
                ("Forecast", forecast_col, COLORS["muted"]),
                ("Ford", 2, COLORS["teal"]),
                ("Toyota", 3, COLORS["gold"]),
            )
            if profile.flash_layout == "oem_split"
            else (
                ("Forecast", forecast_col, COLORS["muted"]),
                ("Actual", actual_col, COLORS["teal"]),
            )
        )
        for label_text, column, color in volume_series:
            chart.add_series({
                "name": label_text,
                "categories": [profile.flash_sheet, table_row + 1, 0, data_last_row, 0],
                "values": [profile.flash_sheet, table_row + 1, column, data_last_row, column],
                "fill": {"color": color}, "border": {"none": True},
            })
        line = book.report.workbook.add_chart({"type": "line"})
        line.add_series({
            "name": "TSL",
            "categories": [profile.flash_sheet, table_row + 1, 0, data_last_row, 0],
            "values": [profile.flash_sheet, table_row + 1, sl_col, data_last_row, sl_col],
            "y2_axis": True, "line": {"color": COLORS["gold"], "width": 2.25},
        })
        chart.combine(line)
        chart.set_title({"name": "Hourly demand and service"})
        chart.set_legend({"position": "bottom"})
        chart.set_y2_axis({"num_format": "0%", "min": 0, "max": 1})
        chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
        chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
        ws.insert_chart(table_row, len(headers) + 1, chart, {"x_scale": 1.05, "y_scale": 1.0})

    target = total_values.get("service_target")
    if target is not None and display_rows:
        ws.conditional_format(table_row + 1, sl_col, data_last_row, sl_col, {
            "type": "cell", "criteria": "<", "value": target,
            "format": book.report.error,
        })
    for column, header in enumerate(headers):
        width = 13
        if header == "Hour":
            width = 10
        elif header == "Data State":
            width = 20
        elif len(header) > 17:
            width = 18
        ws.set_column(column, column, width)
    ws.set_column(len(headers) + 1, len(headers) + 10, 11)
    ws.freeze_panes(table_row + 1, 1)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.repeat_rows(0, table_row)


def _add_control_sheet(
    book: DecisionWorkbook,
    summaries: Sequence[tuple[ServiceProfile, dict[str, Any] | None, int | None]],
    report_day: date,
) -> None:
    ws = book.report.workbook.add_worksheet("CONTROL")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.merge_range("A1:J1", "SERVICE FLASH CONTROL", book.report.title)
    ws.merge_range(
        "A2:J2",
        f"Daily control for {report_day:%Y-%m-%d}  |  actuals: mapped Call-by-Call interactions  |  forecast: Verint",
        book.report.subtitle,
    )
    headers = [
        "Flash", "Cutoff", "Forecast", "Actual", "Handled", "Deviation",
        "Availability", "TSL", "AHT", "Status",
    ]
    for col, header in enumerate(headers):
        ws.write(4, col, header, book.report.header)
    for offset, (profile, total, cutoff) in enumerate(summaries, 5):
        value = total or {}
        status = (
            "CALL DATA MISSING" if total is None
            else "FORECAST MISSING" if value.get("forecast") is None
            else "READY"
        )
        row = [
            profile.label, f"{cutoff:02d}:59" if cutoff is not None else None,
            value.get("forecast"), value.get("offered"), value.get("answered"),
            value.get("forecast_attainment"), value.get("availability"),
            value.get("service_level"), value.get("aht_seconds"), status,
        ]
        for col, item in enumerate(row):
            fmt = (
                book.report.percent if col in {5, 6, 7}
                else book.report.decimal if col == 8
                else book.report.integer if col in {2, 3, 4}
                else book.report.body
            )
            if item is None:
                ws.write_blank(offset, col, None, fmt)
            else:
                ws.write(offset, col, item, fmt)
        ws.write_url(offset, 0, f"internal:'{profile.flash_sheet}'!A1", book.report.body, profile.label)
    if summaries:
        ws.add_table(4, 0, 4 + len(summaries), len(headers) - 1, {
            "name": "tblFlashControl", "style": "Table Style Light 9",
            "columns": [{"header": header, "header_format": book.report.header} for header in headers],
        })
        ws.conditional_format(5, 9, 4 + len(summaries), 9, {
            "type": "text", "criteria": "containing", "value": "MISSING",
            "format": book.report.error,
        })
        chart = book.report.workbook.add_chart({"type": "column"})
        for label, column, color in (
            ("Forecast", 2, COLORS["muted"]), ("Actual", 3, COLORS["teal"]),
        ):
            chart.add_series({
                "name": label,
                "categories": ["CONTROL", 5, 0, 4 + len(summaries), 0],
                "values": ["CONTROL", 5, column, 4 + len(summaries), column],
                "fill": {"color": color}, "border": {"none": True},
            })
        chart.set_title({"name": "Forecast versus actual through cutoff"})
        chart.set_legend({"position": "bottom"})
        chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart("A17", chart, {"x_scale": 1.25, "y_scale": 1.1})
    notes = [
        "Open a Flash name to jump to its hourly sheet.",
        "Deviation follows the reference workbook: actual offered / forecast through the latest actual hour.",
        "Availability means handled / offered. TSL follows the effective metric configured for each Flash.",
        "No mapped calls and missing forecasts remain blank; the workbook never turns missing evidence into zero.",
    ]
    ws.write("A10", "OPERATING NOTES", book.report.section)
    for index, note in enumerate(notes, 10):
        ws.merge_range(index, 0, index, 9, note, book.report.note)
    ws.set_column("A:A", 23)
    ws.set_column("B:J", 16)
    ws.freeze_panes(5, 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)


def _flat_hour_rows(
    profiles: Sequence[ServiceProfile],
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]],
) -> tuple[list[str], list[tuple[Any, ...]]]:
    headers = [
        "profile_id", "flash", "business_date", "hour", "forecast", "offered",
        "answered", "abandoned", "short_abandoned", "answered_within_target",
        "forecast_attainment", "availability", "service_level", "service_target",
        "service_method", "abandon_rate", "aht_seconds", "planned_hc",
        "short_sickness_hc", "long_sickness_hc", "late_early_hc",
        "absence_hc", "absence_rate", "ford_offered", "ford_answered",
        "ford_service_level", "toyota_offered", "toyota_answered",
        "toyota_service_level", "data_state",
    ]
    rows: list[tuple[Any, ...]] = []
    for profile in profiles:
        for row in hourly_by_profile[profile.profile_id]:
            ford, toyota = (
                row["groups"].get("Ford") or {}, row["groups"].get("Toyota") or {},
            )
            rows.append(tuple([
                row.get("profile_id"), row.get("flash"), row.get("business_date"),
                row.get("hour"), row.get("forecast"), row.get("offered"),
                row.get("answered"), row.get("abandoned"), row.get("short_abandoned"),
                row.get("answered_within_target"), row.get("forecast_attainment"),
                row.get("availability"), row.get("service_level"),
                row.get("service_target"), row.get("service_method"),
                row.get("abandon_rate"), row.get("aht_seconds"),
                row.get("planned_hc"), row.get("short_sickness_hc"),
                row.get("long_sickness_hc"), row.get("late_early_hc"),
                row.get("absence_hc"), row.get("absence_rate"),
                ford.get("offered"), ford.get("answered"), ford.get("service_level"),
                toyota.get("offered"), toyota.get("answered"), toyota.get("service_level"),
                row.get("data_state"),
            ]))
    return headers, rows


def build_service_flashes_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Build all Book1 Flash layouts from mapped Call-by-Call interactions."""

    catalog = load_service_profiles(config.home, config.service_profiles)
    profiles = [profile for profile in catalog.profiles if profile.active_on(end)]
    profiles.sort(key=lambda profile: (profile.display_order, profile.profile_id))
    if profile_id is not None:
        # Retain CLI validation without changing the stable all-Flash workbook.
        catalog.select(profile_id, end)
    mapping = load_queue_mapping(config.queue_mapping)
    metrics = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    generated = datetime.now()
    target = (output or report_current_path(config, "service")).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    book = DecisionWorkbook(
        partial, config, "service", "SERVICE FLASHES", start, end, generated,
    )
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    summaries: list[tuple[ServiceProfile, dict[str, Any] | None, int | None]] = []
    group_totals_by_profile: dict[str, dict[str, dict[str, Any] | None]] = {}
    try:
        for profile in profiles:
            source_rows = _profile_rows(conn, profile, start, end)
            hourly, total, group_totals, cutoff = _hourly_model(
                conn, profile, mapping, metrics, end, source_rows,
            )
            hourly_by_profile[profile.profile_id] = hourly
            group_totals_by_profile[profile.profile_id] = group_totals
            summaries.append((profile, total, cutoff))
        _add_control_sheet(book, summaries, end)
        for profile, total, cutoff in summaries:
            _add_flash_sheet(
                book, profile, end, hourly_by_profile[profile.profile_id], total,
                group_totals_by_profile[profile.profile_id], cutoff,
            )
        headers, rows = _flat_hour_rows(profiles, hourly_by_profile)
        book.table(
            "FLASH_DATA", "Flash clean hourly data",
            "One profile/hour for the report day. Actual demand is deduplicated by Call-by-Call interaction key.",
            headers, rows,
        )
        with mapping.file.open("r", encoding="utf-8-sig", newline="") as handle:
            source_mapping = list(csv.DictReader(handle))
        mapping_headers = [
            "mapping_type", "source_system", "source_value", "service_scope",
            "comparison_scope", "designation", "flash_group", "used_by_flash",
        ]
        mapping_rows = []
        for row in source_mapping:
            if str(row.get("mapping_type") or "").strip().lower() != "queue":
                continue
            mapped = mapping.map_actual(
                row.get("source_system"), row.get("source_value"), None, None,
            )
            matching_profile = next(
                (
                    profile for profile in profiles
                    if mapped.service_scope in profile.service_scopes
                ),
                None,
            )
            flash_group = (
                matching_profile.group_for(row.get("source_value"))
                if matching_profile is not None else None
            )
            used = (
                matching_profile is not None
                and (
                    not matching_profile.flash_total_groups
                    or flash_group in matching_profile.flash_total_groups
                )
            )
            mapping_rows.append((
                row.get("mapping_type"), row.get("source_system"),
                row.get("source_value"), mapped.service_scope,
                mapped.comparison_scope, mapped.designation, flash_group,
                "YES" if used else "NO",
            ))
        book.table(
            "QUEUE_MAP", "Queue-to-Flash control",
            "The exact configured queues admitted from Call-by-Call. Edit queue_mapping.csv, validate, then refresh.",
            mapping_headers, mapping_rows,
        )
        exception_headers = [
            "flash", "business_date", "hour", "issue", "actual", "forecast",
            "service_level", "target", "action",
        ]
        exception_rows = []
        for profile in profiles:
            for row in hourly_by_profile[profile.profile_id]:
                issue = None
                action = None
                if row["data_state"] == "FORECAST MISSING":
                    issue, action = "FORECAST MISSING", "Load or map the Verint forecast extract"
                elif row["data_state"] == "NO MAPPED CALLS" and row["hour"] <= (max((item["hour"] for item in hourly_by_profile[profile.profile_id] if item["offered"] is not None), default=-1)):
                    issue, action = "NO MAPPED CALLS", "Confirm zero demand or review queue mapping"
                elif row.get("service_level") is not None and row.get("service_target") is not None and row["service_level"] < row["service_target"]:
                    issue, action = "BELOW TSL TARGET", "Review demand, staffing and long waits"
                if issue:
                    exception_rows.append((
                        profile.label, row["business_date"], row["hour_label"], issue,
                        row["offered"], row["forecast"], row["service_level"],
                        row["service_target"], action,
                    ))
        book.table(
            "EXCEPTIONS", "Flash exceptions",
            "Only missing evidence and below-target service intervals requiring review.",
            exception_headers, exception_rows,
        )
        oem_groups = next(
            (
                profile.flash_total_groups for profile in profiles
                if profile.flash_layout == "oem_split"
            ),
            (),
        )
        book.definitions([
            ("Volume Actual", "One unique inbound interaction in a mapped Flash scope", "Demand", "Transferred call legs are not double-counted inside the same Flash scope"),
            ("OEM visible scope", " and ".join(oem_groups) or "Every configured group", "Matches the Book1 OEM image", "Other mapped groups remain in the hub but are excluded from OEM Flash totals"),
            ("Volume Handled", "Mapped interaction with an inbound handled agent leg", "Service availability numerator", "Agent may be outside the FTE roster; the queue is the service boundary"),
            ("Volume Handled in SL", f"Handled interaction with queue wait <= {rulebook.target_seconds} seconds", "TSL numerator", "Threshold is editable in wfm_rules.toml"),
            ("Short Abandon", f"Unanswered interaction with queue wait < {rulebook.short_abandon_seconds} seconds", "Adjusted TSL denominator", "Configured centrally"),
            ("Deviation", "Actual offered / forecast through the latest actual hour", "Demand tracking", "The label follows Book1; mathematically this is forecast attainment"),
            ("Availability", "Handled / actual offered", "Service availability", "Not agent availability and not adherence"),
            ("TSL", "Configured ratio of summed counters", "Service-level control", "Ford OEM uses gross TSL; other profiles use adjusted TSL"),
            ("AHT", "Sum of inbound talk + hold + wrap / handled interactions", "Workload", "Weighted; never an average of hourly averages"),
            ("Ford NL workload cards", "Dispatch, Follow-up and Mailbox BNL remain N/C", "Data integrity", "Book1 provides labels but no governed source or formula; values are not invented"),
        ])
        clean_calls, unique_interactions = conn.execute(
            """SELECT count(*), count(DISTINCT interaction_key)
               FROM core.clean_call_leg WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()
        mapped_offered = conn.execute(
            """SELECT coalesce(sum(offered),0) FROM mart.call_service_hour
               WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()[0]
        book.audit([
            ("Report", "service", "All Book1 Flash profiles"),
            ("Report day", end, "Visible Flash sheets use the selected end date"),
            ("Selected data period", f"{start} to {end}", "Model boundary"),
            ("Clean Call-by-Call legs", clean_calls, "After stable call-leg deduplication"),
            ("Unique clean interactions", unique_interactions, "Before queue mapping"),
            ("Mapped Flash offered", mapped_offered, "One per interaction/comparison scope"),
            ("Queue mapping", mapping.file.name, mapping.sha256),
            ("Service profiles", catalog.version, catalog.sha256),
            ("OEM visible groups", " | ".join(oem_groups) or "ALL", "Configured in service_profiles.toml"),
            ("Metric catalog", metrics.version, metrics.sha256),
            ("Rulebook", rulebook.version, rulebook.sha256),
            ("Design reference", "TOLEARN/Book1.xlsx", "Four pasted Flash references reconstructed as native Excel"),
            ("Prepared by", "Anass ASSRI", "WFM"),
        ])
        book.close()
        publish_report(config, "service", partial, target, generated)
    except Exception:
        try:
            book.report.close()
        except Exception:
            pass
        partial.unlink(missing_ok=True)
        raise
    return target
