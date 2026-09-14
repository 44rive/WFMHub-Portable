"""Read-only projections for the local WFM Manager Workbench.

The browser never receives raw extracts and never calculates business KPIs.
Every projection below reads governed marts and effective configuration only.
"""

from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Iterable

from .capacity_mapping import CapacityMapping, load_capacity_mapping
from .config import Config
from .database import DatabaseConnection, connect
from .metrics import MetricCatalog, load_metric_catalog
from .service_profiles import ServiceProfileCatalog, load_service_profiles


PRESENT_RESULTS = {
    "Present", "Present - partial time off", "Late", "Early leave",
    "Late + early leave", "Shift in progress", "Late - shift in progress",
}
NO_SHOW_RESULTS = {"No show", "No show - partial time off"}


def _iso(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat(sep="T", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def _number(value: Any) -> float:
    return float(value or 0)


def _ratio(numerator: float, denominator: float) -> float | None:
    return numerator / denominator if denominator else None


def _rows(cursor) -> list[dict[str, Any]]:
    names = [item[0] for item in cursor.description]
    return [dict(zip(names, row)) for row in cursor.fetchall()]


def _day(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


@dataclass(frozen=True)
class DashboardFilter:
    start: date
    end: date
    management_lob: str | None = None
    planning_group: str | None = None
    staff_type: str | None = None
    team_leader: str | None = None
    agent_id: str | None = None
    scenario_fte: float = 0.0

    @classmethod
    def from_values(
        cls,
        start: str | None,
        end: str | None,
        fallback: date,
        management_lob: str | None = None,
        planning_group: str | None = None,
        staff_type: str | None = None,
        team_leader: str | None = None,
        agent_id: str | None = None,
        scenario_fte: str | float | None = None,
    ) -> "DashboardFilter":
        left = date.fromisoformat(start) if start else fallback
        right = date.fromisoformat(end) if end else left
        if left > right:
            raise ValueError("Start date cannot be after end date")
        if (right - left).days > 730:
            raise ValueError("The local console accepts at most 731 days per view")

        def clean(value: str | None) -> str | None:
            text = str(value or "").strip()
            if len(text) > 160:
                raise ValueError("A filter value is too long")
            return None if not text or text.casefold() == "all" else text

        try:
            adjustment = float(scenario_fte or 0)
        except (TypeError, ValueError) as exc:
            raise ValueError("Scenario FTE adjustment must be a number") from exc
        if not -100 <= adjustment <= 100:
            raise ValueError("Scenario FTE adjustment must be between -100 and 100")

        return cls(
            left, right, clean(management_lob), clean(planning_group),
            clean(staff_type), clean(team_leader), clean(agent_id), adjustment,
        )


class DashboardData:
    """Small read-only application service over the durable WFMHub database."""

    def __init__(self, config: Config):
        self.config = config
        self.capacity: CapacityMapping = load_capacity_mapping(config.capacity_mapping)
        self.services: ServiceProfileCatalog = load_service_profiles(
            config.home, config.service_profiles,
        )
        self.metrics: MetricCatalog = load_metric_catalog(
            config.home, config.metric_catalog,
        )

    def _connect(self) -> DatabaseConnection:
        return connect(self.config, read_only=True)

    def _workforce_lob(self, lob: Any, on_date: date) -> str:
        text = str(lob or "Unmapped").strip() or "Unmapped"
        candidates = {
            item.management_lob
            for item in self.services.profiles
            if item.active_on(on_date)
            and any(text.casefold() == value.casefold() for value in item.staffing_lobs)
        }
        if len(candidates) == 1:
            return next(iter(candidates))
        mapped = self.capacity.map_schedule(text, None)
        return mapped.management_lob or text

    def _capacity_for_attendance(self, row: dict[str, Any]):
        return self.capacity.map_schedule(row.get("lob"), row.get("assignment"))

    def _service_target(self, management_lob: str, on_date: date) -> float | None:
        profiles = [
            profile for profile in self.services.profiles
            if profile.management_lob == management_lob and profile.active_on(on_date)
        ]
        if len(profiles) != 1:
            return None
        method = self.metrics.method_for(
            profiles[0].service_level_metric, on_date,
            {"lob": management_lob, "source_system": "CALL_BY_CALL"},
        )
        return method.target if method else None

    def latest_date(self) -> date:
        conn = self._connect()
        try:
            # Forecasts and published schedules can extend weeks into the
            # future. They must never move the default operational day beyond
            # the newest actual call/status evidence.
            value = conn.execute(
                """SELECT max(business_date) FROM (
                     SELECT business_date FROM mart.call_service_15min
                     UNION ALL SELECT extract_date AS business_date FROM raw.agent_status
                     UNION ALL SELECT extract_date AS business_date FROM raw.lilo
                   )"""
            ).fetchone()[0]
            if value is None:
                value = conn.execute(
                    """SELECT max(business_date) FROM (
                         SELECT business_date FROM mart.attendance_agent_day
                         UNION ALL SELECT business_date FROM mart.staffing_interval
                         UNION ALL SELECT business_date FROM mart.forecast_interval
                       )"""
                ).fetchone()[0]
        finally:
            conn.close()
        return _day(value) if value else date.today()

    def meta(self) -> dict[str, Any]:
        conn = self._connect()
        try:
            latest = self.latest_date()
            run = conn.execute(
                """SELECT finished_at, status, files_loaded, files_skipped,
                          files_failed, details
                   FROM meta.refresh_run ORDER BY started_at DESC LIMIT 1"""
            ).fetchone()
            sources = _rows(conn.execute(
                """SELECT source_family, newest_business_date, row_count,
                          rejected_count, status
                   FROM mart.source_health ORDER BY source_family"""
            ))
            assignments = _rows(conn.execute(
                """SELECT DISTINCT agent_id, lob, assignment
                   FROM mart.attendance_agent_day
                   WHERE trim(coalesce(agent_id,''))<>''"""
            ))
            agents = _rows(conn.execute(
                """SELECT agent_id, canonical_name, team_leader, ops_manager,
                          lob, market, language, location
                   FROM core.dim_agent
                   WHERE trim(coalesce(agent_id,''))<>''
                   ORDER BY canonical_name, agent_id"""
            ))
        finally:
            conn.close()

        assignment_by_agent: dict[str, set[tuple[str, str, str]]] = defaultdict(set)
        for row in assignments:
            mapped = self.capacity.map_schedule(row["lob"], row["assignment"])
            assignment_by_agent[str(row["agent_id"])].add(
                (mapped.management_lob, mapped.planning_group, mapped.staff_type),
            )
        agent_rows = []
        for row in agents:
            agent_id = str(row["agent_id"])
            capacities = assignment_by_agent.get(agent_id) or {
                (self._workforce_lob(row["lob"], latest), "", ""),
            }
            for management_lob, planning_group, staff_type in sorted(capacities):
                agent_rows.append({
                    "id": agent_id,
                    "name": row["canonical_name"] or agent_id,
                    "team_leader": row["team_leader"] or "Unassigned",
                    "ops_manager": row["ops_manager"] or "Unassigned",
                    "management_lob": management_lob,
                    "planning_group": planning_group,
                    "staff_type": staff_type,
                })

        capacity_rows = {
            (
                result.management_lob, result.planning_group, result.staff_type,
            )
            for _, _, result in self.capacity.forecast_rows
        } | {
            (
                result.management_lob, result.planning_group, result.staff_type,
            )
            for result in self.capacity.schedule_rows.values()
        }
        lob_order = {
            item.management_lob: item.display_order
            for item in self.services.profiles if item.active_on(latest)
        }
        management_lobs = sorted(
            {row[0] for row in capacity_rows} |
            {item.management_lob for item in self.services.profiles if item.active_on(latest)},
            key=lambda value: (lob_order.get(value, 999), value.casefold()),
        )
        return {
            "product": "WFMHub Manager Workbench",
            "latest_date": latest.isoformat(),
            "database": self.config.database.name,
            "last_refresh": {
                "finished_at": _iso(run[0]) if run else None,
                "status": run[1] if run else "NOT RUN",
                "files_loaded": int(run[2] or 0) if run else 0,
                "files_skipped": int(run[3] or 0) if run else 0,
                "files_failed": int(run[4] or 0) if run else 0,
                "details": run[5] if run else None,
            },
            "sources": [
                {
                    "family": row["source_family"],
                    "newest_date": _iso(row["newest_business_date"]),
                    "rows": int(row["row_count"] or 0),
                    "rejected": int(row["rejected_count"] or 0),
                    "status": row["status"],
                }
                for row in sources
            ],
            "management_lobs": management_lobs,
            "capacity": [
                {"management_lob": row[0], "planning_group": row[1], "staff_type": row[2]}
                for row in sorted(capacity_rows)
            ],
            "agents": agent_rows,
        }

    def _attendance_rows(self, conn: DatabaseConnection, scope: DashboardFilter):
        rows = _rows(conn.execute(
            """SELECT business_date, agent_day_key, agent_id, agent_name,
                      team_leader, ops_manager, lob, language, assignment,
                      scheduled_start, scheduled_end, scheduled_minutes,
                      planned_work_minutes, planning_overlay, first_login,
                      last_logout, actual_first_seen, actual_last_seen,
                      attendance_result, call_action, requires_call,
                      shift_state, actual_evidence, is_provisional,
                      uncoded_late_minutes, uncoded_early_leave_minutes,
                      no_show_minutes, status_covered_minutes, evaluation_as_of
               FROM mart.attendance_agent_day
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date DESC, team_leader, agent_name, agent_id""",
            (scope.start, scope.end),
        ))
        output = []
        for row in rows:
            mapped = self._capacity_for_attendance(row)
            if scope.management_lob and mapped.management_lob != scope.management_lob:
                continue
            if scope.planning_group and mapped.planning_group != scope.planning_group:
                continue
            if scope.staff_type and mapped.staff_type != scope.staff_type:
                continue
            if scope.team_leader and str(row.get("team_leader") or "Unassigned") != scope.team_leader:
                continue
            if scope.agent_id and str(row.get("agent_id") or "") != scope.agent_id:
                continue
            row["management_lob"] = mapped.management_lob
            row["planning_group"] = mapped.planning_group
            row["staff_type"] = mapped.staff_type
            output.append(row)
        return output

    def _service_rows(self, conn: DatabaseConnection, scope: DashboardFilter):
        rows = _rows(conn.execute(
            """SELECT business_date, interval_start, interval_end, source_system,
                      service_scope, comparison_scope, queue, designation,
                      language, offered, answered, abandoned, short_abandoned,
                      abandoned_within_target, answered_within_target,
                      handled_seconds, call_legs, transferred_legs
               FROM mart.call_service_15min
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, comparison_scope, queue""",
            (scope.start, scope.end),
        ))
        output = []
        for row in rows:
            # A reviewed queue can intentionally contribute to more than one
            # management Flash (for example selected Ford NL queues also live
            # in RSA BE). Project through the exact profile allowlists instead
            # of assuming the queue map's comparison scope is the report LOB.
            profiles = [
                profile for profile in self.services.profiles
                if profile.active_on(_day(row["business_date"]))
                and str(row.get("source_system") or "").upper() in profile.source_systems
                and profile.includes_flash_queue(row.get("queue"))
            ]
            for profile in profiles:
                if scope.management_lob and profile.management_lob != scope.management_lob:
                    continue
                projected = dict(row)
                projected["management_lob"] = profile.management_lob
                projected["group"] = profile.group_for(row.get("queue"))
                output.append(projected)
        return output

    @staticmethod
    def _service_bucket(rows: Iterable[dict[str, Any]]) -> dict[str, Any]:
        values = list(rows)
        offered = sum(_number(row.get("offered")) for row in values)
        answered = sum(_number(row.get("answered")) for row in values)
        abandoned = sum(_number(row.get("abandoned")) for row in values)
        abandoned_target = sum(_number(row.get("abandoned_within_target")) for row in values)
        answered_target = sum(_number(row.get("answered_within_target")) for row in values)
        handled_seconds = sum(_number(row.get("handled_seconds")) for row in values)
        return {
            "offered": offered,
            "answered": answered,
            "handled_in_sl": answered_target,
            "abandoned": abandoned,
            "service_level": _ratio(answered_target, offered - abandoned_target),
            "abandon_rate": _ratio(abandoned, offered),
            "aht_seconds": _ratio(handled_seconds, answered),
        }

    def today(self, scope: DashboardFilter) -> dict[str, Any]:
        focus = DashboardFilter(
            scope.end, scope.end, scope.management_lob, scope.planning_group,
            scope.staff_type, scope.team_leader, scope.agent_id,
        )
        conn = self._connect()
        try:
            attendance = self._attendance_rows(conn, focus)
            service = self._service_rows(conn, focus)
            staffing = self._staffing_rows(conn, focus)
        finally:
            conn.close()
        # Profiles may deliberately share queues. A cross-profile total would
        # therefore be a synthetic KPI. Keep each LOB row available, but only
        # calculate the headline and time series for one selected Management LOB.
        scoped_service = service if scope.management_lob else []
        service_summary = self._service_bucket(scoped_service)
        due = [
            row for row in attendance
            if _number(row.get("planned_work_minutes")) > 0
            and str(row.get("shift_state") or "") != "NOT_STARTED"
        ]
        present = [
            row for row in due
            if row.get("attendance_result") in PRESENT_RESULTS
            and (row.get("actual_first_seen") is not None or row.get("first_login") is not None)
        ]
        no_show = [row for row in due if row.get("attendance_result") in NO_SHOW_RESULTS]
        unknown = [row for row in due if row not in present and row not in no_show]
        calls = []
        for row in due:
            needs_evidence_check = row in unknown
            if (
                bool(row.get("requires_call"))
                or _number(row.get("uncoded_late_minutes")) > 0
                or _number(row.get("uncoded_early_leave_minutes")) > 0
                or needs_evidence_check
            ):
                record = self._attendance_record(row)
                if needs_evidence_check and str(record.get("action") or "NONE") == "NONE":
                    record["action"] = "CHECK DATA — POSSIBLE NO SHOW"
                calls.append(record)
        by_lob: dict[str, dict[str, Any]] = {}
        all_lobs = {row["management_lob"] for row in service} | {
            row["management_lob"] for row in attendance
        }
        for lob in sorted(all_lobs):
            lob_service = self._service_bucket(row for row in service if row["management_lob"] == lob)
            lob_att = [row for row in due if row["management_lob"] == lob]
            lob_present = [
                row for row in lob_att
                if row.get("attendance_result") in PRESENT_RESULTS
                and (row.get("actual_first_seen") is not None or row.get("first_login") is not None)
            ]
            lob_no_show = [row for row in lob_att if row.get("attendance_result") in NO_SHOW_RESULTS]
            by_lob[lob] = {
                "management_lob": lob, **lob_service,
                "target": self._service_target(lob, focus.end),
                "due_hc": len(lob_att), "present_hc": len(lob_present),
                "no_show_hc": len(lob_no_show),
                "unknown_hc": max(len(lob_att) - len(lob_present) - len(lob_no_show), 0),
            }
        intervals: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in scoped_service:
            intervals[_iso(row["interval_start"]) or ""].append(row)
        staffing_now = min(
            (sum(_number(row.get("staffing_gap_fte")) for row in values)
             for _, values in self._group(staffing, "interval_start")),
            default=None,
        )
        return {
            "as_of": scope.end.isoformat(),
            "cards": [
                self._card("Service level", service_summary["service_level"], "percent", "Ratio of summed governed counters"),
                self._card("Present HC", len(present), "integer", f"{len(due)} due; late remains present"),
                self._card("No show HC", len(no_show), "integer", "Completed, evidence-backed only"),
                self._card("Possible no show", len(unknown), "integer", "Missing evidence remains unknown"),
            ],
            "service_by_lob": list(by_lob.values()),
            "service_timeline": [
                {"time": key, **self._service_bucket(values)}
                for key, values in sorted(intervals.items())
            ],
            "attendance_actions": calls[:250],
            "capacity_gap_fte": staffing_now,
            "empty": not attendance and not service and not staffing,
            "service_scope_required": not scope.management_lob,
        }

    @staticmethod
    def _card(label: str, value: Any, kind: str, note: str) -> dict[str, Any]:
        return {"label": label, "value": value, "kind": kind, "note": note}

    @staticmethod
    def _group(rows: Iterable[dict[str, Any]], field: str):
        grouped: dict[Any, list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            grouped[row.get(field)].append(row)
        return sorted(grouped.items(), key=lambda item: str(item[0] or ""))

    def _attendance_record(self, row: dict[str, Any]) -> dict[str, Any]:
        return {
            "date": _iso(row.get("business_date")),
            "agent_id": row.get("agent_id"), "agent": row.get("agent_name"),
            "team_leader": row.get("team_leader"),
            "management_lob": row.get("management_lob"),
            "planning_group": row.get("planning_group"), "staff_type": row.get("staff_type"),
            "scheduled_start": _iso(row.get("scheduled_start")),
            "scheduled_end": _iso(row.get("scheduled_end")),
            "first_login": _iso(row.get("first_login")), "last_logout": _iso(row.get("last_logout")),
            "result": row.get("attendance_result"), "action": row.get("call_action"),
            "late_minutes": int(row.get("uncoded_late_minutes") or 0),
            "early_leave_minutes": int(row.get("uncoded_early_leave_minutes") or 0),
            "no_show_minutes": int(row.get("no_show_minutes") or 0),
            "evidence": row.get("actual_evidence"), "provisional": bool(row.get("is_provisional")),
        }

    def _staffing_rows(self, conn: DatabaseConnection, scope: DashboardFilter):
        rows = _rows(conn.execute(
            """SELECT business_date, interval_start, interval_end, lob, language,
                      planning_group, staff_type, capacity_mapping_status,
                      scheduled_agents, observed_agents, productive_agents,
                      auxiliary_agents, scheduled_fte, elapsed_scheduled_fte,
                      observed_fte, productive_fte, staffing_variance_fte,
                      staffing_gap_fte, staffing_state, evidence_basis,
                      evaluation_as_of, gross_scheduled_fte, planned_time_off_fte
               FROM mart.staffing_interval
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, planning_group, staff_type""",
            (scope.start, scope.end),
        ))
        output = []
        for row in rows:
            mapped = self.capacity.map_schedule(row.get("lob"), row.get("staff_type"))
            management_lob = mapped.management_lob
            if scope.management_lob and management_lob != scope.management_lob:
                continue
            if scope.planning_group and row.get("planning_group") != scope.planning_group:
                continue
            if scope.staff_type and row.get("staff_type") != scope.staff_type:
                continue
            row["management_lob"] = management_lob
            output.append(row)
        return output

    def _forecast_rows(self, conn: DatabaseConnection, scope: DashboardFilter):
        rows = _rows(conn.execute(
            """SELECT business_date, interval_start, interval_end,
                      interval_minutes, queue_name, volume_forecast,
                      fte_required, aht_forecast_seconds, source_file
               FROM mart.forecast_interval
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, queue_name, source_file""",
            (scope.start, scope.end),
        ))
        output = []
        for row in rows:
            mapped = self.capacity.map_forecast(row.get("source_file"), row.get("queue_name"))
            if mapped.status != "MAPPED":
                continue
            if scope.management_lob and mapped.management_lob != scope.management_lob:
                continue
            if scope.planning_group and mapped.planning_group != scope.planning_group:
                continue
            if scope.staff_type and mapped.staff_type != scope.staff_type:
                continue
            row["management_lob"] = mapped.management_lob
            row["planning_group"] = mapped.planning_group
            row["staff_type"] = mapped.staff_type
            output.append(row)
        return output

    def staffing(self, scope: DashboardFilter) -> dict[str, Any]:
        conn = self._connect()
        try:
            staffing = self._staffing_rows(conn, scope)
            forecast = self._forecast_rows(conn, scope)
        finally:
            conn.close()
        required_hours = sum(
            _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
            for row in forecast if row.get("fte_required") is not None
        )
        scheduled_hours = sum(_number(row.get("scheduled_fte")) * .25 for row in staffing)
        gross_hours = sum(_number(row.get("gross_scheduled_fte")) * .25 for row in staffing)
        time_off_hours = sum(_number(row.get("planned_time_off_fte")) * .25 for row in staffing)

        required: dict[tuple[str, str, str, str], float] = defaultdict(float)
        for row in forecast:
            key = (
                _iso(row["business_date"]) or "", _iso(row["interval_start"]) or "",
                row["planning_group"], row["staff_type"],
            )
            if row.get("fte_required") is not None:
                required[key] += _number(row["fte_required"])
        scheduled: dict[tuple[str, str, str, str], dict[str, float]] = defaultdict(lambda: defaultdict(float))
        evidence: dict[tuple[str, str, str, str], str] = {}
        for row in staffing:
            key = (
                _iso(row["business_date"]) or "", _iso(row["interval_start"]) or "",
                str(row["planning_group"]), str(row["staff_type"]),
            )
            for field in ("scheduled_fte", "gross_scheduled_fte", "planned_time_off_fte", "observed_fte", "productive_fte"):
                scheduled[key][field] += _number(row.get(field))
            evidence[key] = str(row.get("evidence_basis") or "NONE")
        keys = sorted(set(required) | set(scheduled))
        intervals = []
        actions = []
        for key in keys:
            needed = required.get(key)
            values = scheduled.get(key, {})
            net = values.get("scheduled_fte", 0)
            gap = net - needed if key in required else None
            record = {
                "date": key[0], "time": key[1], "planning_group": key[2],
                "staff_type": key[3], "required_fte": needed,
                "gross_fte": values.get("gross_scheduled_fte", 0),
                "time_off_fte": values.get("planned_time_off_fte", 0),
                "scheduled_fte": net, "observed_fte": values.get("observed_fte", 0),
                "productive_fte": values.get("productive_fte", 0),
                "gap_fte": gap, "evidence": evidence.get(key, "NO_SCHEDULE"),
            }
            intervals.append(record)
            if gap is not None and gap < -0.001:
                actions.append(record)
        return {
            "cards": [
                self._card("Required FTE hours", required_hours, "decimal", "Explicit Verint absolute requirement"),
                self._card("Scheduled coverage", _ratio(scheduled_hours, required_hours), "percent", "Net published FTE hours / required FTE hours"),
                self._card("PTO / Away FTE hours", time_off_hours, "decimal", f"Gross {gross_hours:.1f} FTE hours"),
                self._card("Shortage intervals", len(actions), "integer", "Net schedule below explicit requirement"),
            ],
            "intervals": intervals, "actions": sorted(actions, key=lambda row: row["gap_fte"] or 0)[:500],
            "empty": not staffing and not forecast,
        }

    def demand(self, scope: DashboardFilter) -> dict[str, Any]:
        """Forecast Volume and absolute requirement at governed Staff Type grain."""

        conn = self._connect()
        try:
            rows = self._forecast_rows(conn, scope)
        finally:
            conn.close()
        total_volume = sum(
            _number(row.get("volume_forecast")) for row in rows
            if row.get("volume_forecast") is not None
        )
        required_hours = sum(
            _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
            for row in rows if row.get("fte_required") is not None
        )
        interval_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            interval_groups[_iso(row.get("interval_start")) or ""].append(row)
        profile = []
        for key, values in sorted(interval_groups.items()):
            requirements = [
                _number(row.get("fte_required")) for row in values
                if row.get("fte_required") is not None
            ]
            profile.append({
                "time": key,
                "volume": sum(_number(row.get("volume_forecast")) for row in values),
                "required_fte": sum(requirements) if requirements else None,
            })
        peak = max(
            (row for row in profile if row["required_fte"] is not None),
            key=lambda row: row["required_fte"], default=None,
        )
        source_groups: dict[tuple[str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
        ledger_groups: dict[tuple[str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            source_groups[(
                str(row.get("source_file") or ""), row["management_lob"],
                row["planning_group"], row["staff_type"],
            )].append(row)
            ledger_groups[(
                _iso(row.get("business_date")) or "", row["management_lob"],
                row["planning_group"], row["staff_type"],
            )].append(row)
        source_register = []
        for key, values in sorted(source_groups.items()):
            source_register.append({
                "source_file": key[0], "management_lob": key[1],
                "planning_group": key[2], "staff_type": key[3],
                "latest_date": max(_iso(row.get("business_date")) or "" for row in values),
                "intervals": len(values),
                "volume": sum(_number(row.get("volume_forecast")) for row in values),
            })
        ledger = []
        for key, values in sorted(ledger_groups.items()):
            by_interval: dict[str, float] = defaultdict(float)
            requirement_present = False
            for row in values:
                if row.get("fte_required") is not None:
                    requirement_present = True
                    by_interval[_iso(row.get("interval_start")) or ""] += _number(row["fte_required"])
            peak_item = max(by_interval.items(), key=lambda item: item[1], default=(None, None))
            ledger.append({
                "date": key[0], "management_lob": key[1],
                "planning_group": key[2], "staff_type": key[3],
                "volume": sum(_number(row.get("volume_forecast")) for row in values),
                "required_fte_hours": sum(
                    _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
                    for row in values if row.get("fte_required") is not None
                ) if requirement_present else None,
                "peak_fte": peak_item[1], "peak_interval": peak_item[0],
            })
        latest = max((_day(row["business_date"]) for row in rows), default=None)
        return {
            "cards": [
                self._card("Forecast volume", total_volume, "integer", "Selected Staff Type demand"),
                self._card("Required FTE hours", required_hours, "decimal", "Absolute requirement × interval duration"),
                self._card("Peak requirement", peak["required_fte"] if peak else None, "decimal", _iso(peak["time"]) if peak else "No requirement"),
                self._card("Forecast through", _iso(latest), "date", "Latest mapped forecast interval"),
            ],
            "profile": profile, "sources": source_register,
            "ledger": ledger[:1000], "empty": not rows,
        }

    def scenario(self, scope: DashboardFilter) -> dict[str, Any]:
        """Non-persistent capacity what-if using one explicit FTE adjustment."""

        staffing = self.staffing(scope)
        adjustment = scope.scenario_fte
        intervals = []
        baseline_uncovered = 0.0
        adjusted_uncovered = 0.0
        recovered = 0.0
        for row in staffing["intervals"]:
            required = row.get("required_fte")
            baseline = row.get("scheduled_fte")
            if required is None:
                adjusted_gap = None
                baseline_gap = None
            else:
                baseline_gap = _number(baseline) - _number(required)
                adjusted_gap = _number(baseline) + adjustment - _number(required)
                baseline_uncovered += max(-baseline_gap, 0) * .25
                adjusted_uncovered += max(-adjusted_gap, 0) * .25
            record = dict(row)
            record.update({
                "baseline_gap_fte": baseline_gap,
                "adjusted_fte": _number(baseline) + adjustment,
                "adjusted_gap_fte": adjusted_gap,
            })
            intervals.append(record)
        recovered = max(baseline_uncovered - adjusted_uncovered, 0)
        peak_gap = min(
            (row["adjusted_gap_fte"] for row in intervals if row["adjusted_gap_fte"] is not None),
            default=None,
        )
        return {
            "cards": [
                self._card("Baseline uncovered", baseline_uncovered, "decimal", "FTE hours before scenario"),
                self._card("Recovered", recovered, "decimal", "FTE hours recovered by explicit adjustment"),
                self._card("Applied adjustment", adjustment, "decimal", "FTE across selected scope; not persisted"),
                self._card("Remaining exposure", adjusted_uncovered, "decimal", "Adjusted uncovered FTE hours"),
            ],
            "adjustment_fte": adjustment,
            "peak_gap_fte": peak_gap,
            "intervals": intervals,
            "actions": sorted(
                [row for row in intervals if row.get("adjusted_gap_fte") is not None and row["adjusted_gap_fte"] < 0],
                key=lambda row: row["adjusted_gap_fte"],
            )[:500],
            "empty": staffing["empty"],
        }

    def service(self, scope: DashboardFilter) -> dict[str, Any]:
        conn = self._connect()
        try:
            rows = self._service_rows(conn, scope)
        finally:
            conn.close()
        # Do not invent a portfolio-wide SL across intentionally overlapping
        # service profiles. The browser always starts on one governed LOB.
        scoped_rows = rows if scope.management_lob else []
        summary = self._service_bucket(scoped_rows)
        timeline = []
        for key, values in self._group(scoped_rows, "interval_start"):
            timeline.append({"time": _iso(key), **self._service_bucket(values)})
        by_lob = []
        for key, values in self._group(rows, "management_lob"):
            by_lob.append({
                "management_lob": key,
                "target": self._service_target(str(key), scope.end),
                **self._service_bucket(values),
            })
        by_queue = []
        for key, values in self._group(rows, "queue"):
            first = values[0]
            by_queue.append({
                "queue": key, "management_lob": first["management_lob"],
                "group": first["group"], **self._service_bucket(values),
            })
        by_queue.sort(key=lambda row: (row["service_level"] is None, row["service_level"] or 0, -row["offered"]))
        return {
            "cards": [
                self._card("Service level", summary["service_level"], "percent", "Answered in target / offered less abandons in target"),
                self._card("Offered volume", summary["offered"], "integer", "Exact configured Flash queues"),
                self._card("Handled volume", summary["answered"], "integer", f"{summary['handled_in_sl']:.0f} handled in SL"),
                self._card("Abandon rate", summary["abandon_rate"], "percent", f"AHT {summary['aht_seconds'] or 0:.0f}s"),
            ],
            "timeline": timeline, "by_lob": by_lob, "queues": by_queue[:500],
            "target": self._service_target(scope.management_lob, scope.end)
            if scope.management_lob else None,
            "empty": not rows,
            "service_scope_required": not scope.management_lob,
        }

    def attendance(self, scope: DashboardFilter) -> dict[str, Any]:
        conn = self._connect()
        try:
            attendance = self._attendance_rows(conn, scope)
            integrity = _rows(conn.execute(
                """SELECT * FROM mart.schedule_integrity_agent_day
                   WHERE business_date BETWEEN ? AND ? AND requires_review=1
                   ORDER BY business_date DESC, is_recurring DESC,
                            displaced_minutes DESC, agent_name""",
                (scope.start, scope.end),
            ))
            gaps = _rows(conn.execute(
                """SELECT * FROM mart.correction_candidate
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date DESC, priority, agent_name, gap_start""",
                (scope.start, scope.end),
            ))
            timeline = _rows(conn.execute(
                """SELECT * FROM mart.shift_timeline_segment
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date DESC, agent_name, segment_start""",
                (scope.start, scope.end),
            ))
            controls = _rows(conn.execute(
                """SELECT a.business_date, a.agent_id, a.agent_name,
                          a.team_leader, a.lob, a.scheduled_start, a.scheduled_end,
                          c.status_coverage_percent, c.break_minutes,
                          c.break_overrun_minutes, c.lunch_minutes,
                          c.lunch_overrun_minutes, c.measurement_basis
                   FROM mart.conformance_agent_day c
                   JOIN mart.attendance_agent_day a
                     ON a.agent_day_key=c.agent_day_key
                   WHERE a.business_date BETWEEN ? AND ?
                     AND a.shift_state='COMPLETE'
                   ORDER BY a.business_date DESC,
                            (c.break_overrun_minutes+c.lunch_overrun_minutes) DESC,
                            a.agent_name""",
                (scope.start, scope.end),
            ))
        finally:
            conn.close()
        allowed = {str(row["agent_day_key"]) for row in attendance}
        agent_ids = {str(row["agent_id"]) for row in attendance}
        integrity = [row for row in integrity if str(row.get("agent_day_key")) in allowed]
        gaps = [row for row in gaps if str(row.get("agent_id")) in agent_ids]
        timeline = [row for row in timeline if str(row.get("agent_day_key")) in allowed]
        controls = [row for row in controls if str(row.get("agent_id")) in agent_ids]
        timeline_by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in timeline:
            timeline_by_day[str(row["agent_day_key"])].append({
                "start": _iso(row.get("segment_start")), "end": _iso(row.get("segment_end")),
                "minutes": int(row.get("segment_minutes") or 0),
                "actual_status": row.get("actual_status"), "category": row.get("actual_category"),
                "mismatch": row.get("mismatch_type"), "is_gap": bool(row.get("is_gap")),
                "source": row.get("observed_source"),
            })
        review = []
        for row in integrity[:120]:
            review.append({
                "date": _iso(row.get("business_date")), "agent_day_key": row.get("agent_day_key"),
                "agent_id": row.get("agent_id"), "agent": row.get("agent_name"),
                "team_leader": row.get("team_leader"),
                "management_lob": self._workforce_lob(row.get("lob"), _day(row["business_date"])),
                "scheduled_start": _iso(row.get("scheduled_start")), "scheduled_end": _iso(row.get("scheduled_end")),
                "observed_start": _iso(row.get("observed_start")), "observed_end": _iso(row.get("observed_end")),
                "classification": row.get("classification"),
                "displaced_minutes": int(row.get("displaced_minutes") or 0),
                "internal_gap_minutes": int(row.get("internal_gap_minutes") or 0),
                "recurrence_count": int(row.get("recurrence_count") or 0),
                "segments": timeline_by_day.get(str(row["agent_day_key"]), []),
            })
        gap_records = [{
            "date": _iso(row.get("business_date")), "gap_id": row.get("correction_id"),
            "agent_id": row.get("agent_id"), "agent": row.get("agent_name"),
            "team_leader": row.get("team_leader"),
            "management_lob": self._workforce_lob(row.get("lob"), _day(row["business_date"])),
            "start": _iso(row.get("gap_start")), "end": _iso(row.get("gap_end")),
            "minutes": int(row.get("gap_minutes") or 0), "issue": row.get("detected_issue"),
            "suggested_activity": row.get("suggested_activity"),
            "reconciliation": row.get("verint_reconciliation"), "source": row.get("observed_source"),
        } for row in gaps]
        alert_records = [{
            "date": _iso(row.get("business_date")), "agent_id": row.get("agent_id"),
            "agent": row.get("agent_name"), "team_leader": row.get("team_leader"),
            "management_lob": self._workforce_lob(row.get("lob"), _day(row["business_date"])),
            "break_minutes": int(row.get("break_minutes") or 0),
            "break_overrun": int(row.get("break_overrun_minutes") or 0),
            "meal_minutes": int(row.get("lunch_minutes") or 0),
            "meal_overrun": int(row.get("lunch_overrun_minutes") or 0),
            "coverage": row.get("status_coverage_percent"), "evidence": row.get("measurement_basis"),
        } for row in controls if _number(row.get("break_overrun_minutes")) > 0 or _number(row.get("lunch_overrun_minutes")) > 0]
        return {
            "cards": [
                self._card("Review agent-days", len(review), "integer", "Published schedule versus observed boundaries"),
                self._card("Residual gaps", len(gap_records), "integer", "Only intervals still unsupported by final Activities"),
                self._card("Residual gap hours", sum(row["minutes"] for row in gap_records) / 60, "decimal", "Exact start/end fragments"),
                self._card("Break / meal alerts", len(alert_records), "integer", "Completed shifts with evidence"),
            ],
            "review": review, "gaps": gap_records[:1000], "break_meal": alert_records[:500],
            "empty": not attendance and not integrity and not gaps,
        }

    def history(self, scope: DashboardFilter) -> dict[str, Any]:
        conn = self._connect()
        try:
            service = self._service_rows(conn, scope)
            staffing = self._staffing_rows(conn, scope)
            forecast = self._forecast_rows(conn, scope)
            absence = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, team_leader, lob,
                          planned_net_minutes, final_absence_minutes,
                          final_vacation_minutes, final_unpaid_minutes,
                          final_shrinkage_minutes, final_unmapped_minutes,
                          final_ledger_status
                   FROM mart.verint_final_absence_agent_day
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date, lob, team_leader, agent_name""",
                (scope.start, scope.end),
            ))
            patterns = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, team_leader, lob,
                          classification, pattern_family, recurrence_count,
                          eligible_day_count, displaced_minutes, confidence
                   FROM mart.schedule_integrity_agent_day
                   WHERE business_date BETWEEN ? AND ? AND is_recurring=1
                   ORDER BY recurrence_count DESC, displaced_minutes DESC,
                            business_date DESC, agent_name""",
                (scope.start, scope.end),
            ))
        finally:
            conn.close()
        valid_agents = None
        if scope.team_leader or scope.agent_id or scope.management_lob:
            valid_agents = {
                str(row["agent_id"])
                for row in self._attendance_filter_from_dimension(scope)
            }
        if valid_agents is not None:
            absence = [row for row in absence if str(row["agent_id"]) in valid_agents]
            patterns = [row for row in patterns if str(row["agent_id"]) in valid_agents]
        summary = self._service_bucket(service)
        required_hours = sum(
            _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
            for row in forecast if row.get("fte_required") is not None
        )
        scheduled_hours = sum(_number(row.get("scheduled_fte")) * .25 for row in staffing)
        planned = sum(_number(row.get("planned_net_minutes")) for row in absence)
        absent = sum(_number(row.get("final_absence_minutes")) for row in absence)
        shrink = sum(_number(row.get("final_shrinkage_minutes")) for row in absence)
        by_day: dict[str, dict[str, Any]] = defaultdict(lambda: {
            "offered": 0.0, "answered": 0.0, "answered_target": 0.0,
            "abandoned_target": 0.0, "required_hours": 0.0,
            "scheduled_hours": 0.0, "observed_hours": 0.0,
            "productive_hours": 0.0, "planned_minutes": 0.0,
            "absence_minutes": 0.0, "shrinkage_minutes": 0.0,
        })
        for row in service:
            key = _iso(row["business_date"]) or ""
            by_day[key]["offered"] += _number(row.get("offered"))
            by_day[key]["answered"] += _number(row.get("answered"))
            by_day[key]["answered_target"] += _number(row.get("answered_within_target"))
            by_day[key]["abandoned_target"] += _number(row.get("abandoned_within_target"))
        for row in forecast:
            key = _iso(row["business_date"]) or ""
            by_day[key]["required_hours"] += _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
        for row in staffing:
            key = _iso(row["business_date"]) or ""
            by_day[key]["scheduled_hours"] += _number(row.get("scheduled_fte")) * .25
            by_day[key]["observed_hours"] += _number(row.get("observed_fte")) * .25
            by_day[key]["productive_hours"] += _number(row.get("productive_fte")) * .25
        for row in absence:
            key = _iso(row["business_date"]) or ""
            by_day[key]["planned_minutes"] += _number(row.get("planned_net_minutes"))
            by_day[key]["absence_minutes"] += _number(row.get("final_absence_minutes"))
            by_day[key]["shrinkage_minutes"] += _number(row.get("final_shrinkage_minutes"))
        daily = []
        for key, values in sorted(by_day.items()):
            daily.append({
                "date": key, "offered": values["offered"],
                "service_level": _ratio(values["answered_target"], values["offered"] - values["abandoned_target"]),
                "required_fte_hours": values["required_hours"],
                "scheduled_fte_hours": values["scheduled_hours"],
                "observed_fte_hours": values["observed_hours"],
                "productive_fte_hours": values["productive_hours"],
                "absence_rate": _ratio(values["absence_minutes"], values["planned_minutes"]),
                "shrinkage_rate": _ratio(values["shrinkage_minutes"], values["planned_minutes"]),
            })
        pattern_records = [{
            "date": _iso(row["business_date"]), "agent_id": row["agent_id"],
            "agent": row["agent_name"], "team_leader": row["team_leader"],
            "management_lob": self._workforce_lob(row["lob"], _day(row["business_date"])),
            "classification": row["classification"], "pattern": row["pattern_family"],
            "recurrence": int(row["recurrence_count"] or 0),
            "eligible_days": int(row["eligible_day_count"] or 0),
            "displaced_minutes": int(row["displaced_minutes"] or 0),
            "confidence": row["confidence"],
        } for row in patterns]
        return {
            "cards": [
                self._card("Service level", summary["service_level"], "percent", "Ratio of summed service counters"),
                self._card("Required FTE hours", required_hours, "decimal", "Explicit Verint requirement"),
                self._card("Scheduled FTE hours", scheduled_hours, "decimal", "Net published capacity"),
                self._card("Final absence / shrinkage", _ratio(absent, planned), "percent", f"Shrinkage {_ratio(shrink, planned) or 0:.1%}"),
            ],
            "daily": daily, "patterns": pattern_records[:500],
            "final_absence": {
                "planned_minutes": planned, "absence_minutes": absent,
                "shrinkage_minutes": shrink, "absence_rate": _ratio(absent, planned),
                "shrinkage_rate": _ratio(shrink, planned),
            },
            "empty": not daily,
        }

    def manager_desk(self, scope: DashboardFilter) -> dict[str, Any]:
        """Prioritized current-day, post-day and forward WFM evidence."""

        today = self.today(scope)
        forward_scope = DashboardFilter(
            scope.end, scope.end + timedelta(days=6), scope.management_lob,
            scope.planning_group, scope.staff_type, scope.team_leader,
            scope.agent_id,
        )
        forward = self.staffing(forward_scope)
        review_scope = DashboardFilter(
            scope.end - timedelta(days=6), scope.end, scope.management_lob,
            scope.planning_group, scope.staff_type, scope.team_leader,
            scope.agent_id,
        )
        review = self.attendance(review_scope)
        service_rows = [
            row for row in today["service_by_lob"]
            if row.get("service_level") is not None
        ]
        selected_service = None
        if scope.management_lob:
            selected_service = next(
                (row for row in service_rows if row["management_lob"] == scope.management_lob),
                None,
            )
        service_focus = selected_service or min(
            service_rows, key=lambda row: row["service_level"], default=None,
        )
        uncovered = sum(
            max(-_number(row.get("gap_fte")), 0) * .25
            for row in forward["actions"]
        )
        work_queue = []
        for row in today["attendance_actions"]:
            work_queue.append({
                "priority": "P1" if "CALL" in str(row.get("action") or "") else "CHECK",
                "deadline": "NOW",
                "type": "PEOPLE",
                "title": f"{row.get('agent') or row.get('agent_id')} · {row.get('result') or 'Unknown'}",
                "detail": f"{row.get('management_lob') or 'Unmapped'} · {row.get('evidence') or 'Evidence missing'}",
                "source": "Attendance Pulse",
            })
        for row in forward["actions"][:8]:
            work_queue.append({
                "priority": "P1" if _number(row.get("gap_fte")) <= -2 else "P2",
                "deadline": str(row.get("time") or "")[-8:-3],
                "type": "CAPACITY",
                "title": f"{row.get('planning_group')} shortage {row.get('gap_fte'):.1f} FTE",
                "detail": f"{row.get('date')} · {row.get('staff_type')}",
                "source": "Staff Preparation",
            })
        label = "Selected LOB service" if selected_service else "Lowest LOB service"
        service_note = (
            str(service_focus.get("management_lob")) if service_focus
            else "No service evidence"
        )
        return {
            "cards": [
                self._card(label, service_focus.get("service_level") if service_focus else None, "percent", service_note),
                self._card("Calls requiring action", len(today["attendance_actions"]), "integer", "Confirmed calls and evidence checks"),
                self._card("Residual gaps", len(review["gaps"]), "integer", "Exact unresolved intervals"),
                self._card("Forward uncovered", uncovered, "decimal", "FTE hours across next seven days"),
            ],
            "lob_matrix": today["service_by_lob"],
            "work_queue": sorted(work_queue, key=lambda row: (row["priority"] != "P1", row["deadline"]))[:12],
            "forward_risk": forward["actions"][:100],
            "horizons": {
                "today": {
                    "call_now": len(today["attendance_actions"]),
                    "no_show": today["cards"][2]["value"],
                    "unknown": today["cards"][3]["value"],
                },
                "post_day": {
                    "residual_gaps": len(review["gaps"]),
                    "gap_hours": sum(_number(row.get("minutes")) for row in review["gaps"]) / 60,
                    "break_meal": len(review["break_meal"]),
                },
                "forward": {
                    "shortage_intervals": len(forward["actions"]),
                    "uncovered_fte_hours": uncovered,
                    "pto_away_fte_hours": forward["cards"][2]["value"],
                },
            },
            "empty": today["empty"] and forward["empty"] and review["empty"],
        }

    def realisations(self, scope: DashboardFilter) -> dict[str, Any]:
        """Plan-to-delivery reconciliation without blending service contracts."""

        conn = self._connect()
        try:
            service = self._service_rows(conn, scope)
            staffing = self._staffing_rows(conn, scope)
            forecast = self._forecast_rows(conn, scope)
            absence = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, team_leader, lob,
                          planned_net_minutes, final_absence_minutes,
                          final_shrinkage_minutes, final_ledger_status
                   FROM mart.verint_final_absence_agent_day
                   WHERE business_date BETWEEN ? AND ?""",
                (scope.start, scope.end),
            ))
        finally:
            conn.close()
        if any((scope.management_lob, scope.planning_group, scope.staff_type, scope.team_leader, scope.agent_id)):
            valid_agents = {
                str(row["agent_id"])
                for row in self._attendance_filter_from_dimension(scope)
            }
            absence = [row for row in absence if str(row.get("agent_id")) in valid_agents]

        by_day: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
        by_lob: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
        for row in service:
            for bucket in (by_day[_iso(row["business_date"]) or ""], by_lob[row["management_lob"]]):
                bucket["offered"] += _number(row.get("offered"))
                bucket["answered_target"] += _number(row.get("answered_within_target"))
                bucket["abandoned_target"] += _number(row.get("abandoned_within_target"))
        for row in forecast:
            hours = _number(row.get("fte_required")) * _number(row.get("interval_minutes")) / 60
            for bucket in (by_day[_iso(row["business_date"]) or ""], by_lob[row["management_lob"]]):
                bucket["forecast"] += _number(row.get("volume_forecast"))
                bucket["required_hours"] += hours
        for row in staffing:
            for bucket in (by_day[_iso(row["business_date"]) or ""], by_lob[row["management_lob"]]):
                bucket["scheduled_hours"] += _number(row.get("scheduled_fte")) * .25
                bucket["observed_hours"] += _number(row.get("observed_fte")) * .25
                bucket["productive_hours"] += _number(row.get("productive_fte")) * .25
        for row in absence:
            management_lob = self._workforce_lob(row.get("lob"), _day(row["business_date"]))
            for bucket in (by_day[_iso(row["business_date"]) or ""], by_lob[management_lob]):
                bucket["planned_minutes"] += _number(row.get("planned_net_minutes"))
                bucket["absence_minutes"] += _number(row.get("final_absence_minutes"))
                bucket["shrinkage_minutes"] += _number(row.get("final_shrinkage_minutes"))

        def complete(key: str, values: dict[str, float]) -> dict[str, Any]:
            return {
                "key": key, "actual_volume": values["offered"],
                "forecast_volume": values["forecast"],
                "volume_variance": _ratio(values["offered"] - values["forecast"], values["forecast"]),
                "service_level": _ratio(values["answered_target"], values["offered"] - values["abandoned_target"]),
                "target": self._service_target(key, scope.end) if key in {item.management_lob for item in self.services.profiles} else None,
                "required_fte_hours": values["required_hours"],
                "scheduled_fte_hours": values["scheduled_hours"],
                "observed_fte_hours": values["observed_hours"],
                "productive_fte_hours": values["productive_hours"],
                "scheduled_coverage": _ratio(values["scheduled_hours"], values["required_hours"]),
                "absence_rate": _ratio(values["absence_minutes"], values["planned_minutes"]),
                "shrinkage_rate": _ratio(values["shrinkage_minutes"], values["planned_minutes"]),
            }

        daily = [{"date": key, **complete(key, values)} for key, values in sorted(by_day.items())]
        lob_rows = [{"management_lob": key, **complete(key, values)} for key, values in sorted(by_lob.items())]
        total_forecast = sum(row["forecast_volume"] for row in daily)
        total_actual = sum(row["actual_volume"] for row in daily)
        total_required = sum(row["required_fte_hours"] for row in daily)
        total_scheduled = sum(row["scheduled_fte_hours"] for row in daily)
        total_observed = sum(row["observed_fte_hours"] for row in daily)
        total_productive = sum(row["productive_fte_hours"] for row in daily)
        lowest = min(
            (row for row in lob_rows if row["service_level"] is not None),
            key=lambda row: row["service_level"], default=None,
        )
        return {
            "cards": [
                self._card("Lowest LOB service", lowest["service_level"] if lowest else None, "percent", lowest["management_lob"] if lowest else "No service evidence"),
                self._card("Actual vs forecast", _ratio(total_actual - total_forecast, total_forecast), "percent", "Volume variance; not a performance score"),
                self._card("Scheduled coverage", _ratio(total_scheduled, total_required), "percent", "Net scheduled / required FTE hours"),
                self._card("Observed productive", _ratio(total_productive, total_scheduled), "percent", f"Observed {(_ratio(total_observed, total_scheduled) or 0):.1%}"),
            ],
            "daily": daily, "by_lob": lob_rows,
            "capacity_chain": {
                "required": total_required, "scheduled": total_scheduled,
                "observed": total_observed, "productive": total_productive,
            },
            "empty": not daily,
        }

    def absence(self, scope: DashboardFilter) -> dict[str, Any]:
        """Final Verint Activities components and completeness exceptions."""

        conn = self._connect()
        try:
            days = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, team_leader, lob,
                          planned_net_minutes, final_absence_minutes,
                          final_vacation_minutes, final_unpaid_minutes,
                          final_shrinkage_minutes, final_unmapped_minutes,
                          final_ledger_status
                   FROM mart.verint_final_absence_agent_day
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date DESC, team_leader, agent_name""",
                (scope.start, scope.end),
            ))
            events = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, team_leader, lob,
                          activity, category, event_start, event_end, minutes,
                          counts_as_absence, counts_as_vacation, counts_as_unpaid,
                          counts_as_shrinkage, mapped, evidence_type, source_file
                   FROM mart.verint_final_absence_event
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date DESC, team_leader, agent_name, event_start""",
                (scope.start, scope.end),
            ))
            exceptions = _rows(conn.execute(
                """SELECT business_date, agent_id, agent_name, activity, category,
                          event_start, event_end, minutes, exception_type, source_file
                   FROM mart.verint_final_exception
                   WHERE business_date BETWEEN ? AND ?
                   ORDER BY business_date DESC, exception_type, agent_name""",
                (scope.start, scope.end),
            ))
        finally:
            conn.close()
        if any((scope.management_lob, scope.planning_group, scope.staff_type, scope.team_leader, scope.agent_id)):
            valid_agents = {
                str(row["agent_id"])
                for row in self._attendance_filter_from_dimension(scope)
            }
            days = [row for row in days if str(row.get("agent_id")) in valid_agents]
            events = [row for row in events if str(row.get("agent_id")) in valid_agents]
            exceptions = [row for row in exceptions if str(row.get("agent_id")) in valid_agents]
        planned = sum(_number(row.get("planned_net_minutes")) for row in days)
        absent = sum(_number(row.get("final_absence_minutes")) for row in days)
        shrinkage = sum(_number(row.get("final_shrinkage_minutes")) for row in days)
        final_states = {"CLEAR", "ABSENCE_RECORDED"}
        ready = [row for row in days if str(row.get("final_ledger_status")) in final_states]
        review = [row for row in days if str(row.get("final_ledger_status")) not in final_states]
        components: dict[str, dict[str, Any]] = {}
        for row in events:
            category = str(row.get("category") or "UNMAPPED")
            item = components.setdefault(category, {
                "category": category, "minutes": 0, "events": 0,
                "counts_as_absence": False, "counts_as_vacation": False,
                "counts_as_unpaid": False, "counts_as_shrinkage": False,
            })
            item["minutes"] += int(row.get("minutes") or 0)
            item["events"] += 1
            for flag in ("counts_as_absence", "counts_as_vacation", "counts_as_unpaid", "counts_as_shrinkage"):
                item[flag] = item[flag] or bool(row.get(flag))
        return {
            "cards": [
                self._card("Final absence", _ratio(absent, planned), "percent", "Verint Activities final components"),
                self._card("Final shrinkage", _ratio(shrinkage, planned), "percent", "Reported separately from absence"),
                self._card("Final-ready rows", len(ready), "integer", "CLEAR or ABSENCE_RECORDED"),
                self._card("Rows for review", len(review) + len(exceptions), "integer", "Incomplete or exceptional evidence"),
            ],
            "components": sorted(components.values(), key=lambda row: -row["minutes"]),
            "days": [
                {
                    **row,
                    "date": _iso(row.get("business_date")),
                    "management_lob": self._workforce_lob(row.get("lob"), _day(row["business_date"])),
                }
                for row in days
            ][:1000],
            "events": [{**row, "date": _iso(row.get("business_date")), "start": _iso(row.get("event_start")), "end": _iso(row.get("event_end"))} for row in events[:1000]],
            "exceptions": [{**row, "date": _iso(row.get("business_date")), "start": _iso(row.get("event_start")), "end": _iso(row.get("event_end"))} for row in exceptions[:1000]],
            "completeness": {
                "total": len(days), "ready": len(ready), "review": len(review),
                "rate": _ratio(len(ready), len(days)),
            },
            "empty": not days and not events and not exceptions,
        }

    def patterns(self, scope: DashboardFilter) -> dict[str, Any]:
        """Supported recurrence evidence, never an inference about intent."""

        history = self.history(scope)
        rows = history["patterns"]
        families: dict[str, set[str]] = defaultdict(set)
        agents: set[str] = set()
        for row in rows:
            agent_id = str(row.get("agent_id") or "")
            agents.add(agent_id)
            families[str(row.get("pattern") or "Other")].add(agent_id)
        ranked = sorted(families.items(), key=lambda item: (-len(item[1]), item[0]))
        cards = [
            self._card(label, len(agent_ids), "integer", "Agents meeting configured recurrence evidence")
            for label, agent_ids in ranked[:4]
        ]
        while len(cards) < 4:
            cards.append(self._card("No additional pattern", None, "integer", "No governed evidence"))
        return {
            "cards": cards, "patterns": rows,
            "families": [{"pattern": key, "agents": len(value)} for key, value in ranked],
            "recurring_agents": len(agents), "empty": not rows,
        }

    def readiness(self, scope: DashboardFilter) -> dict[str, Any]:
        """Source health and current quality backlog."""

        meta = self.meta()
        conn = self._connect()
        try:
            issues = _rows(conn.execute(
                """SELECT detected_at, source_family, source_file, business_date,
                          agent_id, issue_type, severity, details
                   FROM meta.quality_issue
                   WHERE business_date IS NULL OR business_date BETWEEN ? AND ?
                   ORDER BY detected_at DESC, severity, issue_type LIMIT 500""",
                (scope.start, scope.end),
            ))
        finally:
            conn.close()
        ready = sum(1 for row in meta["sources"] if str(row.get("status") or "").upper() in {"READY", "OK", "CURRENT"})
        blocking = sum(1 for row in issues if str(row.get("severity") or "").upper() in {"ERROR", "CRITICAL"})
        return {
            "cards": [
                self._card("Required feeds ready", ready, "integer", f"{len(meta['sources'])} configured source families"),
                self._card("Latest actual date", meta["latest_date"], "date", "Calls / Agent Status evidence"),
                self._card("Rejected source rows", sum(int(row.get("rejected") or 0) for row in meta["sources"]), "integer", "Visible data-quality boundary"),
                self._card("Blocking issues", blocking, "integer", "ERROR or CRITICAL quality findings"),
            ],
            "sources": meta["sources"],
            "issues": [{**row, "detected_at": _iso(row.get("detected_at")), "business_date": _iso(row.get("business_date"))} for row in issues],
            "last_refresh": meta["last_refresh"], "empty": not meta["sources"],
        }

    def mappings(self, scope: DashboardFilter) -> dict[str, Any]:
        """Effective read-only service and capacity configuration."""

        profiles = []
        for profile in sorted(self.services.profiles, key=lambda item: item.display_order):
            if not profile.active_on(scope.end):
                continue
            profiles.append({
                "id": profile.profile_id, "label": profile.label,
                "management_lob": profile.management_lob,
                "service_scopes": list(profile.service_scopes),
                "staffing_lobs": list(profile.staffing_lobs),
                "queue_count": len(profile.flash_queues),
                "queues": list(profile.flash_queues),
                "target": self._service_target(profile.management_lob, scope.end),
                "effective_from": _iso(profile.effective_from),
                "effective_to": _iso(profile.effective_to),
            })
        capacity_rows = []
        with self.config.capacity_mapping.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle):
                capacity_rows.append(dict(row))
        return {
            "cards": [
                self._card("Service profiles", len(profiles), "integer", f"Catalog {self.services.version}"),
                self._card("Exact service queues", sum(row["queue_count"] for row in profiles), "integer", "Profile memberships; overlaps allowed"),
                self._card("Capacity mappings", len(capacity_rows), "integer", "Staff Type and schedule assignments"),
                self._card("Metric methods", len(self.metrics.methods), "integer", f"Catalog {self.metrics.version}"),
            ],
            "service_profiles": profiles, "capacity_mappings": capacity_rows,
            "hashes": {
                "service_profiles": self.services.sha256,
                "capacity_mapping": self.capacity.sha256,
                "metric_catalog": self.metrics.sha256,
            },
            "empty": not profiles and not capacity_rows,
        }

    def _attendance_filter_from_dimension(self, scope: DashboardFilter) -> list[dict[str, Any]]:
        conn = self._connect()
        try:
            return self._attendance_rows(conn, scope)
        finally:
            conn.close()
