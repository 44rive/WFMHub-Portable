"""Read-only projections for the local WFM operations console.

The browser never receives raw extracts and never calculates business KPIs.
Every projection below reads governed marts and effective configuration only.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Iterable

from .capacity_mapping import CapacityMapping, load_capacity_mapping
from .config import Config
from .database import DatabaseConnection, connect
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

        return cls(
            left, right, clean(management_lob), clean(planning_group),
            clean(staff_type), clean(team_leader), clean(agent_id),
        )


class DashboardData:
    """Small read-only application service over the durable WFMHub database."""

    def __init__(self, config: Config):
        self.config = config
        self.capacity: CapacityMapping = load_capacity_mapping(config.capacity_mapping)
        self.services: ServiceProfileCatalog = load_service_profiles(
            config.home, config.service_profiles,
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

    def latest_date(self) -> date:
        conn = self._connect()
        try:
            value = conn.execute(
                """SELECT max(business_date) FROM (
                     SELECT business_date FROM mart.call_service_15min
                     UNION ALL SELECT business_date FROM mart.attendance_agent_day
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
            "product": "WFMHub Operations Console",
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
                      last_logout, attendance_result, call_action, requires_call,
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
        due = [row for row in attendance if _number(row.get("planned_work_minutes")) > 0]
        present = [row for row in due if row.get("attendance_result") in PRESENT_RESULTS]
        no_show = [row for row in due if row.get("attendance_result") in NO_SHOW_RESULTS]
        unknown = [row for row in due if row not in present and row not in no_show]
        calls = [
            self._attendance_record(row) for row in attendance
            if bool(row.get("requires_call"))
            or _number(row.get("uncoded_late_minutes")) > 0
            or _number(row.get("uncoded_early_leave_minutes")) > 0
        ]
        by_lob: dict[str, dict[str, Any]] = {}
        all_lobs = {row["management_lob"] for row in service} | {
            row["management_lob"] for row in attendance
        }
        for lob in sorted(all_lobs):
            lob_service = self._service_bucket(row for row in service if row["management_lob"] == lob)
            lob_att = [row for row in due if row["management_lob"] == lob]
            lob_present = [row for row in lob_att if row.get("attendance_result") in PRESENT_RESULTS]
            lob_no_show = [row for row in lob_att if row.get("attendance_result") in NO_SHOW_RESULTS]
            by_lob[lob] = {
                "management_lob": lob, **lob_service,
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
            by_lob.append({"management_lob": key, **self._service_bucket(values)})
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

    def _attendance_filter_from_dimension(self, scope: DashboardFilter) -> list[dict[str, Any]]:
        conn = self._connect()
        try:
            return self._attendance_rows(conn, scope)
        finally:
            conn.close()
