"""Copy this file, remove the leading underscore, then paste your custom code."""


def run(ctx):
    # The database connection behind ctx is read-only. Use logical names such
    # as mart.agent_pcs_day, core.clean_call_leg or mart.attendance_agent_day.
    result = ctx.query(
        """SELECT agent_id, agent_name, sum(handled_calls) AS handled_calls,
                  CASE WHEN sum(pcs_score_count)>0
                       THEN 1.0*sum(pcs_score_sum)/sum(pcs_score_count) END AS pcs_average
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY agent_id, agent_name
           ORDER BY pcs_average""",
        [ctx.start, ctx.end],
    )

    # Paste more Python here. result.headers and result.rows are ordinary
    # Python lists, so no pandas installation is required.

    return ctx.write_csv("my_agent_analysis", result)
