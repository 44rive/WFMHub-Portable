// Create one copy of this query per model CSV.
// Replace TABLE_FILE.csv with the file named in output/model_data/<report>/manifest.json.
// Set the query to "Only Create Connection" and tick "Add this data to the Data Model".
// ExistingTypes is essential: without it Excel imports CSV numbers as text and SUM fails.
let
    ModelRoot = Excel.CurrentWorkbook(){[Name="pModelDataPath"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(ModelRoot & "\\TABLE_FILE.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    RequestedTypes = {
        {"business_date", type date}, {"period_start", type date}, {"period_end", type date},
        {"generated_at", type datetime}, {"scheduled_start", type datetime},
        {"scheduled_end", type datetime}, {"actual_first_seen", type datetime},
        {"actual_last_seen", type datetime}, {"evaluation_as_of", type datetime},
        {"hour_start", type datetime}, {"interval_start", type datetime},
        {"interval_end", type datetime}, {"gap_start", type datetime},
        {"gap_end", type datetime}, {"segment_start", type datetime},
        {"segment_end", type datetime}, {"injected_date", type date},

        {"inbound_call_legs", Int64.Type}, {"pcs_status_1", Int64.Type},
        {"q1_nonblank", Int64.Type}, {"valid_q1", Int64.Type},
        {"q1_score_sum", type number}, {"pcs_average", type number},
        {"participation_rate", type number}, {"score_le_3", Int64.Type},
        {"score_gt_3", Int64.Type}, {"invalid_q1", Int64.Type},
        {"pcs_score_sum", type number}, {"valid_responses", Int64.Type},
        {"participating_responses", Int64.Type}, {"eligible_calls", Int64.Type},

        {"scheduled_working", Int64.Type}, {"call_actions", Int64.Type},
        {"no_shows", Int64.Type}, {"late", Int64.Type},
        {"missing_evidence", Int64.Type}, {"uncoded_late_minutes", type number},
        {"no_show_minutes", type number}, {"source_loaded", Int64.Type},
        {"is_provisional", Int64.Type}, {"gap_minutes", type number},
        {"segment_minutes", type number}, {"verint_overlap_minutes", type number},
        {"is_gap", Int64.Type},

        {"offered", type number}, {"answered", type number}, {"abandoned", type number},
        {"short_abandoned", type number}, {"answered_within_target", type number},
        {"handled_seconds", type number}, {"service_level", type number},
        {"service_availability", type number}, {"aht_seconds", type number},
        {"actual_volume", type number}, {"forecast_volume", type number},
        {"variance_volume", type number}, {"absolute_error_volume", type number},
        {"forecast_attainment", type number}, {"absolute_error_rate", type number},

        {"scheduled_agents", type number}, {"observed_agents", type number},
        {"productive_agents", type number}, {"scheduled_fte", type number},
        {"elapsed_scheduled_fte", type number}, {"observed_fte", type number},
        {"productive_fte", type number}, {"staffing_variance_fte", type number},
        {"staffing_gap_fte", type number},

        {"planned_net_hours", type number}, {"final_absence_hours", type number},
        {"final_vacation_hours", type number}, {"final_unpaid_hours", type number},
        {"final_shrinkage_hours", type number}, {"final_unmapped_hours", type number},
        {"final_absence_rate", type number}, {"final_absence_day", Int64.Type},
        {"planned_hours", type number}, {"absence_hours", type number},
        {"absence_rate", type number}, {"vacation_hours", type number},
        {"unpaid_hours", type number}, {"shrinkage_hours", type number},
        {"shrinkage_rate", type number},

        {"aht_earned", type number}, {"productivity_earned", type number},
        {"pcs_earned", type number}, {"participation_earned", type number},
        {"qm_earned", type number}, {"absence_earned", type number},
        {"extra_pcs_earned", type number}, {"gross_achievement", type number},
        {"voc_malus", type number}, {"final_achievement", type number},
        {"reference_bonus", type number}, {"proration", type number},
        {"scenario_payout", type number}, {"released_payout", type number},
        {"actual_value", type number}, {"earned_weight", type number},
        {"tier1_target", type number}, {"tier2_target", type number},

        {"finding_rank", Int64.Type}, {"current_value", type number},
        {"reference_value", type number}, {"target_value", type number},
        {"delta_value", type number}, {"Current Value", type number},
        {"Reference Value", type number}, {"Delta", type number},
        {"Numerator", type number}, {"Denominator", type number}, {"Target", type number}
    },
    ExistingTypes = List.Select(
        RequestedTypes,
        each List.Contains(Table.ColumnNames(PromotedHeaders), _{0})
    ),
    Typed = Table.TransformColumnTypes(PromotedHeaders, ExistingTypes, "en-US")
in
    Typed
