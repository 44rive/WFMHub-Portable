// Query name: PCS Agent Day
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = Excel.CurrentWorkbook(){[Name="pFeedFolder"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_AgentDay.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"business_date", type date},
        {"agent_id", type text}, {"agent_name", type text},
        {"team_leader", type text}, {"ops_manager", type text},
        {"lob", type text}, {"language", type text},
        {"inbound_call_legs", Int64.Type}, {"pcs_status_1", Int64.Type},
        {"q1_nonblank", Int64.Type}, {"valid_q1", Int64.Type},
        {"q1_score_sum", type number}, {"pcs_average", type number},
        {"participation_rate", type number}, {"score_le_3", Int64.Type},
        {"score_gt_3", Int64.Type}, {"invalid_q1", Int64.Type},
        {"sample_state", type text}
    }, "en-US")
in
    Typed
