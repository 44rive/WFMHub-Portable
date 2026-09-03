// Query name: PCS Actions
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = Excel.CurrentWorkbook(){[Name="pFeedFolder"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_Actions.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"coaching_key", type text}, {"business_date", type date},
        {"call_start", type datetime}, {"call_reference_number", type text},
        {"call_id", type text}, {"agent_id", type text}, {"agent_name", type text},
        {"team_leader", type text}, {"ops_manager", type text}, {"lob", type text},
        {"language", type text}, {"queue", type text}, {"q1_answer", type text},
        {"q1_score", type number}, {"q2_answer", type text},
        {"customer_comment", type text}, {"pcs_status", type text},
        {"post_call_survey_mode", type text}, {"source_file", type text},
        {"coaching_status", type text}, {"coaching_date", type date},
        {"coach", type text}, {"coaching_comment", type text}
    }, "en-US")
in
    Typed
