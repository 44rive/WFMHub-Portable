// Query name: PCS Summary
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = Excel.CurrentWorkbook(){[Name="pFeedFolder"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_Summary.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"period", type text}, {"period_start", type date}, {"period_end", type date},
        {"pcs_average", type number}, {"participation_rate", type number},
        {"valid_q1", Int64.Type}, {"pcs_status_1", Int64.Type},
        {"score_le_3", Int64.Type}, {"score_gt_3", Int64.Type},
        {"inbound_call_legs", Int64.Type}, {"coaching_completed", Int64.Type},
        {"actions_rate", type number}
    }, "en-US")
in
    Typed
