// Query name: PCS Trend
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = Excel.CurrentWorkbook(){[Name="pFeedFolder"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_Trend.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"business_date", type date}, {"pcs_score_sum", type number},
        {"valid_responses", Int64.Type}, {"participating_responses", Int64.Type},
        {"eligible_calls", Int64.Type}, {"pcs_average", type number},
        {"participation_rate", type number}, {"score_le_3", Int64.Type},
        {"score_gt_3", Int64.Type}, {"coaching_completed", Int64.Type},
        {"actions_rate", type number}
    }, "en-US")
in
    Typed
