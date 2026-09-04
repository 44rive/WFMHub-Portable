// Query name: PCS Dates
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = "__PCS_FEED_FOLDER__",
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_Dates.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"date", type date}, {"year", Int64.Type},
        {"month_number", Int64.Type}, {"month_key", type text},
        {"month_label", type text}, {"iso_week", Int64.Type},
        {"weekday", type text}, {"day_of_month", Int64.Type}
    }, "en-US")
in
    Typed
