// Query name: PCS Agents
// Load To: Only Create Connection + Add this data to the Data Model
let
    FeedRoot = "__PCS_FEED_FOLDER__",
    Source = Csv.Document(
        File.Contents(FeedRoot & "\\PCS_Agents.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(PromotedHeaders, {
        {"agent_id", type text}, {"agent_name", type text},
        {"team_leader", type text}, {"ops_manager", type text},
        {"lob", type text}, {"language", type text}
    }, "en-US")
in
    Typed
