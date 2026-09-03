// Create one copy of this query per model CSV.
// Replace TABLE_FILE.csv with the file named in output/model_data/<report>/manifest.json.
// Set the query to "Only Create Connection" and tick "Add this data to the Data Model".
let
    ModelRoot = Excel.CurrentWorkbook(){[Name="pModelDataPath"]}[Content]{0}[Column1],
    Source = Csv.Document(
        File.Contents(ModelRoot & "\\TABLE_FILE.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders
