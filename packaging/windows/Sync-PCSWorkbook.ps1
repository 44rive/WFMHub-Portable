[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet("Sync")][string]$Action,
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$FeedFolder,
    [switch]$OpenAfter
)

# The six read-only CSV tables are resized in place. Human coaching, formulas,
# charts and dropdown names are never deleted or re-created.
$ErrorActionPreference = "Stop"
$script:Excel = $null
$script:Workbook = $null

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Set-SetupValue {
    param([string]$Setting, [string]$Value)
    $sheet = $script:Workbook.Worksheets.Item("SETUP")
    try {
        $table = $sheet.ListObjects.Item("tblSetup")
        try {
            $labels = $table.ListColumns.Item("Setting").DataBodyRange
            $values = $table.ListColumns.Item("Value").DataBodyRange
            try {
                for ($row = 1; $row -le $labels.Rows.Count; $row++) {
                    if ([string]$labels.Cells.Item($row, 1).Value2 -eq $Setting) {
                        $values.Cells.Item($row, 1).Value2 = $Value
                        return
                    }
                }
            }
            finally { Release-ComObject $labels; Release-ComObject $values }
        }
        finally { Release-ComObject $table }
    }
    finally { Release-ComObject $sheet }
    throw "PCS SETUP is missing '$Setting'; rebuild the tracker from WFMHub."
}

function Sync-FeedTable {
    param([string]$FileName, [string]$SheetName, [string]$TableName, [string[]]$NumericColumns)
    $path = Join-Path $FeedFolder $FileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "PCS feed missing: $path. Run Update PCS data first."
    }
    $records = @(Import-Csv -LiteralPath $path -Encoding UTF8)
    $sheet = $script:Workbook.Worksheets.Item($SheetName)
    try {
        $table = $sheet.ListObjects.Item($TableName)
        try {
            $headers = @()
            for ($column = 1; $column -le $table.ListColumns.Count; $column++) {
                $headers += [string]$table.ListColumns.Item($column).Name
            }
            if ($records.Count -gt 0) {
                $incoming = @($records[0].PSObject.Properties | ForEach-Object { $_.Name })
                if (($incoming -join [char]31) -ne ($headers -join [char]31)) {
                    throw "$FileName columns do not match $TableName. Run Update PCS data with the matching Hub release."
                }
            }
            $rowCount = [Math]::Max(1, $records.Count)
            $columnCount = $headers.Count
            if ($rowCount -gt 1000000) { throw "$FileName exceeds Excel's row limit." }
            $oldData = $table.DataBodyRange
            if ($null -ne $oldData) {
                try { $oldData.ClearContents() }
                finally { Release-ComObject $oldData }
            }
            $top = $sheet.Cells.Item(4, 1)
            $bottom = $sheet.Cells.Item(4 + $rowCount, $columnCount)
            $newRange = $sheet.Range($top, $bottom)
            try { $table.Resize($newRange) }
            finally {
                Release-ComObject $newRange
                Release-ComObject $top
                Release-ComObject $bottom
            }
            $data = [object[,]]::new($rowCount, $columnCount)
            for ($row = 0; $row -lt $records.Count; $row++) {
                for ($column = 0; $column -lt $columnCount; $column++) {
                    $raw = [string]$records[$row].($headers[$column])
                    if ($NumericColumns -contains $headers[$column] -and $raw -ne "") {
                        $parsed = 0.0
                        if (-not [double]::TryParse($raw,
                                [System.Globalization.NumberStyles]::Float,
                                [System.Globalization.CultureInfo]::InvariantCulture,
                                [ref]$parsed)) {
                            throw "$FileName has a nonnumeric '$($headers[$column])' at data row $($row + 1)."
                        }
                        $data[$row, $column] = $parsed
                    }
                    else { $data[$row, $column] = $raw }
                }
            }
            $destination = $sheet.Range($sheet.Cells.Item(5, 1), $sheet.Cells.Item(4 + $rowCount, $columnCount))
            try { $destination.Value2 = $data }
            finally { Release-ComObject $destination }
            Write-Output "${TableName}: $($records.Count) rows"
        }
        finally { Release-ComObject $table }
    }
    finally { Release-ComObject $sheet }
}

try {
    $resolvedWorkbook = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (-not (Test-Path -LiteralPath $resolvedWorkbook -PathType Leaf)) {
        throw "PCS tracker not found: $resolvedWorkbook"
    }
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    $script:Excel.ScreenUpdating = $false
    $script:Excel.EnableEvents = $false
    $script:Workbook = $script:Excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    if ($script:Workbook.ReadOnly) {
        throw "PCS tracker is open or locked by OneDrive. Close Excel, wait for sync, then retry."
    }
    if ($script:Workbook.Queries.Count -gt 0) {
        throw "This is the old Power Query tracker. Run Update PCS data with this release to migrate it once."
    }
    # Recalculate once after all six bulk writes, not after every table resize.
    $script:Excel.Calculation = -4135  # xlCalculationManual
    Sync-FeedTable "PCS_FILTER_LIST_CURRENT.csv" "_PCS_FILTERS" "tblPcsFilters" @()
    Sync-FeedTable "PCS_LOB_SCORECARD_CURRENT.csv" "_PCS_LOB" "tblPcsLob" @(
        "Rank", "Valid Q1", "PCS", "Prior PCS", "Change", "Participation", "Coaching Due")
    Sync-FeedTable "PCS_AGENT_SCORECARD_CURRENT.csv" "_PCS_AGENT" "tblPcsAgent" @(
        "Rank", "PCS", "Prior PCS", "Change", "Participation", "Valid Q1", "Coaching Due")
    Sync-FeedTable "PCS_DAILY_SCORECARD_CURRENT.csv" "_PCS_DAILY" "tblPcsDaily" @(
        "Rank", "PCS", "Participation", "Valid Q1", "Coaching Due")
    Sync-FeedTable "PCS_RESULTS_CURRENT.csv" "PERFORMANCE" "tblPcsPerformance" @(
        "PCS Average", "Participation Rate", "Valid Q1", "PCS Status 1",
        "Q1 Nonblank", "Score <= 3", "Score > 3", "Inbound Call Legs")
    Sync-FeedTable "PCS_COACHING_OPPORTUNITY_CURRENT.csv" "_PCS_COACH" "tblPcsCoachingView" @(
        "Rank", "Q1 Score")
    foreach ($name in @("PCS_PERIOD_LIST", "PCS_LOB_LIST", "PCS_TL_ACTIVE", "PCS_AGENT_ACTIVE")) {
        $defined = $script:Workbook.Names.Item($name)
        try {
            if ([string]$defined.RefersTo -like "*#REF!*") {
                throw "PCS selector $name contains #REF!; the workbook was changed outside the Hub."
            }
        }
        finally { Release-ComObject $defined }
    }
    Set-SetupValue "Feed Sync" "READY"
    Set-SetupValue "Local Feed Folder" ([System.IO.Path]::GetFullPath($FeedFolder))
    Set-SetupValue "Workbook Last Refreshed" ([DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss"))
    $script:Excel.CalculateFullRebuild()
    $script:Excel.Calculation = -4105  # xlCalculationAutomatic
    $script:Workbook.Save()
    Write-Output "PCS TRACKER SYNCED | $resolvedWorkbook"
}
catch {
    Write-Error ("PCS feed sync failed: " + $_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $script:Workbook) {
        try { $script:Workbook.Close($false) } catch {}
        Release-ComObject $script:Workbook
    }
    if ($null -ne $script:Excel) {
        try { $script:Excel.Calculation = -4105 } catch {}
        try { $script:Excel.EnableEvents = $true } catch {}
        try { $script:Excel.Quit() } catch {}
        Release-ComObject $script:Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ($OpenAfter) { Start-Process -FilePath ([System.IO.Path]::GetFullPath($WorkbookPath)) }
