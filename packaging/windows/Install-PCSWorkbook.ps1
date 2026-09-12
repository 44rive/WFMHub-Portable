[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Install", "Refresh")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [ValidateSet("LOCAL", "SHAREPOINT")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$FeedFolder,

    [Parameter(Mandatory = $true)]
    [string]$FilterQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$AgentQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$CoachingQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$LobQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultsQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$DailyQueryPath,

    [switch]$OpenAfter
)

$ErrorActionPreference = "Stop"
$script:Excel = $null
$script:Workbook = $null
$script:Saved = $false

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-SetupTable {
    $sheet = $script:Workbook.Worksheets.Item("SETUP")
    try {
        return $sheet.ListObjects.Item("tblSetup")
    }
    finally {
        Release-ComObject $sheet
    }
}

function Set-SetupValue {
    param([string]$Setting, [string]$Value)
    $table = Get-SetupTable
    try {
        $settingColumn = $table.ListColumns.Item("Setting").DataBodyRange
        $valueColumn = $table.ListColumns.Item("Value").DataBodyRange
        try {
            for ($row = 1; $row -le $settingColumn.Rows.Count; $row++) {
                if ([string]$settingColumn.Cells.Item($row, 1).Value2 -eq $Setting) {
                    $valueColumn.Cells.Item($row, 1).Value2 = $Value
                    return
                }
            }
            throw "SETUP is missing the '$Setting' row. Rebuild the PCS tracker before installing."
        }
        finally {
            Release-ComObject $settingColumn
            Release-ComObject $valueColumn
        }
    }
    finally {
        Release-ComObject $table
    }
}

function Remove-WorkbookQuery {
    param([string]$Name)
    for ($index = $script:Workbook.Queries.Count; $index -ge 1; $index--) {
        $query = $script:Workbook.Queries.Item($index)
        try {
            if ([string]$query.Name -eq $Name) {
                $query.Delete()
            }
        }
        finally {
            Release-ComObject $query
        }
    }
    for ($index = $script:Workbook.Connections.Count; $index -ge 1; $index--) {
        $connection = $script:Workbook.Connections.Item($index)
        try {
            if ([string]$connection.Name -eq "Query - $Name") {
                $connection.Delete()
            }
        }
        finally {
            Release-ComObject $connection
        }
    }
}

function Remove-StarterTable {
    param(
        [string]$SheetName,
        [string]$TableName,
        [int]$FirstRow,
        [int]$LastColumn
    )
    $sheet = $script:Workbook.Worksheets.Item($SheetName)
    try {
        try {
            $table = $sheet.ListObjects.Item($TableName)
            $table.Delete()
            Release-ComObject $table
        }
        catch {
            # A partially installed workbook may already have lost the starter table.
        }
        $used = $sheet.UsedRange
        try {
            $usedLastRow = [int]$used.Row + [int]$used.Rows.Count - 1
        }
        finally {
            Release-ComObject $used
        }
        $lastRow = [Math]::Max($FirstRow + 1, $usedLastRow)
        $first = $sheet.Cells.Item($FirstRow, 1)
        $last = $sheet.Cells.Item($lastRow, $LastColumn)
        $range = $sheet.Range($first, $last)
        try { $range.ClearContents() } finally {
            Release-ComObject $range
            Release-ComObject $first
            Release-ComObject $last
        }
    }
    finally {
        Release-ComObject $sheet
    }
}

function Set-TableColumnFormat {
    param(
        [string]$SheetName,
        [string]$TableName,
        [string]$ColumnName,
        [string]$NumberFormat
    )
    $sheet = $script:Workbook.Worksheets.Item($SheetName)
    try {
        $table = $sheet.ListObjects.Item($TableName)
        try {
            $column = $table.ListColumns.Item($ColumnName)
            try {
                $range = $column.DataBodyRange
                if ($null -ne $range) {
                    try { $range.NumberFormat = $NumberFormat }
                    finally { Release-ComObject $range }
                }
            }
            finally { Release-ComObject $column }
        }
        finally { Release-ComObject $table }
    }
    finally { Release-ComObject $sheet }
}

function Set-SelectorNames {
    $keys = "'_PCS_FILTERS'!`$A:`$A"
    $values = "'_PCS_FILTERS'!`$C:`$C"
    $teamGroup = '"TL|"&SUBSTITUTE(OVERVIEW!$H$3,"|","/")'
    $agentGroup = '"AGENT|"&SUBSTITUTE(OVERVIEW!$H$3,"|","/")&"|"&SUBSTITUTE(OVERVIEW!$O$3,"|","/")'
    $groups = [ordered]@{
        "PCS_PERIOD_LIST" = '"PERIOD"'
        "PCS_LOB_LIST" = '"LOB"'
        "PCS_TL_ACTIVE" = $teamGroup
        "PCS_AGENT_ACTIVE" = $agentGroup
    }
    foreach ($entry in $groups.GetEnumerator()) {
        $group = [string]$entry.Value
        $first = "MATCH($group,$keys,0)"
        $formula = "=INDEX($values,$first):INDEX($values,$first+COUNTIF($keys,$group)-1)"
        $existing = $null
        try {
            $existing = $script:Workbook.Names.Item([string]$entry.Key)
            $existing.Delete()
        }
        catch {}
        finally { Release-ComObject $existing }
        $created = $script:Workbook.Names.Add([string]$entry.Key, $formula)
        Release-ComObject $created
    }
}

function Assert-SelectorIntegrity {
    $overview = $script:Workbook.Worksheets.Item("OVERVIEW")
    try {
        $selectedLob = [string]$overview.Range("H3").Value2
        $selectedTeam = [string]$overview.Range("O3").Value2
    }
    finally { Release-ComObject $overview }
    $checks = @(
        @("PCS_PERIOD_LIST", "PERIOD", "Current MTD"),
        @("PCS_LOB_LIST", "LOB", "All"),
        @("PCS_TL_ACTIVE", ("TL|" + $selectedLob.Replace("|", "/")), "All"),
        @("PCS_AGENT_ACTIVE", ("AGENT|" + $selectedLob.Replace("|", "/") + "|" + $selectedTeam.Replace("|", "/")), "All")
    )
    $filterSheet = $script:Workbook.Worksheets.Item("_PCS_FILTERS")
    try {
        foreach ($check in $checks) {
            $definedName = $null
            $resolved = $null
            try {
                $definedName = $script:Workbook.Names.Item([string]$check[0])
                $resolved = $definedName.RefersToRange
                if ($null -eq $resolved -or $resolved.Rows.Count -lt 1) {
                    throw "$($check[0]) does not resolve to a populated range"
                }
                $seenRequired = $false
                foreach ($cell in $resolved.Cells) {
                    try {
                        $actualGroup = [string]$filterSheet.Cells.Item($cell.Row, 1).Value2
                        if ($actualGroup -ne [string]$check[1]) {
                            throw "$($check[0]) crossed into '$actualGroup' instead of '$($check[1])' at row $($cell.Row)"
                        }
                        if ([string]$cell.Value2 -eq [string]$check[2]) {
                            $seenRequired = $true
                        }
                    }
                    finally { Release-ComObject $cell }
                }
                if (-not $seenRequired) {
                    throw "$($check[0]) is missing required value '$($check[2])'"
                }
            }
            finally {
                Release-ComObject $resolved
                Release-ComObject $definedName
            }
        }
    }
    finally { Release-ComObject $filterSheet }
}

function Assert-PresentationIntegrity {
    $problems = @()
    foreach ($sheetName in @("OVERVIEW", "COACHING", "_PCS_CALC")) {
        $sheet = $script:Workbook.Worksheets.Item($sheetName)
        try {
            $used = $sheet.UsedRange
            try {
                $formulaCells = $null
                try { $formulaCells = $used.SpecialCells(-4123) } catch {}
                if ($null -ne $formulaCells) {
                    try {
                        foreach ($cell in $formulaCells.Cells) {
                            try {
                                $formula = [string]$cell.Formula
                                if ($formula -like "*#REF!*") {
                                    $problems += "$sheetName!$([string]$cell.Address) contains #REF!"
                                }
                                if ($formula -like "*tblPcs*") {
                                    $problems += "$sheetName!$([string]$cell.Address) still uses a replaceable table reference"
                                }
                            }
                            finally { Release-ComObject $cell }
                        }
                    }
                    finally { Release-ComObject $formulaCells }
                }
            }
            finally { Release-ComObject $used }
        }
        finally { Release-ComObject $sheet }
    }

    foreach ($requiredName in @("PCS_LOB_DATA", "PCS_AGENT_DATA", "PCS_DAILY_DATA", "PCS_COACH_DATA")) {
        $definedName = $null
        try {
            $definedName = $script:Workbook.Names.Item($requiredName)
            $refersTo = [string]$definedName.RefersTo
            if ([string]::IsNullOrWhiteSpace($refersTo) -or $refersTo -like "*#REF!*") {
                $problems += "$requiredName is broken"
            }
        }
        catch {
            $problems += "$requiredName is missing"
        }
        finally { Release-ComObject $definedName }
    }

    foreach ($check in @(
        @("OVERVIEW", "A6", "PCS_LOB_DATA"),
        @("OVERVIEW", "H6", "PCS_LOB_DATA"),
        @("_PCS_CALC", "B2", "PCS_LOB_DATA"),
        @("_PCS_CALC", "E2", "PCS_DAILY_DATA"),
        @("COACHING", "A5", "PCS_COACH_DATA")
    )) {
        $sheet = $script:Workbook.Worksheets.Item($check[0])
        try {
            $cell = $sheet.Range($check[1])
            try {
                $formula = [string]$cell.Formula
                if ($formula -notlike "*$($check[2])*") {
                    $problems += "$($check[0])!$($check[1]) lost its governed lookup formula"
                }
            }
            finally { Release-ComObject $cell }
        }
        finally { Release-ComObject $sheet }
    }

    if ($problems.Count -gt 0) {
        throw ("PCS presentation integrity failed: " + ($problems -join "; "))
    }
}

function Add-QueryTable {
    param(
        [string]$QueryName,
        [string]$Formula,
        [string]$SheetName,
        [string]$TableName,
        [string]$DestinationAddress
    )
    [void]$script:Workbook.Queries.Add($QueryName, $Formula)
    $sheet = $script:Workbook.Worksheets.Item($SheetName)
    $destination = $sheet.Range($DestinationAddress)
    $source = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$QueryName;Extended Properties=`"`""
    try {
        $table = $sheet.ListObjects.Add(0, $source, $null, 1, $destination)
        $table.Name = $TableName
        $table.TableStyle = "TableStyleLight9"
        $queryTable = $table.QueryTable
        try {
            $queryTable.CommandType = 2
            $queryTable.CommandText = "SELECT * FROM [$QueryName]"
            # Overwrite the query destination in place. Insert/delete-cell
            # refresh can shift workbook names by one row at group boundaries.
            $queryTable.RefreshStyle = 0
            $queryTable.BackgroundQuery = $false
            $queryTable.AdjustColumnWidth = $false
            $queryTable.PreserveFormatting = $true
            [void]$queryTable.Refresh($false)
        }
        finally {
            Release-ComObject $queryTable
            Release-ComObject $table
        }
    }
    finally {
        Release-ComObject $destination
        Release-ComObject $sheet
    }
}

function Wait-ForRefresh {
    param([int]$TimeoutSeconds = 300)
    try { $script:Excel.CalculateUntilAsyncQueriesDone() } catch {}
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $refreshing = $false
        foreach ($sheet in $script:Workbook.Worksheets) {
            try {
                foreach ($table in $sheet.ListObjects) {
                    try {
                        if ($null -ne $table.QueryTable -and $table.QueryTable.Refreshing) {
                            $refreshing = $true
                        }
                    }
                    catch {}
                    finally { Release-ComObject $table }
                }
            }
            finally { Release-ComObject $sheet }
        }
        if (-not $refreshing -and $script:Excel.CalculationState -eq 0) {
            return
        }
        Start-Sleep -Milliseconds 500
    }
    throw "Excel did not finish the PCS refresh within $TimeoutSeconds seconds."
}

try {
    $resolvedWorkbook = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (-not (Test-Path -LiteralPath $resolvedWorkbook -PathType Leaf)) {
        throw "PCS workbook not found: $resolvedWorkbook"
    }
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    $script:Workbook = $script:Excel.Workbooks.Open($resolvedWorkbook, 0, $false)
    if ($script:Workbook.ReadOnly) {
        throw "PCS workbook opened read-only. It is probably already open in Excel or locked by OneDrive. Close it, wait for sync to finish, then retry."
    }

    if ($Action -eq "Install") {
        foreach ($path in @($FilterQueryPath, $AgentQueryPath, $CoachingQueryPath, $LobQueryPath, $ResultsQueryPath, $DailyQueryPath)) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Power Query definition not found: $path"
            }
        }
        Remove-StarterTable "_PCS_FILTERS" "tblPcsFilters" 4 3
        Remove-StarterTable "_PCS_LOB" "tblPcsLob" 4 12
        Remove-StarterTable "_PCS_AGENT" "tblPcsAgent" 4 14
        Remove-StarterTable "_PCS_DAILY" "tblPcsDaily" 4 9
        Remove-StarterTable "PERFORMANCE" "tblPcsPerformance" 4 21
        Remove-StarterTable "_PCS_COACH" "tblPcsCoachingView" 4 14
        Remove-WorkbookQuery "PCS_DATA"
        Remove-WorkbookQuery "COACHING_QUEUE"
        Remove-WorkbookQuery "PCS_FILTERS"
        Remove-WorkbookQuery "PCS_COACHING_VIEW"
        Remove-WorkbookQuery "PCS_LOB"
        Remove-WorkbookQuery "PCS_AGENT"
        Remove-WorkbookQuery "PCS_DAILY"
        Remove-WorkbookQuery "PCS_RESULTS"
        Remove-WorkbookQuery "PCS_SCOPE"
        Set-SetupValue "Connection Mode" $Mode
        Set-SetupValue "Local Feed Folder" ([System.IO.Path]::GetFullPath($FeedFolder))
        Add-QueryTable "PCS_FILTERS" ([System.IO.File]::ReadAllText($FilterQueryPath)) "_PCS_FILTERS" "tblPcsFilters" "A4"
        Add-QueryTable "PCS_LOB" ([System.IO.File]::ReadAllText($LobQueryPath)) "_PCS_LOB" "tblPcsLob" "A4"
        Add-QueryTable "PCS_AGENT" ([System.IO.File]::ReadAllText($AgentQueryPath)) "_PCS_AGENT" "tblPcsAgent" "A4"
        Add-QueryTable "PCS_DAILY" ([System.IO.File]::ReadAllText($DailyQueryPath)) "_PCS_DAILY" "tblPcsDaily" "A4"
        Add-QueryTable "PCS_RESULTS" ([System.IO.File]::ReadAllText($ResultsQueryPath)) "PERFORMANCE" "tblPcsPerformance" "A4"
        Add-QueryTable "PCS_COACHING_VIEW" ([System.IO.File]::ReadAllText($CoachingQueryPath)) "_PCS_COACH" "tblPcsCoachingView" "A4"
        Set-SetupValue "Power Query Installed" "YES"
        Set-SetupValue "Last Installer Result" "Installed successfully"
    }

    $script:Workbook.RefreshAll()
    Wait-ForRefresh
    foreach ($format in @(
        @("_PCS_LOB", "tblPcsLob", "Data Through", "yyyy-mm-dd"),
        @("_PCS_LOB", "tblPcsLob", "Feed Refreshed At", "yyyy-mm-dd hh:mm"),
        @("_PCS_AGENT", "tblPcsAgent", "Data Through", "yyyy-mm-dd"),
        @("_PCS_AGENT", "tblPcsAgent", "Feed Refreshed At", "yyyy-mm-dd hh:mm"),
        @("_PCS_DAILY", "tblPcsDaily", "Date", "yyyy-mm-dd"),
        @("_PCS_DAILY", "tblPcsDaily", "Data Through", "yyyy-mm-dd"),
        @("_PCS_DAILY", "tblPcsDaily", "Feed Refreshed At", "yyyy-mm-dd hh:mm"),
        @("PERFORMANCE", "tblPcsPerformance", "Period Start", "yyyy-mm-dd"),
        @("PERFORMANCE", "tblPcsPerformance", "Period End", "yyyy-mm-dd"),
        @("PERFORMANCE", "tblPcsPerformance", "Data Through", "yyyy-mm-dd"),
        @("PERFORMANCE", "tblPcsPerformance", "Feed Refreshed At", "yyyy-mm-dd hh:mm"),
        @("_PCS_COACH", "tblPcsCoachingView", "Date", "yyyy-mm-dd"),
        @("_PCS_COACH", "tblPcsCoachingView", "Data Through", "yyyy-mm-dd"),
        @("_PCS_COACH", "tblPcsCoachingView", "Feed Refreshed At", "yyyy-mm-dd hh:mm")
    )) {
        Set-TableColumnFormat $format[0] $format[1] $format[2] $format[3]
    }
    Set-SelectorNames
    $script:Excel.CalculateFullRebuild()
    Assert-SelectorIntegrity
    Assert-PresentationIntegrity
    Set-SetupValue "Workbook Last Refreshed" ([DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss"))
    Set-SetupValue "Last Installer Result" "Refresh completed"
    $script:Workbook.Save()
    $script:Saved = $true
    Write-Output "PCS WORKBOOK READY | $Action | $Mode | $resolvedWorkbook"
}
catch {
    Write-Error ("PCS Excel automation failed: " + $_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $script:Workbook) {
        try { $script:Workbook.Close($script:Saved) } catch {}
        Release-ComObject $script:Workbook
    }
    if ($null -ne $script:Excel) {
        try { $script:Excel.Quit() } catch {}
        Release-ComObject $script:Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($OpenAfter) {
    Start-Process -FilePath ([System.IO.Path]::GetFullPath($WorkbookPath))
}
