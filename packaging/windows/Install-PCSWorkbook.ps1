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
    [string]$DataQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$CoachingQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$LobQueryPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultsQueryPath,

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
            $table.Unlist()
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
        try { $range.Clear() } finally {
            Release-ComObject $range
            Release-ComObject $first
            Release-ComObject $last
        }
    }
    finally {
        Release-ComObject $sheet
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
            $queryTable.RefreshStyle = 1
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

    if ($Action -eq "Install") {
        foreach ($path in @($DataQueryPath, $CoachingQueryPath, $LobQueryPath, $ResultsQueryPath)) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Power Query definition not found: $path"
            }
        }
        Remove-WorkbookQuery "PCS_DATA"
        Remove-WorkbookQuery "COACHING_QUEUE"
        Remove-WorkbookQuery "PCS_LOB"
        Remove-WorkbookQuery "PCS_RESULTS"
        Remove-StarterTable "OVERVIEW" "tblPcsLob" 34 19
        Remove-StarterTable "RESULTS" "tblResults" 4 21
        Remove-StarterTable "COACHING_QUEUE" "tblCoachingQueue" 4 13
        Remove-StarterTable "PCS_DATA" "tblPcsData" 4 26
        Set-SetupValue "Connection Mode" $Mode
        Set-SetupValue "Local Feed Folder" ([System.IO.Path]::GetFullPath($FeedFolder))
        Add-QueryTable "PCS_LOB" ([System.IO.File]::ReadAllText($LobQueryPath)) "OVERVIEW" "tblPcsLob" "A34"
        Add-QueryTable "PCS_RESULTS" ([System.IO.File]::ReadAllText($ResultsQueryPath)) "RESULTS" "tblResults" "A4"
        Add-QueryTable "COACHING_QUEUE" ([System.IO.File]::ReadAllText($CoachingQueryPath)) "COACHING_QUEUE" "tblCoachingQueue" "A4"
        Add-QueryTable "PCS_DATA" ([System.IO.File]::ReadAllText($DataQueryPath)) "PCS_DATA" "tblPcsData" "A4"
        Set-SetupValue "Power Query Installed" "YES"
        Set-SetupValue "Last Installer Result" "Installed successfully"
    }

    $script:Workbook.RefreshAll()
    Wait-ForRefresh
    $script:Excel.CalculateFullRebuild()
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
