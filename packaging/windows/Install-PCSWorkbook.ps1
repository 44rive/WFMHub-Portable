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

function Save-FormulaState {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($sheet in $script:Workbook.Worksheets) {
        try {
            if ($sheet.Name -in @("PCS_DATA", "COACHING_QUEUE")) {
                continue
            }
            $used = $sheet.UsedRange
            try {
                foreach ($cell in $used.Cells) {
                    try {
                        if ($cell.HasFormula) {
                            $formula = $null
                            try { $formula = [string]$cell.Formula2 } catch { $formula = [string]$cell.Formula }
                            $items.Add([pscustomobject]@{
                                Sheet = [string]$sheet.Name
                                Address = [string]$cell.Address($false, $false)
                                Formula = $formula
                            })
                        }
                    }
                    finally {
                        Release-ComObject $cell
                    }
                }
            }
            finally {
                Release-ComObject $used
            }
        }
        finally {
            Release-ComObject $sheet
        }
    }
    return $items
}

function Restore-FormulaState {
    param([object[]]$Items)
    foreach ($item in $Items) {
        $sheet = $script:Workbook.Worksheets.Item($item.Sheet)
        $cell = $sheet.Range($item.Address)
        try {
            try { $cell.Formula2 = $item.Formula } catch { $cell.Formula = $item.Formula }
        }
        finally {
            Release-ComObject $cell
            Release-ComObject $sheet
        }
    }
}

function Save-NameState {
    $items = @{}
    foreach ($name in $script:Workbook.Names) {
        try {
            $plain = ([string]$name.Name).Split("!")[-1]
            if ($plain.StartsWith("PCS_")) {
                $items[$plain] = [string]$name.RefersTo
            }
        }
        finally {
            Release-ComObject $name
        }
    }
    return $items
}

function Restore-NameState {
    param([hashtable]$Items)
    foreach ($name in $Items.Keys) {
        try {
            $existing = $script:Workbook.Names.Item($name)
            $existing.RefersTo = $Items[$name]
            Release-ComObject $existing
        }
        catch {
            [void]$script:Workbook.Names.Add($name, $Items[$name])
        }
    }
}

function Remove-StarterTable {
    param([string]$SheetName, [string]$TableName, [int]$LastColumn)
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
        $lastRow = [Math]::Max(5, [int]$sheet.UsedRange.Rows.Count)
        $first = $sheet.Cells.Item(4, 1)
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
        [string]$TableName
    )
    [void]$script:Workbook.Queries.Add($QueryName, $Formula)
    $sheet = $script:Workbook.Worksheets.Item($SheetName)
    $destination = $sheet.Range("A4")
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
        foreach ($path in @($DataQueryPath, $CoachingQueryPath)) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Power Query definition not found: $path"
            }
        }
        $formulaState = Save-FormulaState
        $nameState = Save-NameState
        Remove-WorkbookQuery "PCS_DATA"
        Remove-WorkbookQuery "COACHING_QUEUE"
        Remove-StarterTable "PCS_DATA" "tblPcsData" 26
        Remove-StarterTable "COACHING_QUEUE" "tblCoachingQueue" 13
        Set-SetupValue "Connection Mode" $Mode
        Set-SetupValue "Local Feed Folder" ([System.IO.Path]::GetFullPath($FeedFolder))
        Add-QueryTable "PCS_DATA" ([System.IO.File]::ReadAllText($DataQueryPath)) "PCS_DATA" "tblPcsData"
        Add-QueryTable "COACHING_QUEUE" ([System.IO.File]::ReadAllText($CoachingQueryPath)) "COACHING_QUEUE" "tblCoachingQueue"
        Restore-FormulaState $formulaState
        Restore-NameState $nameState
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
