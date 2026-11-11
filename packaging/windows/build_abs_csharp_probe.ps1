param([switch]$Test)
$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$SourceRoot = Join-Path $RepoRoot 'miniapps/abs_csharp_probe'
$BuildRoot = Join-Path $RepoRoot 'build/abs-csharp-probe'
$Stage = Join-Path $BuildRoot 'WFMHub-ABS-CSharp-Probe'
$Collector = Join-Path $Stage '_system/collector'
if (Test-Path $BuildRoot) { Remove-Item -LiteralPath $BuildRoot -Recurse -Force }
New-Item -ItemType Directory -Path $Collector -Force | Out-Null
$ReferenceRoot = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8'
if (-not (Test-Path $ReferenceRoot)) { throw 'The Windows build machine requires the .NET Framework 4.8 targeting pack.' }
$Compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (-not (Test-Path $Compiler)) { throw 'The Windows build machine is missing the .NET Framework C# compiler.' }
function Find-Reference([string]$Name) {
    $Item = Get-ChildItem -LiteralPath $ReferenceRoot -Filter $Name -Recurse | Select-Object -First 1
    if ($null -eq $Item) { throw "Missing Framework build reference: $Name" }
    return $Item.FullName
}
$References = @('mscorlib.dll','System.dll','System.Core.dll','System.Web.Extensions.dll','UIAutomationClient.dll','UIAutomationTypes.dll','WindowsBase.dll') | ForEach-Object { '/reference:' + (Find-Reference $_) }
$Executable = Join-Path $Collector 'ABSProbe.exe'
& $Compiler /nologo /noconfig /nostdlib+ /target:exe /platform:x64 /optimize+ /warnaserror+ "/out:$Executable" @References (Join-Path $SourceRoot 'ABSProbe.cs')
if ($LASTEXITCODE -ne 0) { throw 'C# diagnostic compilation failed.' }
Copy-Item (Join-Path $SourceRoot 'ABSProbe.exe.config') $Collector
if ($Test) {
    $Tests = Join-Path $BuildRoot 'NativeTests.exe'
    $TestReferences = @(('/reference:' + (Find-Reference 'System.Windows.Forms.dll')), ('/reference:' + (Find-Reference 'System.Drawing.dll')))
    & $Compiler /nologo /noconfig /nostdlib+ /target:exe /platform:x64 /optimize+ /warnaserror+ /main:WfmHubAbsProbe.NativeTests "/out:$Tests" @References @TestReferences (Join-Path $SourceRoot 'ABSProbe.cs') (Join-Path $SourceRoot 'NativeTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Native test compilation failed.' }
    Copy-Item (Join-Path $SourceRoot 'ABSProbe.exe.config') ($Tests + '.config')
    & $Tests
    if ($LASTEXITCODE -ne 0) { throw 'Native Windows tests failed.' }
}
foreach ($Name in @('CHECK_LAUNCH.cmd','PROBE_ABS.cmd','README.md')) {
    if ($Name.EndsWith('.cmd')) {
        $Text = [IO.File]::ReadAllText((Join-Path $SourceRoot $Name)).Replace("`r`n","`n").Replace("`n","`r`n")
        [IO.File]::WriteAllText((Join-Path $Stage $Name), $Text, [Text.UTF8Encoding]::new($false))
    } else { Copy-Item (Join-Path $SourceRoot $Name) $Stage }
}
$Code = New-Item -ItemType Directory -Path (Join-Path $Stage '_system/source') -Force
Copy-Item (Join-Path $SourceRoot 'ABSProbe.cs') $Code.FullName
[IO.File]::WriteAllText((Join-Path $Stage 'VERSION.txt'), "abs-csharp-probe-0.1.0`n")
$Manifest = Get-ChildItem -LiteralPath $Stage -File -Recurse | Sort-Object FullName | ForEach-Object {
    $Relative = $_.FullName.Substring($Stage.Length + 1).Replace([char]92,[char]47)
    (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $Relative
}
[IO.File]::WriteAllLines((Join-Path $Stage 'SHA256SUMS.txt'), $Manifest, [Text.UTF8Encoding]::new($false))
$Dist = Join-Path $RepoRoot 'dist'
New-Item -ItemType Directory -Path $Dist -Force | Out-Null
$Archive = Join-Path $Dist 'WFMHub-ABS-CSharp-Probe-v0.1.0-win-x64.zip'
if (Test-Path $Archive) { Remove-Item -LiteralPath $Archive }
Compress-Archive -LiteralPath $Stage -DestinationPath $Archive -CompressionLevel Optimal
$Digest = (Get-FileHash -LiteralPath $Archive -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText(($Archive + '.sha256'), $Digest + '  ' + [IO.Path]::GetFileName($Archive) + "`n", [Text.UTF8Encoding]::new($false))
Write-Output "Built $Archive"
