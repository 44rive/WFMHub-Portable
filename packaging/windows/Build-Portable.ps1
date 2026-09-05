[CmdletBinding()]
param(
    [string]$Version = "0.16.0",
    [string]$PythonVersion = "3.13.7"
)

$ErrorActionPreference = "Stop"
$Repository = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
Push-Location $Repository
try {
    py -3.13 packaging\windows\build_portable.py --clean --version $Version --python-version $PythonVersion
    if ($LASTEXITCODE -ne 0) { throw "Portable build failed with exit code $LASTEXITCODE" }
}
finally {
    Pop-Location
}
