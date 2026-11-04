param([switch]$OneFile)

$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
$previousMode = $env:CSVINTERPRETER_ONEFILE
try {
    # No system packages: optional libraries installed elsewhere must not leak
    # into the executable through OpenPyXL's optional imports.
    if (-not (Test-Path -LiteralPath '.venv-build/Scripts/python.exe')) {
        python -m venv .venv-build
        if ($LASTEXITCODE -ne 0) { throw 'Failed to create the build environment.' }
    }
    $buildPython = Join-Path $PSScriptRoot '.venv-build/Scripts/python.exe'
    & $buildPython -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'Failed to install build dependencies.' }

    $env:CSVINTERPRETER_ONEFILE = if ($OneFile) { '1' } else { '0' }
    & $buildPython -m PyInstaller --noconfirm --clean CSVInterpreter.spec
    if ($LASTEXITCODE -ne 0) { throw 'Failed to build the application.' }
} finally {
    $env:CSVINTERPRETER_ONEFILE = $previousMode
    Pop-Location
}
