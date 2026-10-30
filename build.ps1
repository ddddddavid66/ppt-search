param([switch]$InstallDependencies)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$taskPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $taskPython)) {
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw 'Cannot create project virtual environment.' }
    $InstallDependencies = $true
}
if ($InstallDependencies) {
    & $taskPython -m pip install -r requirements-lock.txt
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
}
& $taskPython -m tools.build
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
