$ErrorActionPreference = 'Stop'

$venvPath = Join-Path $PSScriptRoot '.venv'
$python = Join-Path $venvPath 'Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPath)) {
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        throw 'Python not found. Make sure Python 3.10 or newer is installed and added to PATH.'
    }

    & python -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
    if ($LASTEXITCODE -ne 0) {
        throw 'Python must be version 3.10 or newer.'
    }

    & python -m venv $venvPath
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to create the virtual environment.'
    }
}

if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw "Virtual environment exists but its Python executable is missing: $python"
}

& $python -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
if ($LASTEXITCODE -ne 0) {
    throw 'The virtual environment requires Python 3.10 or newer.'
}

& $python -m pip install -e $PSScriptRoot
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to install the pipeline and its dependencies.'
}

Write-Host 'Setup complete. Run the pipeline with .\.venv\Scripts\pf-build-all-transactions.exe.'
