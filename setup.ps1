$ErrorActionPreference = 'Stop'

$venvPath = Join-Path $PSScriptRoot '.venv'
$python = Join-Path $venvPath 'Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPath)) {
    if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
        throw 'Python launcher not found. Install Python 3.10 or newer and the Windows Python launcher.'
    }

    & py -3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
    if ($LASTEXITCODE -ne 0) {
        throw 'The Python launcher must select Python 3.10 or newer.'
    }

    & py -3 -m venv $venvPath
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
