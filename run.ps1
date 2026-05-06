#Requires -Version 5
<#
.SYNOPSIS
  Launch the Hermes WeChat Sidecar.
.DESCRIPTION
  Activates .venv and runs `python -m src.main`.
  Pre-flight: warns if .env is missing or still contains REPLACE_ME.
#>

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    Write-Error ".venv not found. Run:`n  py -3.11 -m venv .venv`n  .\.venv\Scripts\Activate.ps1`n  pip install -r requirements.txt"
    exit 1
}

if (-not (Test-Path ".\.env")) {
    Write-Error ".env not found. Copy .env.example to .env and fill in ADAPTER_AUTH_TOKEN."
    exit 1
}

$envContent = Get-Content .\.env -Raw
if ($envContent -match "REPLACE_ME") {
    Write-Warning ".env still contains REPLACE_ME placeholders. Sidecar will likely fail."
}

& .\.venv\Scripts\python.exe -m src.main
