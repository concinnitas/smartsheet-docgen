# Run from this folder in PowerShell:  .\setup.ps1
# If blocked: Set-ExecutionPolicy -Scope Process Bypass
$ErrorActionPreference = "Stop"
if (-not (Test-Path .venv)) { py -3 -m venv .venv }
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
if (-not (Test-Path config\.env)) { Copy-Item config\.env.example config\.env; Write-Host "Created config\.env - add your Smartsheet token." }
.\.venv\Scripts\python.exe -m pytest -q tests
