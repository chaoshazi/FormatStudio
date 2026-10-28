$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    python -m venv .venv
}
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

if (-not (Test-Path "web\node_modules")) {
    Push-Location web
    npm install
    Pop-Location
}

Start-Process -WindowStyle Hidden -FilePath ".\.venv\Scripts\python.exe" -ArgumentList "run_server.py" -WorkingDirectory $root
Push-Location web
npm run dev
Pop-Location