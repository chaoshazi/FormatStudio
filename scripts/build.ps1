$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Host "==> Building frontend"
Push-Location web
npm install
npm run build
Pop-Location

Write-Host "==> Installing build dependencies"
.\.venv\Scripts\python.exe -m pip install pyinstaller

if (-not (Test-Path "vendor\libreoffice\program\soffice.exe")) {
    Write-Warning "vendor\libreoffice\program\soffice.exe not found. The installer will not include LibreOffice."
}

Write-Host "==> Running PyInstaller"
.\.venv\Scripts\pyinstaller.exe --noconfirm --clean format-studio.spec

$iscc = Get-Command iscc.exe -ErrorAction SilentlyContinue
if ($iscc) {
    Write-Host "==> Building installer"
    & $iscc.Source "installer\format-studio.iss"
} else {
    Write-Warning "Inno Setup (iscc.exe) not found. Skipping installer."
}

Write-Host "Done. Output in dist\"