$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectDir

python -m pip install --upgrade pyinstaller

pyinstaller `
  --name "NotaFacil" `
  --onefile `
  --noconsole `
  --icon "static_v3\icons\app.ico" `
  --add-data "templates_v3;templates_v3" `
  --add-data "static_v3;static_v3" `
  "app_pwa.py"

Write-Host ""
Write-Host "EXE gerado em: $ProjectDir\dist\NotaFacil.exe"
