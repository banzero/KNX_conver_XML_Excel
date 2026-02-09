$ErrorActionPreference = "Stop"

$RootDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $RootDir
$env:PYINSTALLER_CONFIG_DIR = "$RootDir\\.pyinstaller"
New-Item -ItemType Directory -Path $env:PYINSTALLER_CONFIG_DIR -Force | Out-Null

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    python -m venv .venv-build
    .\.venv-build\Scripts\Activate.ps1
    python -m pip install --upgrade pip
    python -m pip install pyinstaller
}

pyinstaller --clean --noconfirm --onefile --name knx-web-tool knx_web_tool.py

if (Test-Path package) { Remove-Item package -Recurse -Force }
New-Item -ItemType Directory -Path package | Out-Null
Copy-Item dist\knx-web-tool.exe package\
Copy-Item README_web_tool.md package\README_web_tool.md

if (Test-Path knx-web-tool-windows.zip) { Remove-Item knx-web-tool-windows.zip -Force }
Compress-Archive -Path package\* -DestinationPath knx-web-tool-windows.zip -Force

Write-Output "Created: $RootDir\knx-web-tool-windows.zip"
