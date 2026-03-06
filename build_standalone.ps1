param(
    [string]$PythonExe = "python",
    [string]$VenvPath = ".venv"
)

$ErrorActionPreference = "Stop"

Write-Host "Creating virtual environment..."
& $PythonExe -m venv $VenvPath

$VenvPython = Join-Path $VenvPath "Scripts\python.exe"
$VenvPip = Join-Path $VenvPath "Scripts\pip.exe"

Write-Host "Installing dependencies..."
& $VenvPython -m pip install --upgrade pip
& $VenvPip install -r requirements.txt pyinstaller

Write-Host "Building one-file Windows EXE..."
& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name FreeSpeech `
    --icon "assets\icon-512.ico" `
    --add-data "themes\red.json;themes" `
    --add-data "assets\icon-512.png;assets" `
    --add-data "assets\icon-512-maskable.png;assets" `
    --add-data "assets\icon-512.ico;assets" `
    --add-data "assets\favicon.ico;assets" `
    --add-data "assets\easteregg.mp3;assets" `
    --add-data "assets\easteregg2.mp3;assets" `
    --add-data "assets\easteregg2.ogg;assets" `
    freespeech\main.py

Write-Host ""
Write-Host "Build complete:"
Write-Host "  dist\FreeSpeech.exe"
