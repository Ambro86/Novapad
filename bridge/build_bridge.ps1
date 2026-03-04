param(
    [string]$Python = "py",
    [string]$OutDir = "..\\dll"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$bridgeScript = Join-Path $scriptDir "faster_whisper_bridge.py"
$resolvedOutDir = Join-Path $scriptDir $OutDir

& $Python -3 -m pip install --upgrade pip
& $Python -3 -m pip install faster-whisper pyinstaller

& $Python -3 -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name faster_whisper_bridge `
    --distpath $resolvedOutDir `
    $bridgeScript

Write-Host "Bridge built in $resolvedOutDir\\faster_whisper_bridge.exe"
