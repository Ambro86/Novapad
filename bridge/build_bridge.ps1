param(
    [string]$Python = "py",
    [string]$PythonVersion = "3.14",
    [string]$OutDir = "..\\dll"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$bridgeScript = Join-Path $scriptDir "faster_whisper_bridge.py"
$resolvedOutDir = Join-Path $scriptDir $OutDir
$launcherArgs = if ($Python -ieq "py") { @("-$PythonVersion") } else { @() }

& $Python @launcherArgs -m pip install --upgrade pip
& $Python @launcherArgs -m pip install "faster-whisper==1.2.1" pyinstaller

& $Python @launcherArgs -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name faster_whisper_bridge `
    --distpath $resolvedOutDir `
    $bridgeScript

Write-Host "Bridge built in $resolvedOutDir\\faster_whisper_bridge.exe"
