param(
    [string]$Python = "py",
    [string]$PythonVersion = "3.14",
    [string]$OutDir = "..\dll"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resolvedOutDir = Join-Path $scriptDir $OutDir
$pythonSelector = if ($Python -ieq "py") { "-$PythonVersion" } else { $null }

if ($pythonSelector) {
    & $Python $pythonSelector -m pip install --upgrade pip
    & $Python $pythonSelector -m pip install `
        "google-genai" `
        "google-api-core" `
        "onnxruntime>=1.20,<2" `
        "numpy>=2,<3" `
        "pyinstaller>=6.10"
} else {
    & $Python -m pip install --upgrade pip
    & $Python -m pip install `
        "google-genai" `
        "google-api-core" `
        "onnxruntime>=1.20,<2" `
        "numpy>=2,<3" `
        "pyinstaller>=6.10"
}
if ($LASTEXITCODE -ne 0) {
    throw "Installing audio-description worker dependencies failed with exit code $LASTEXITCODE."
}

Push-Location $scriptDir
try {
    if ($pythonSelector) {
        & $Python $pythonSelector -m PyInstaller --noconfirm --clean audio_description_bridge.spec
    } else {
        & $Python -m PyInstaller --noconfirm --clean audio_description_bridge.spec
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Building audio_description_bridge.exe failed with exit code $LASTEXITCODE."
    }
    New-Item -ItemType Directory -Force -Path $resolvedOutDir | Out-Null
    Copy-Item -Force ".\dist\audio_description_bridge.exe" (Join-Path $resolvedOutDir "audio_description_bridge.exe")
} finally {
    Pop-Location
}

Write-Host "Audio-description bridge built in $resolvedOutDir\audio_description_bridge.exe"
