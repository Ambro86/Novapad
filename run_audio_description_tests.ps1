param(
    [string]$Python = "py",
    [string]$PythonVersion = "3.14",
    [switch]$InstallPythonTestDependencies,
    [switch]$SkipRust
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Bridge = Join-Path $Root "bridge"

function Invoke-Python([string[]]$Arguments) {
    if ($Python -ieq "py") {
        & $Python "-$PythonVersion" @Arguments
    } else {
        & $Python @Arguments
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Python command failed with exit code $LASTEXITCODE."
    }
}

if ($InstallPythonTestDependencies) {
    Invoke-Python @("-m", "pip", "install", "-r", (Join-Path $Bridge "requirements-test.txt"))
}

Write-Host "Running 89 audio-description worker and bridge-protocol tests..."
Invoke-Python @((Join-Path $Bridge "run_audio_description_tests.py"))

if (-not $SkipRust) {
    Write-Host "Running 17 Sonarpad Rust audio-description regression tests..."
    Push-Location $Root
    try {
        cargo test omni_port_ -- --nocapture
        if ($LASTEXITCODE -ne 0) {
            throw "cargo test failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
}

Write-Host "Audio-description test run completed."
