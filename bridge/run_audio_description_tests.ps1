param(
    [string]$Python = "py",
    [string]$PythonVersion = "3.14",
    [switch]$SkipRust
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$pythonSelector = if ($Python -ieq "py") { "-$PythonVersion" } else { $null }

if ($pythonSelector) {
    & $Python $pythonSelector -m pip install -r (Join-Path $scriptDir "requirements-test.txt")
    & $Python $pythonSelector (Join-Path $scriptDir "run_audio_description_tests.py")
} else {
    & $Python -m pip install -r (Join-Path $scriptDir "requirements-test.txt")
    & $Python (Join-Path $scriptDir "run_audio_description_tests.py")
}
if ($LASTEXITCODE -ne 0) {
    throw "The Python audio-description tests failed."
}

if (-not $SkipRust) {
    Push-Location $projectDir
    try {
        cargo test omni_port_ -- --nocapture
        if ($LASTEXITCODE -ne 0) {
            throw "The Rust audio-description regression tests failed."
        }
    } finally {
        Pop-Location
    }
}
