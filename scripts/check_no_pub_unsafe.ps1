$ErrorActionPreference = "Stop"

function Get-RustFiles {
    param([string[]]$Roots)

    $files = @()
    foreach ($root in $Roots) {
        if (Test-Path $root) {
            $files += Get-ChildItem -Path $root -Recurse -File -Filter *.rs |
                Where-Object { $_.FullName -notmatch '\\target\\' -and $_.Name -notlike '*.bac' }
        }
    }
    return $files
}

$roots = @("src", "sapi4_bridge/src")
$files = Get-RustFiles -Roots $roots

if (-not $files -or $files.Count -eq 0) {
    Write-Host "[check_no_pub_unsafe] No Rust source files found in configured roots."
    exit 0
}

# Reject all public unsafe functions, including pub(crate)/pub(super)/pub(in ...).
$pattern = '\bpub(?:\s*\([^)]*\))?\s+unsafe\s+fn\b'
$violations = @()

foreach ($file in $files) {
    $content = [System.IO.File]::ReadAllText($file.FullName)
    $matches = [System.Text.RegularExpressions.Regex]::Matches($content, $pattern)
    if ($matches.Count -gt 0) {
        $lines = $content -split "`r?`n"
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -match $pattern) {
                $violations += "{0}:{1}: {2}" -f $file.FullName, ($i + 1), $lines[$i].Trim()
            }
        }
    }
}

if ($violations.Count -gt 0) {
    Write-Host "[check_no_pub_unsafe] ERROR: public unsafe functions are forbidden."
    $violations | ForEach-Object { Write-Host $_ }
    exit 1
}

Write-Host "[check_no_pub_unsafe] OK: no public unsafe functions found."
