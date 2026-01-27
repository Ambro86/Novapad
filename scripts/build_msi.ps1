$ErrorActionPreference = "Stop"

$repo = Resolve-Path (Join-Path $PSScriptRoot "..")
Push-Location $repo
try {
    $wixDir = Join-Path $repo "target\\release\\.cargo-packager\\wix\\x64"
    $lightExe = Join-Path $env:LOCALAPPDATA ".cargo-packager\\WixTools\\light.exe"
    $locFile = Join-Path $wixDir "locale.wxl"
    $wixobj = Join-Path $wixDir "main.wixobj"
    $outTemp = Join-Path $wixDir "output.msi"

    if (-not (Test-Path $lightExe)) {
        throw "WixTools not found: $lightExe"
    }

    if (-not (Test-Path $wixobj)) {
        throw "Missing wixobj. Run: cargo packager --release --format wix"
    }

    & $lightExe -sval -ext WixUIExtension -ext WixUtilExtension -cultures:en-us -loc $locFile -out $outTemp $wixobj

    $versionLine = Select-String -Path (Join-Path $repo "Cargo.toml") -Pattern '^\s*version\s*=\s*"(.+)"'
    if (-not $versionLine) {
        throw "Version not found in Cargo.toml"
    }
    $version = $versionLine.Matches[0].Groups[1].Value
    $msiName = "novapad_${version}_x64_en-US.msi"
    $dest = Join-Path $repo "target\\release\\$msiName"
    Move-Item $outTemp -Destination $dest -Force
} finally {
    Pop-Location
}
