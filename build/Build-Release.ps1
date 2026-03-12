param(
    [string]$Version = "",
    [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process Bypass -Force

function Get-VersionValue {
    param([string]$RepoRoot)

    if (-not [string]::IsNullOrWhiteSpace($Version)) {
        return $Version.Trim()
    }

    $versionFile = Join-Path $RepoRoot "VERSION"
    if (-not (Test-Path $versionFile)) {
        throw "VERSION file not found."
    }

    return ([string](Get-Content $versionFile -Raw)).Trim()
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Ensure-Ps2Exe {
    if (-not (Get-Module -ListAvailable ps2exe)) {
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module ps2exe
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$releaseVersion = Get-VersionValue -RepoRoot $repoRoot
$artifactRoot = Join-Path $repoRoot "artifacts"
$releaseRoot = Join-Path $artifactRoot "release"
$portableRoot = Join-Path $releaseRoot "OverlayMirror-win64"
$portableExe = Join-Path $portableRoot "OverlayMirror.exe"
$portableZip = Join-Path $releaseRoot ("OverlayMirror-win64-{0}.zip" -f $releaseVersion)
$sourceZip = Join-Path $releaseRoot ("OverlayMirror-source-{0}.zip" -f $releaseVersion)
$checksumsFile = Join-Path $releaseRoot "SHA256SUMS.txt"
$vendorRoot = Join-Path $repoRoot "vendor\webview2"

if ($Clean -and (Test-Path $artifactRoot)) {
    Remove-Item -Path $artifactRoot -Recurse -Force
}

Ensure-Directory -Path $releaseRoot
Ensure-Directory -Path $portableRoot

if (-not (Test-Path (Join-Path $vendorRoot "Microsoft.Web.WebView2.Core.dll"))) {
    throw "Missing vendor\\webview2\\Microsoft.Web.WebView2.Core.dll"
}

if (-not (Test-Path (Join-Path $vendorRoot "Microsoft.Web.WebView2.WinForms.dll"))) {
    throw "Missing vendor\\webview2\\Microsoft.Web.WebView2.WinForms.dll"
}

Ensure-Ps2Exe

Invoke-ps2exe `
    -inputFile (Join-Path $repoRoot "OverlayMirror.ps1") `
    -outputFile $portableExe `
    -noConsole `
    -STA `
    -DPIAware `
    -winFormsDPIAware `
    -supportOS `
    -longPaths `
    -title "Overlay Mirror" `
    -description "Tray-first overlay tool for window preview, text, and web pages." `
    -company "eshap" `
    -product "Overlay Mirror" `
    -copyright "Copyright (c) 2026 eshap" `
    -version $releaseVersion

Copy-Item (Join-Path $vendorRoot "Microsoft.Web.WebView2.Core.dll") $portableRoot -Force
Copy-Item (Join-Path $vendorRoot "Microsoft.Web.WebView2.WinForms.dll") $portableRoot -Force
Copy-Item (Join-Path $repoRoot "README.md") $portableRoot -Force
Copy-Item (Join-Path $repoRoot "LICENSE") $portableRoot -Force

if (Test-Path $portableZip) {
    Remove-Item $portableZip -Force
}

Compress-Archive -Path (Join-Path $portableRoot "*") -DestinationPath $portableZip -Force

if (Test-Path $sourceZip) {
    Remove-Item $sourceZip -Force
}

$sourceItems = Get-ChildItem -Path $repoRoot -Force | Where-Object {
    $_.Name -notin @(".git", "artifacts")
}
Compress-Archive -Path $sourceItems.FullName -DestinationPath $sourceZip -Force

$hashItems = @(
    (Get-FileHash -Path $portableExe -Algorithm SHA256)
    (Get-FileHash -Path $portableZip -Algorithm SHA256)
    (Get-FileHash -Path $sourceZip -Algorithm SHA256)
)

$hashEntries = $hashItems | ForEach-Object {
    "{0} *{1}" -f $_.Hash, [System.IO.Path]::GetFileName($_.Path)
}

Set-Content -Path $checksumsFile -Value $hashEntries -Encoding ASCII

Write-Output ("Built EXE      : {0}" -f $portableExe)
Write-Output ("Portable ZIP   : {0}" -f $portableZip)
Write-Output ("Source ZIP     : {0}" -f $sourceZip)
Write-Output ("Checksums file : {0}" -f $checksumsFile)
