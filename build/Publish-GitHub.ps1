param(
    [Parameter(Mandatory = $true)]
    [string]$Owner,

    [Parameter(Mandatory = $true)]
    [string]$Repo,

    [string]$Token = $env:GITHUB_TOKEN,
    [string]$Version = "",
    [string]$Description = "Tray-first Windows overlay app for window mirroring, text overlays, and embedded web pages.",
    [switch]$Private
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy -Scope Process Bypass -Force

if ([string]::IsNullOrWhiteSpace($Token)) {
    throw "GitHub token is required. Pass -Token or set GITHUB_TOKEN."
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$versionFile = Join-Path $repoRoot "VERSION"
$releaseVersion = if ([string]::IsNullOrWhiteSpace($Version)) {
    ([string](Get-Content $versionFile -Raw)).Trim()
}
else {
    $Version.Trim()
}

$releaseRoot = Join-Path $repoRoot "artifacts\release"
$portableZip = Join-Path $releaseRoot ("OverlayMirror-win64-{0}.zip" -f $releaseVersion)
$sourceZip = Join-Path $releaseRoot ("OverlayMirror-source-{0}.zip" -f $releaseVersion)
$portableExe = Join-Path $releaseRoot "OverlayMirror-win64\OverlayMirror.exe"

foreach ($requiredPath in @($portableZip, $sourceZip, $portableExe)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Missing release artifact: $requiredPath"
    }
}

$headers = @{
    Authorization = "Bearer $Token"
    Accept = "application/vnd.github+json"
    "User-Agent" = "OverlayMirrorPublisher"
}

$repoApi = "https://api.github.com/repos/$Owner/$Repo"
$repoExists = $true

try {
    Invoke-RestMethod -Uri $repoApi -Headers $headers -Method Get | Out-Null
}
catch {
    $repoExists = $false
}

if (-not $repoExists) {
    $createBody = @{
        name = $Repo
        description = $Description
        private = [bool]$Private
        has_issues = $true
        has_projects = $false
        has_wiki = $false
    } | ConvertTo-Json

    if ((Invoke-RestMethod -Uri "https://api.github.com/user" -Headers $headers -Method Get).login -eq $Owner) {
        Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $headers -Method Post -Body $createBody | Out-Null
    }
    else {
        Invoke-RestMethod -Uri "https://api.github.com/orgs/$Owner/repos" -Headers $headers -Method Post -Body $createBody | Out-Null
    }
}

$originUrl = "https://github.com/$Owner/$Repo.git"

$remoteNames = git -C $repoRoot remote
if ($remoteNames -notcontains "origin") {
    git -C $repoRoot remote add origin $originUrl
}
else {
    git -C $repoRoot remote set-url origin $originUrl
}

$existingUserName = git -C $repoRoot config --get user.name 2>$null
if ([string]::IsNullOrWhiteSpace($existingUserName)) {
    git -C $repoRoot config user.name $Owner
}

$existingUserEmail = git -C $repoRoot config --get user.email 2>$null
if ([string]::IsNullOrWhiteSpace($existingUserEmail)) {
    git -C $repoRoot config user.email "$Owner@users.noreply.github.com"
}

git -C $repoRoot add .
if (git -C $repoRoot diff --cached --quiet) {
}
else {
    git -C $repoRoot commit -m ("Release {0}" -f $releaseVersion)
}

$encodedToken = [Uri]::EscapeDataString($Token)
$pushUrl = "https://$encodedToken@github.com/$Owner/$Repo.git"
git -C $repoRoot push -u $pushUrl master

$tag = "v$releaseVersion"
$releaseName = "Overlay Mirror $releaseVersion"

$releaseLookup = $null
try {
    $releaseLookup = Invoke-RestMethod -Uri "$repoApi/releases/tags/$tag" -Headers $headers -Method Get
}
catch {
    $releaseLookup = $null
}

if ($null -eq $releaseLookup) {
    $releaseBody = @{
        tag_name = $tag
        target_commitish = "master"
        name = $releaseName
        draft = $false
        prerelease = $false
        generate_release_notes = $true
    } | ConvertTo-Json

    $releaseLookup = Invoke-RestMethod -Uri "$repoApi/releases" -Headers $headers -Method Post -Body $releaseBody
}

$uploadUrlTemplate = $releaseLookup.upload_url -replace '\{\?name,label\}$', ''

foreach ($assetPath in @($portableZip, $sourceZip, $portableExe)) {
    $assetName = [System.IO.Path]::GetFileName($assetPath)

    $existingAsset = $releaseLookup.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
    if ($null -ne $existingAsset) {
        Invoke-RestMethod -Uri "$repoApi/releases/assets/$($existingAsset.id)" -Headers $headers -Method Delete | Out-Null
    }

    Invoke-RestMethod `
        -Uri ($uploadUrlTemplate + "?name=" + [Uri]::EscapeDataString($assetName)) `
        -Headers @{
            Authorization = "Bearer $Token"
            Accept = "application/vnd.github+json"
            "User-Agent" = "OverlayMirrorPublisher"
            "Content-Type" = "application/octet-stream"
        } `
        -Method Post `
        -InFile $assetPath | Out-Null
}

Write-Output ("Repository: https://github.com/{0}/{1}" -f $Owner, $Repo)
Write-Output ("Release   : https://github.com/{0}/{1}/releases/tag/{2}" -f $Owner, $Repo, $tag)
