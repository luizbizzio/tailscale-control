# License: PolyForm Internal Use License 1.0.0
# Copyright (c) 2026 Luiz Bizzio

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'TailscaleControl'),
    [string]$SourceDirectory,
    [string]$ReleaseTag = 'latest',
    [string]$ReleaseAssetBase,
    [switch]$PreferRemote,
    [switch]$SkipHashCheck,
    [switch]$NoLaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AppName = 'Tailscale Control'
$script:AppVersion = '1.0.0'
$script:PowerShellExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1)
if ([string]::IsNullOrWhiteSpace([string]$script:PowerShellExe)) {
    $script:PowerShellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
}
$script:WScriptExe = Join-Path $env:WINDIR 'System32\wscript.exe'
$script:InstalledScriptPath = Join-Path $InstallRoot 'tailscale-control.ps1'
$script:InstalledIconsDir = Join-Path $InstallRoot 'assets\icons'
$script:InstalledIconPath = Join-Path $script:InstalledIconsDir 'tailscale-control.ico'
$script:LauncherVbsPath = Join-Path $InstallRoot 'TailscaleControlLauncher.vbs'
$script:StartMenuShortcutPath = Join-Path ([Environment]::GetFolderPath('Programs')) 'Tailscale Control.lnk'
$script:GitHubReleasesBase = 'https://github.com/luizbizzio/tailscale-control/releases'
$script:ExpectedScriptSha256 = '97ea5df2ef1a2089d1e50804112f53597e861b95e69a5d12ff6a613e2ecc71c2'
$script:IconAssets = @(
    [pscustomobject]@{
        FileName = 'tailscale-control.ico'
        RelativePath = 'assets/icons/tailscale-control.ico'
        ExpectedSha256 = '2c2c83a52aafd6def4c074fd3127d7de9f414453e8b92db30b4d1d11c1ea0e3a'
    },
    [pscustomobject]@{
        FileName = 'tailscale.ico'
        RelativePath = 'assets/icons/tailscale.ico'
        ExpectedSha256 = '1eff7d0ee72515e6bfd9a3301900c23a288b389e202c55a51fdd735243c411f0'
    },
    [pscustomobject]@{
        FileName = 'tailscale-mtu.ico'
        RelativePath = 'assets/icons/tailscale-mtu.ico'
        ExpectedSha256 = '01e024c88d947c4f40e9659498a2ae6ed10ed970a9130806000e2ffd17fcf0b5'
    }
)

function Write-Step {
    param([string]$Message)
    Write-Host ('[{0}] {1}' -f (Get-Date).ToString('HH:mm:ss'), $Message)
}

function Initialize-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-FileSha256Hex {
    param([string]$Path)
    return ([string](Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash).ToLowerInvariant()
}

function Assert-Hash {
    param(
        [string]$Path,
        [string]$Expected,
        [string]$Label
    )
    if ($SkipHashCheck) {
        Write-Step ('Skipping SHA-256 verification for ' + $Label + '.')
        return
    }
    if ([string]::IsNullOrWhiteSpace([string]$Expected)) {
        throw ('Missing expected SHA-256 for ' + $Label + '.')
    }
    $actual = Get-FileSha256Hex -Path $Path
    if ($actual -ne $Expected.ToLowerInvariant()) {
        throw ($Label + ' SHA-256 mismatch. Expected ' + $Expected + ' but got ' + $actual + '.')
    }
}

function Assert-IcoFile {
    param(
        [string]$Path,
        [string]$Label
    )
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 6) {
        throw ($Label + ' is too small to be a valid icon file.')
    }
    if ($bytes[0] -ne 0 -or $bytes[1] -ne 0 -or $bytes[2] -ne 1 -or $bytes[3] -ne 0) {
        throw ($Label + ' does not look like a valid .ico file.')
    }
}

function Resolve-ReleaseAssetBase {
    param(
        [string]$Tag,
        [string]$ExplicitBase
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$ExplicitBase)) {
        return ([string]$ExplicitBase).TrimEnd('/')
    }
    $tagValue = [string]$Tag
    if ([string]::IsNullOrWhiteSpace($tagValue)) { $tagValue = 'latest' }
    if ($tagValue -eq 'latest') {
        return ($script:GitHubReleasesBase + '/latest/download')
    }
    return ($script:GitHubReleasesBase + '/download/' + [uri]::EscapeDataString($tagValue))
}

$script:ResolvedReleaseAssetBase = Resolve-ReleaseAssetBase -Tag $ReleaseTag -ExplicitBase $ReleaseAssetBase
$script:RemoteScriptUrl = ($script:ResolvedReleaseAssetBase.TrimEnd('/') + '/tailscale-control.ps1')

function Invoke-Download {
    param(
        [string]$Url,
        [string]$OutFile,
        [string]$ExpectedSha256,
        [string]$Label,
        [switch]$Icon
    )
    Write-Step ('Downloading ' + $Label + '...')
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
    if ((Get-Item -LiteralPath $OutFile).Length -le 0) {
        throw ($Label + ' download returned an empty file.')
    }
    if ($Icon) {
        Assert-IcoFile -Path $OutFile -Label $Label
    }
    elseif ($Label -eq 'script') {
        $raw = Get-Content -LiteralPath $OutFile -Raw -Encoding UTF8
        if ($raw -match '(?i)<html' -or $raw -match '(?i)<!DOCTYPE html') {
            throw 'Downloaded script looks like HTML instead of PowerShell.'
        }
        if ($raw -notmatch '\$script:AppName\s*=\s*[''\"]Tailscale Control[''\"]') {
            throw 'Downloaded script did not pass the content sanity check.'
        }
    }
    Assert-Hash -Path $OutFile -Expected $ExpectedSha256 -Label $Label
}

function Resolve-SourceDirectory {
    if (-not [string]::IsNullOrWhiteSpace([string]$SourceDirectory)) {
        return $SourceDirectory
    }
    if ($PSCommandPath) {
        return (Split-Path -Parent $PSCommandPath)
    }
    return $null
}

function Resolve-LocalIconPath {
    param(
        [string]$BaseDirectory,
        [string]$FileName
    )
    if ([string]::IsNullOrWhiteSpace([string]$BaseDirectory)) { return $null }
    $assetPath = Join-Path (Join-Path $BaseDirectory 'assets\icons') $FileName
    if (Test-Path -LiteralPath $assetPath) { return $assetPath }
    $legacyPath = Join-Path $BaseDirectory $FileName
    if (Test-Path -LiteralPath $legacyPath) { return $legacyPath }
    return $null
}

function Get-AppVersionFromFile {
    param([string]$Path)
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ($raw -match '\$script:AppVersion\s*=\s*[''\"]([^''\"]+)[''\"]') { return [string]$Matches[1] }
        if ($raw -match '\$AppVersion\s*=\s*[''\"]([^''\"]+)[''\"]') { return [string]$Matches[1] }
    }
    catch { }
    return 'unknown'
}

function Write-AppLauncherVbs {
    param([string]$ScriptPath)
    $content = @"
Set oShell = CreateObject("WScript.Shell")
cmd = Chr(34) & "$($script:PowerShellExe)" & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & "$ScriptPath" & Chr(34)
oShell.Run cmd, 0, False
"@
    Set-Content -LiteralPath $script:LauncherVbsPath -Value $content -Encoding ASCII
}

function Initialize-StartMenuShortcut {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($script:StartMenuShortcutPath)
    $shortcut.TargetPath = $script:WScriptExe
    $shortcut.Arguments = '"' + $script:LauncherVbsPath + '"'
    $shortcut.WorkingDirectory = $InstallRoot
    if (Test-Path -LiteralPath $script:InstalledIconPath) {
        $shortcut.IconLocation = $script:InstalledIconPath
    }
    else {
        $shortcut.IconLocation = "$env:SystemRoot\System32\SHELL32.dll,220"
    }
    $shortcut.Description = 'Open Tailscale Control'
    $shortcut.Save()
}

function Install-FileAtomically {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )
    $destinationDirectory = Split-Path -Parent $DestinationPath
    Initialize-Directory -Path $destinationDirectory
    $tmpPath = $DestinationPath + '.tmp'
    if (Test-Path -LiteralPath $tmpPath) {
        Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $SourcePath -Destination $tmpPath -Force
    if (Test-Path -LiteralPath $DestinationPath) {
        Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue
    }
    Move-Item -LiteralPath $tmpPath -Destination $DestinationPath -Force
}

$resolvedSourceDir = Resolve-SourceDirectory
$localScriptPath = if ($resolvedSourceDir) { Join-Path $resolvedSourceDir 'tailscale-control.ps1' } else { $null }
$useLocalScript = $false
if (-not $PreferRemote) {
    if (-not [string]::IsNullOrWhiteSpace([string]$localScriptPath) -and (Test-Path -LiteralPath $localScriptPath)) { $useLocalScript = $true }
}

Initialize-Directory -Path $InstallRoot
Initialize-Directory -Path $script:InstalledIconsDir
$tempRoot = Join-Path $InstallRoot '.install-tmp'
Initialize-Directory -Path $tempRoot
$tempIconsDir = Join-Path $tempRoot 'assets\icons'
Initialize-Directory -Path $tempIconsDir
$tempScriptPath = Join-Path $tempRoot 'tailscale-control.ps1'

if ($useLocalScript) {
    Write-Step ('Using local script from ' + $localScriptPath)
    Copy-Item -LiteralPath $localScriptPath -Destination $tempScriptPath -Force
}
else {
    Invoke-Download -Url $script:RemoteScriptUrl -OutFile $tempScriptPath -ExpectedSha256 $script:ExpectedScriptSha256 -Label 'script'
}

foreach ($asset in $script:IconAssets) {
    $tempIconPath = Join-Path $tempIconsDir $asset.FileName
    $localIconPath = Resolve-LocalIconPath -BaseDirectory $resolvedSourceDir -FileName $asset.FileName
    if (-not $PreferRemote -and -not [string]::IsNullOrWhiteSpace([string]$localIconPath)) {
        Write-Step ('Using local icon from ' + $localIconPath)
        Copy-Item -LiteralPath $localIconPath -Destination $tempIconPath -Force
        Assert-IcoFile -Path $tempIconPath -Label $asset.FileName
    }
    else {
        $iconUrl = ($script:ResolvedReleaseAssetBase.TrimEnd('/') + '/' + $asset.FileName)
        Invoke-Download -Url $iconUrl -OutFile $tempIconPath -ExpectedSha256 $asset.ExpectedSha256 -Label $asset.FileName -Icon
    }
}

Install-FileAtomically -SourcePath $tempScriptPath -DestinationPath $script:InstalledScriptPath
foreach ($asset in $script:IconAssets) {
    $sourceIcon = Join-Path $tempIconsDir $asset.FileName
    $destinationIcon = Join-Path $script:InstalledIconsDir $asset.FileName
    Install-FileAtomically -SourcePath $sourceIcon -DestinationPath $destinationIcon
}
Write-AppLauncherVbs -ScriptPath $script:InstalledScriptPath
Initialize-StartMenuShortcut

try {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
catch { }

$version = Get-AppVersionFromFile -Path $script:InstalledScriptPath
if ([string]::IsNullOrWhiteSpace([string]$version)) { $version = [string]$script:AppVersion }
Write-Step ('Install/update completed. Version ' + $version + '.')
Write-Step ('Script path: ' + $script:InstalledScriptPath)
Write-Step ('Icons path: ' + $script:InstalledIconsDir)
Write-Step ('Remote asset base: ' + $script:ResolvedReleaseAssetBase)
Write-Step ('Launcher path: ' + $script:LauncherVbsPath)

if (-not $NoLaunch) {
    Write-Step 'Launching Tailscale Control without console...'
    Start-Process -FilePath $script:WScriptExe -ArgumentList @($script:LauncherVbsPath)
}
