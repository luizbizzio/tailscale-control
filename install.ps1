# License: PolyForm Internal Use License 1.0.0
# Copyright (c) 2026 Luiz Bizzio

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'TailscaleControl'),
    [string]$SourceDirectory,
    [string]$ReleaseTag,
    [string]$ReleaseAssetBase,
    [switch]$PreferRemote,
    [switch]$NoLaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$DefaultReleaseTag = 'v1.0.0'

if ([string]::IsNullOrWhiteSpace($DefaultReleaseTag) -or $DefaultReleaseTag -eq '__RELEASE_TAG__') {
    throw 'Default release tag was not filled. Run the GitHub Action before publishing install.ps1.'
}

function ConvertFrom-ReleaseTagToVersion {
    param([string]$Tag)

    $value = [string]$Tag
    if ([string]::IsNullOrWhiteSpace($value)) {
        return 'unknown'
    }

    $value = $value.Trim()

    if ($value -eq 'latest') {
        return 'latest'
    }

    if ($value.StartsWith('v', [System.StringComparison]::OrdinalIgnoreCase)) {
        $value = $value.Substring(1)
    }

    return $value
}

if ([string]::IsNullOrWhiteSpace([string]$ReleaseTag)) {
    $ReleaseTag = $DefaultReleaseTag
}

$script:AppName = 'Tailscale Control'
$script:AppVersion = ConvertFrom-ReleaseTagToVersion -Tag $DefaultReleaseTag
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
$script:IconAssets = @(
    [pscustomobject]@{
        FileName = 'tailscale-control.ico'
        RelativePath = 'assets/icons/tailscale-control.ico'
    },
    [pscustomobject]@{
        FileName = 'tailscale.ico'
        RelativePath = 'assets/icons/tailscale.ico'
    },
    [pscustomobject]@{
        FileName = 'tailscale-mtu.ico'
        RelativePath = 'assets/icons/tailscale-mtu.ico'
    }
)

function Write-Step {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Cyan
    )

    try {
        Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $Message" -ForegroundColor $Color
    } catch {
        Write-Output $Message
    }
}

function Initialize-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
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

function Assert-PowerShellScript {
    param(
        [string]$Path,
        [string]$Label
    )

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ($raw -match '(?i)<html' -or $raw -match '(?i)<!DOCTYPE html') {
        throw ($Label + ' looks like HTML instead of PowerShell.')
    }
    if ($raw -notmatch '\$script:AppName\s*=\s*[''"]Tailscale Control[''"]') {
        throw ($Label + ' did not pass the content sanity check.')
    }
}

function Resolve-ReleaseTag {
    param([string]$Tag)

    $tagValue = [string]$Tag
    if ([string]::IsNullOrWhiteSpace($tagValue)) {
        return $DefaultReleaseTag
    }
    $tagValue = $tagValue.Trim()
    if ($tagValue -eq 'latest') {
        return 'latest'
    }
    if ($tagValue -match '^\d+(\.\d+){1,3}([\-+].*)?$') {
        return ('v' + $tagValue)
    }
    return $tagValue
}

function Resolve-ReleaseAssetBase {
    param(
        [string]$Tag,
        [string]$ExplicitBase
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$ExplicitBase)) {
        return ([string]$ExplicitBase).TrimEnd('/')
    }

    $tagValue = Resolve-ReleaseTag -Tag $Tag
    if ($tagValue -eq 'latest') {
        return 'https://github.com/luizbizzio/tailscale-control/releases/latest/download'
    }
    return ('https://github.com/luizbizzio/tailscale-control/releases/download/' + [uri]::EscapeDataString($tagValue))
}

$script:ResolvedReleaseTag = Resolve-ReleaseTag -Tag $ReleaseTag
$script:ResolvedReleaseAssetBase = Resolve-ReleaseAssetBase -Tag $ReleaseTag -ExplicitBase $ReleaseAssetBase
$script:RemoteScriptUrl = ($script:ResolvedReleaseAssetBase.TrimEnd('/') + '/tailscale-control.ps1')

function Invoke-Download {
    param(
        [string]$Url,
        [string]$OutFile,
        [string]$Label,
        [switch]$Icon
    )

    Write-Step ('Downloading ' + $Label + '...')
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $OutFile)) {
        throw ($Label + ' was not downloaded.')
    }
    if ((Get-Item -LiteralPath $OutFile).Length -le 0) {
        throw ($Label + ' download returned an empty file.')
    }
    if ($Icon) {
        Assert-IcoFile -Path $OutFile -Label $Label
    } elseif ($Label -eq 'script') {
        Assert-PowerShellScript -Path $OutFile -Label $Label
    }
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
        if ($raw -match '\$script:AppVersion\s*=\s*[''"]([^''"]+)[''"]') { return [string]$Matches[1] }
        if ($raw -match '\$AppVersion\s*=\s*[''"]([^''"]+)[''"]') { return [string]$Matches[1] }
    }
    catch { }
    return 'unknown'
}

function Assert-AppLauncherVbs {
    param(
        [string]$Path,
        [string]$ScriptPath
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw 'Launcher VBS was not created.'
    }

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 80) {
        throw 'Launcher VBS is too small.'
    }

    $hasNonZeroByte = $false
    foreach ($byte in $bytes) {
        if ($byte -ne 0) {
            $hasNonZeroByte = $true
            break
        }
    }

    if (-not $hasNonZeroByte) {
        throw 'Launcher VBS is corrupted.'
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding ASCII
    if ($raw -notmatch 'WScript\.Shell' -or $raw -notmatch 'oShell\.Run') {
        throw 'Launcher VBS did not pass the content sanity check.'
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$ScriptPath)) {
        if ($raw -notmatch [regex]::Escape([string]$ScriptPath)) {
            throw 'Launcher VBS points to the wrong script path.'
        }
    }
}


function Move-FileAtomically {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        throw 'Atomic file source does not exist.'
    }

    $destinationDirectory = Split-Path -Parent $DestinationPath
    if (-not [string]::IsNullOrWhiteSpace([string]$destinationDirectory)) {
        if (-not (Test-Path -LiteralPath $destinationDirectory)) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
        }
    }

    $backupPath = $DestinationPath + '.bak'
    if (Test-Path -LiteralPath $backupPath) {
        Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $DestinationPath) {
        [System.IO.File]::Replace($SourcePath, $DestinationPath, $backupPath, $true)
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }
    else {
        [System.IO.File]::Move($SourcePath, $DestinationPath)
    }
}

function Write-AppLauncherVbs {
    param(
        [string]$ScriptPath,
        [string]$LauncherPath = $script:LauncherVbsPath,
        [switch]$BackgroundMode
    )

    if ([string]::IsNullOrWhiteSpace([string]$ScriptPath)) {
        throw 'Script path cannot be empty.'
    }

    if ([string]::IsNullOrWhiteSpace([string]$LauncherPath)) {
        throw 'Launcher path cannot be empty.'
    }

    $launcherDirectory = Split-Path -Parent $LauncherPath
    Initialize-Directory -Path $launcherDirectory

    $extra = ''
    if ($BackgroundMode) { $extra = ' -Background' }

    $content = @"
Set oShell = CreateObject("WScript.Shell")
cmd = Chr(34) & "$($script:PowerShellExe)" & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & "$ScriptPath" & Chr(34) & "$extra"
oShell.Run cmd, 0, False
"@

    $tmpPath = $LauncherPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        Set-Content -LiteralPath $tmpPath -Value $content -Encoding ASCII
        Assert-AppLauncherVbs -Path $tmpPath -ScriptPath $ScriptPath

        Move-FileAtomically -SourcePath $tmpPath -DestinationPath $LauncherPath
        Assert-AppLauncherVbs -Path $LauncherPath -ScriptPath $ScriptPath
    }
    finally {
        if (Test-Path -LiteralPath $tmpPath) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
    }
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
    Move-FileAtomically -SourcePath $tmpPath -DestinationPath $DestinationPath
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {
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

try {
    Write-Step ('Selected release: ' + $script:ResolvedReleaseTag)
    Write-Step ('Release asset base: ' + $script:ResolvedReleaseAssetBase)

    if ($useLocalScript) {
        Write-Step ('Using local script from ' + $localScriptPath)
        Copy-Item -LiteralPath $localScriptPath -Destination $tempScriptPath -Force
        Assert-PowerShellScript -Path $tempScriptPath -Label 'script'
    }
    else {
        Invoke-Download -Url $script:RemoteScriptUrl -OutFile $tempScriptPath -Label 'script'
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
            Invoke-Download -Url $iconUrl -OutFile $tempIconPath -Label $asset.FileName -Icon
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

    $version = Get-AppVersionFromFile -Path $script:InstalledScriptPath
    if ([string]::IsNullOrWhiteSpace([string]$version)) { $version = [string]$script:AppVersion }
    Write-Step ('Install/update completed. Version ' + $version + '.') Green
    Write-Step ('Script path: ' + $script:InstalledScriptPath)
    Write-Step ('Icons path: ' + $script:InstalledIconsDir)
    Write-Step ('Remote asset base: ' + $script:ResolvedReleaseAssetBase)
    Write-Step ('Launcher path: ' + $script:LauncherVbsPath)

    if (-not $NoLaunch) {
        Write-Step 'Launching Tailscale Control without console...' Green
        Start-Process -FilePath $script:WScriptExe -ArgumentList ('"' + $script:LauncherVbsPath + '"')
    }
} finally {
    try {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}
