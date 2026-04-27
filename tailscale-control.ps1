# License: PolyForm Internal Use License 1.0.0
# Copyright (c) 2026 Luiz Bizzio

[CmdletBinding()]
param(
    [switch]$Background
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    $script:Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [Console]::OutputEncoding = $script:Utf8NoBomEncoding
    [Console]::InputEncoding = $script:Utf8NoBomEncoding
    $OutputEncoding = $script:Utf8NoBomEncoding
}
catch { }

$script:AppName = 'Tailscale Control'
$script:ScriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
$script:AppRoot = Join-Path $env:LOCALAPPDATA 'TailscaleControl'
$script:ConfigPath = Join-Path $script:AppRoot 'config.json'
$script:LogPath = Join-Path $script:AppRoot 'tailscale-control.log'
$script:StartupVbsPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'TailscaleControlLauncher.vbs'
$script:InstalledScriptPath = Join-Path $script:AppRoot 'tailscale-control.ps1'
$script:SourceRoot = Split-Path -Parent $script:ScriptPath
$script:InstalledIconsDir = Join-Path $script:AppRoot 'assets\icons'
$script:SourceIconsDir = Join-Path $script:SourceRoot 'assets\icons'
$script:LauncherVbsPath = Join-Path $script:AppRoot 'TailscaleControlLauncher.vbs'
$script:StartMenuShortcutPath = Join-Path ([Environment]::GetFolderPath('Programs')) 'Tailscale Control.lnk'
$script:WScriptExe = Join-Path $env:WINDIR 'System32\wscript.exe'
$script:PowerShellExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1)
if ([string]::IsNullOrWhiteSpace([string]$script:PowerShellExe)) {
    $script:PowerShellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
}
$script:ProgramDataRoot = Join-Path $env:ProgramData 'TailscaleControl'
$script:TailscaleClientTaskPath = '\TailscaleControl\'
$script:TailscaleClientCheckTaskName = 'TailscaleClientCheck'
$script:TailscaleClientUpdateTaskName = 'TailscaleClientUpdate'
$script:TailscaleClientElevatedRunnerPath = Join-Path $script:ProgramDataRoot 'elevated-tailscale-client-maintenance.ps1'
$script:TailscaleClientElevatedLauncherPath = Join-Path $script:ProgramDataRoot 'elevated-tailscale-client-maintenance.vbs'
$script:TailscaleClientCheckResultPath = Join-Path $script:AppRoot 'tailscale-client-check-result.json'
$script:TailscaleClientUpdateResultPath = Join-Path $script:AppRoot 'tailscale-client-update-result.json'
$script:IsBusy = $false
$script:IsRefreshing = $false
$script:IsClientMaintenanceTaskRunning = $false
$script:ClientMaintenanceWorker = $null
$script:IsClientTaskSetupRunning = $false
$script:ClientTaskSetupWorker = $null
$script:ClientTaskSetupOperation = ''
$script:TailscaleClientElevatedTaskReadyCache = @{ Check = $false; Update = $false }
$script:TailscaleClientElevatedTaskCacheInitialized = $false
$script:TailscaleClientElevatedTaskCacheCheckedUtc = [datetime]::MinValue
$script:IsClientTaskReadinessRefreshRunning = $false
$script:IsControlMaintenanceTaskRunning = $false
$script:ControlMaintenanceWorker = $null
$script:IsMtuInstallRunning = $false
$script:Exiting = $false
$script:Snapshot = $null
$script:PingState = @{}
$script:IsPingDiagnosticsTaskRunning = $false
$script:PingDiagnosticsWorker = $null
$script:IsDiagnosticsCommandTaskRunning = $false
$script:DiagnosticsCommandTask = $null
$script:IsAsyncActionRunning = $false
$script:AsyncActionTask = $null
$script:LastActionStartedAsync = $false
$script:DiagnosticsBusyButtons = @()
$script:DiagnosticsBusyFocusedButton = $null
$script:DiagnosticsContentMode = 'Selection'
$script:IsDnsResolveTaskRunning = $false
$script:DnsResolveTask = $null
$script:DnsResolveAutoDomain = ''
$script:IsPublicIpTaskRunning = $false
$script:PublicIpTask = $null
$script:NotifyIcon = $null
$script:OverlayForm = $null
$script:UiTimer = $null
$script:MainForm = $null
$script:MainFormPresentedOnce = $false
$script:HotkeyWindow = $null
$script:HotkeyPollTimer = $null
$script:HotkeyControls = @{}
$script:RegisteredHotkeys = @{}
$script:IsCapturingHotkey = $false
$script:Palette = $null
$script:IsApplyingMachineColumnLayout = $false
$script:MachineColumnsInitialized = $false
$script:StaticHotkeyNames = @('ToggleConnect','ToggleExitNode','ToggleDns','ToggleSubnets','ToggleIncoming','ShowSettings')
$script:QuickAccountSwitchMinimumRows = 2
$script:QuickAccountSwitchMaximumRows = 12
$script:QuickAccountSwitchNames = @('SwitchAccount1','SwitchAccount2')
$script:QuickAccountSwitchRows = @{}
$script:QuickAccountSwitchAccounts = @()
$script:QuickAccountSwitchAvailable = $false
$script:QuickAccountSwitchSectionPanel = $null
$script:QuickAccountSwitchHeader = $null
$script:QuickAccountSwitchSeparator = $null
$script:HotkeyIds = @{ ToggleConnect = 1; ToggleExitNode = 2; ToggleDns = 3; ToggleSubnets = 4; ToggleIncoming = 5; ShowSettings = 6; SwitchAccount1 = 101; SwitchAccount2 = 102 }
$script:ActionLabels = @{ ToggleConnect = 'Toggle Connect'; ToggleExitNode = 'Toggle Exit Node'; ToggleDns = 'Toggle DNS'; ToggleSubnets = 'Toggle Subnets'; ToggleIncoming = 'Toggle Incoming'; ShowSettings = 'Toggle Settings'; SwitchAccount1 = 'Switch Account 1'; SwitchAccount2 = 'Switch Account 2' }
$script:AppVersion = '1.0.0'
$script:ReleaseTag = 'v1.0.0'
$script:UpdateInstallerUrl = 'https://github.com/luizbizzio/tailscale-control/releases/latest/download/install.ps1'
$script:AppUserModelId = 'LuizBizzio.TailscaleControl'
$script:ActivitySeparator = '------------------'
$script:LogTailLineCount = 120
$script:LogMaxLineCount = 1200
$script:LogTrimToLineCount = 800
$script:LogWriteCountSinceOptimize = 0
$script:LastLogOptimizeAt = [datetime]::MinValue
$script:UiTextMaxChars = 120000
$script:UiTextMaxLines = 1200
$script:ActivityTextMaxChars = 90000
$script:ActivityOutputClearLineCount = 0
$script:MetricsTextMaxChars = 160000
$script:RefreshCountSinceGc = 0
$script:LastGcAt = [datetime]::MinValue
$script:TrayNetworkDevicesSignature = ''
$script:TraySelectAccountSignature = ''
$script:TrayExitNodeSignature = ''

foreach ($scriptVariableName in @(
    'AutoUpdateOverride',
    'btnAdminPanel',
    'btnActivityClearLog',
    'btnActivityClearOutput',
    'btnCheckClientUpdate',
    'btnInstallClientAutoUpdateTask',
    'btnCheckControlRepo',
    'btnCheckControlUpdate',
    'btnCheckMtuRepo',
    'btnExportDiagnostics',
    'btnCmdPingAll',
    'btnCmdPingDns',
    'btnCmdPingIPv4',
    'btnCmdPingIPv6',
    'btnCmdWhois',
    'btnDetailClearMetrics',
    'btnDetailMetrics',
    'btnDiagDns',
    'btnDiagIPs',
    'btnDiagMetrics',
    'btnDiagNetcheck',
    'btnDiagStatus',
    'btnDnsResolveRun',
    'btnInstallMtu',
    'btnOpenControlPath',
    'btnOpenMtu',
    'btnPublicIpRun',
    'btnRefresh',
    'btnRunClientUpdate',
    'btnUpdate',
    'btnToggleConnect',
    'btnToggleDns',
    'btnToggleExit',
    'btnToggleIncoming',
    'btnToggleSubnets',
    'btnUninstall',
    'chkCheckUpdateEvery',
    'chkAllowLan',
    'chkCloseToBackground',
    'chkControlCheckUpdateEvery',
    'chkExportRedactSensitive',
    'chkLogRefreshActivity',
    'chkShowCurrentDeviceInfoInTray',
    'chkShowTrayIcon',
    'chkStartMinimized',
    'chkStartup',
    'chkTogglePopups',
    'chkToggleSounds',
    'cmbDnsResolveResolver',
    'cmbExitNode',
    'gridMachines',
    'gridPing',
    'grpActions',
    'grpActivity',
    'grpHotkeys',
    'grpMachines',
    'grpMaintenance',
    'grpPreferences',
    'grpSummary',
    'headerPanel',
    'indDnsState',
    'indExitState',
    'indIncomingState',
    'indRoutesState',
    'IsCompactLayout',
    'IsLoadingConfig',
    'IsSavingSettings',
    'lblBackend',
    'lblBanner',
    'lblCheckUpdateHours',
    'lblConnSummary',
    'lblControlAuthor',
    'lblControlAutoUpdate',
    'lblControlCheckUpdateHours',
    'lblControlLastCheck',
    'lblControlLatestVersion',
    'lblControlPath',
    'lblControlRepo',
    'lblControlUpdateStatus',
    'lblControlVersion',
    'lblDevice',
    'lblDiagSelection',
    'lblDnsInUse',
    'lblDnsName',
    'lblDnsResolveHelp',
    'lblDnsResolveOther',
    'lblDnsResolveServerPreview',
    'lblDnsState',
    'lblExitState',
    'lblFooter',
    'lblIntro',
    'lblIPv4',
    'lblIPv6',
    'lblMachineHelp',
    'lblMaintAutoUpdate',
    'lblMaintLastCheck',
    'lblMaintLatestVersion',
    'lblMaintLocalVersion',
    'lblMaintMtuAuthor',
    'lblMaintMtuCheckInterval',
    'lblMaintMtuDesiredIPv4',
    'lblMaintMtuDesiredIPv6',
    'lblMaintMtuLastError',
    'lblMaintMtuLastResult',
    'lblMaintMtuRepo',
    'lblMaintMtuService',
    'lblMaintMtuStatus',
    'lblMaintMtuVersion',
    'lblMaintUpdateStatus',
    'lblMetricsInfo',
    'lblMtuIPv4',
    'lblMtuIPv6',
    'lblPingHelp',
    'lblPingNote',
    'lblPublicIpHelp',
    'lblRefreshValue',
    'lblRoutesState',
    'lblTailnet',
    'lblTitle',
    'lblToggleSoundVolumeValue',
    'lblUser',
    'lblUserEmail',
    'lblVersion',
    'numCheckUpdateHours',
    'numControlCheckUpdateHours',
    'radDnsResolveNoCache',
    'radDnsResolveUseCache',
    'radPublicIpDetailed',
    'radPublicIpFast',
    'statusStrip',
    'toolStatusLabel',
    'trkOverlay',
    'trkOverlayOpacity',
    'trkRefresh',
    'trkToggleSoundVolume',
    'txtDnsResolveDomain',
    'txtDnsResolveOtherServer',
    'txtDnsResolveOutput',
    'txtLog',
    'txtMachineDetails',
    'txtMachineFilter',
    'txtMetricsSummary',
    'txtPingDetails',
    'txtPublicIpOutput',
    'ToolTip'
)) {
    if ($null -eq (Get-Variable -Scope Script -Name $scriptVariableName -ErrorAction SilentlyContinue)) {
        Set-Variable -Scope Script -Name $scriptVariableName -Value $null
    }
}
$script:IsCompactLayout = $false
$script:IsLoadingConfig = $false
$script:IsSavingSettings = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$summaryLabelCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;

public sealed class TailscaleControlSummaryLabel : Label {
    public bool UseEndEllipsis { get; set; }

    public TailscaleControlSummaryLabel() {
        this.AutoSize = false;
        this.UseMnemonic = false;
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e) {
        Rectangle rect = this.ClientRectangle;
        if (rect.Width <= 0 || rect.Height <= 0) { return; }

        if (this.BackColor == Color.Transparent && this.Parent != null) {
            using (SolidBrush brush = new SolidBrush(this.Parent.BackColor)) {
                e.Graphics.FillRectangle(brush, rect);
            }
        }
        else {
            using (SolidBrush brush = new SolidBrush(this.BackColor)) {
                e.Graphics.FillRectangle(brush, rect);
            }
        }

        TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix;
        if (this.UseEndEllipsis) {
            flags |= TextFormatFlags.EndEllipsis;
        }

        Color textColor = this.Enabled ? this.ForeColor : SystemColors.GrayText;
        TextRenderer.DrawText(e.Graphics, this.Text ?? String.Empty, this.Font, rect, textColor, flags);
    }
}
"@
if (-not ('TailscaleControlSummaryLabel' -as [type])) {
    try { Add-Type -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' -TypeDefinition $summaryLabelCode } catch { }
}

$script:ToolTip = New-Object System.Windows.Forms.ToolTip
$script:ToolTip.AutoPopDelay = 20000
$script:ToolTip.InitialDelay = 400
$script:ToolTip.ReshowDelay = 100
$script:ToolTip.ShowAlways = $true

function Set-AppToolTip {
    param($Control,[string]$Text)
    try {
        if ($null -ne $Control -and -not [string]::IsNullOrWhiteSpace($Text) -and $null -ne $script:ToolTip) {
            $script:ToolTip.SetToolTip($Control, $Text)
        }
    } catch { }
}

function Set-AppToolTips {
    $tooltipItems = @(
        @{ Control = $script:chkStartup; Text = "Adds or removes the startup launcher so the app opens after Windows sign-in.`r`nDefault: On" },
        @{ Control = $script:chkStartMinimized; Text = "Keeps startup quiet by hiding the main window after launch.`r`nDefault: On" },
        @{ Control = $script:chkCloseToBackground; Text = "Keeps tray actions and hotkeys alive when you close the window instead of exiting the app.`r`nDefault: On" },
        @{ Control = $script:chkShowTrayIcon; Text = "Shows the tray icon so status and quick actions stay available without opening the main window.`r`nDefault: On" },
        @{ Control = $script:chkAllowLan; Text = "When an exit node is enabled, also allows access to local LAN devices such as your router, printer, or NAS.`r`nDefault: Off" },
        @{ Control = $script:chkTogglePopups; Text = "Shows a small confirmation overlay after a quick action or hotkey finishes.`r`nDefault: On" },
        @{ Control = $script:chkShowCurrentDeviceInfoInTray; Text = "Shows this device's tailnet, user, MagicDNS, IPs, DNS, and Tailscale version in the tray menu.`r`nDefault: Off" },
        @{ Control = $script:chkLogRefreshActivity; Text = "Writes automatic refresh cycles into Activity. Useful for debugging, noisy for normal use.`r`nDefault: Off" },
        @{ Control = $script:chkToggleSounds; Text = "Plays a short Windows sound when a toggle action changes a Tailscale setting.`r`nDefault: Off" },
        @{ Control = $script:trkOverlayOpacity; Text = "Changes how transparent the overlay is. Higher values make it more visible.`r`nDefault: 90%" },
        @{ Control = $script:trkOverlay; Text = "Changes how long the overlay remains visible after an action finishes.`r`nDefault: 1.5 seconds" },
        @{ Control = $script:trkRefresh; Text = "Sets how often the UI refreshes Tailscale status, devices, DNS, and MTU information.`r`nDefault: 10 seconds" },
        @{ Control = $script:trkToggleSoundVolume; Text = "Controls the volume used by the optional toggle sound.`r`nDefault: 20%" },
        @{ Control = $script:cmbExitNode; Text = "Selects which exit node the Toggle Exit Node action should prefer. Empty means the app will not force a preferred node.`r`nDefault: Empty" },
        @{ Control = $script:btnExportDiagnostics; Text = "Creates a diagnostic JSON file in the export folder, then opens that folder." },
        @{ Control = $script:chkExportRedactSensitive; Text = "Masks usernames, tailnet names, IP addresses, tokens, keys, and private network details in exported diagnostics.`r`nDefault: On" },
        @{ Control = $script:btnInstallClientAutoUpdateTask; Text = "Creates the elevated Windows scheduled tasks used for Tailscale client checks and updates. UAC is needed only for setup." },
        @{ Control = $script:chkCheckUpdateEvery; Text = "Runs Tailscale client check/update on this interval through the installed elevated task.`r`nDefault: Off" },
        @{ Control = $script:numCheckUpdateHours; Text = "Interval used by automatic Tailscale client check/update.`r`nDefault: 24 hours" },
        @{ Control = $script:chkControlCheckUpdateEvery; Text = "Automatically checks and updates Tailscale Control on this interval.`r`nDefault: Off" },
        @{ Control = $script:numControlCheckUpdateHours; Text = "Sets the interval for automatic Tailscale Control check/update runs.`r`nDefault: 24 hours" },
        @{ Control = $script:btnDiagStatus; Text = "Reads Tailscale status and summarizes backend state, DNS, peers, routes, and connection data." },
        @{ Control = $script:btnDiagNetcheck; Text = "Runs Tailscale netcheck to inspect UDP reachability, NAT behavior, port mapping, and DERP latency." },
        @{ Control = $script:btnDiagDns; Text = "Shows Tailscale DNS, MagicDNS, split DNS, resolvers, and what Windows DNS looks like to Tailscale." },
        @{ Control = $script:btnDiagIPs; Text = "Shows the local device Tailscale IPv4 and IPv6 addresses." },
        @{ Control = $script:btnDiagMetrics; Text = "Reads local Tailscale metrics and summarizes useful counters instead of showing raw text only." },
        @{ Control = $script:btnAdminPanel; Text = "Opens the Tailscale admin console in your browser." },
        @{ Control = $script:btnCmdPingAll; Text = "Tests the selected device by MagicDNS, Tailscale IPv4, and Tailscale IPv6 when available." },
        @{ Control = $script:btnCmdPingDns; Text = "Pings the selected device using its MagicDNS name." },
        @{ Control = $script:btnCmdPingIPv4; Text = "Pings the selected device using its Tailscale IPv4 address." },
        @{ Control = $script:btnCmdPingIPv6; Text = "Pings the selected device using its Tailscale IPv6 address." },
        @{ Control = $script:btnCmdWhois; Text = "Runs tailscale whois for the selected target to show owner and identity details known by Tailscale." },
        @{ Control = $script:lblDnsResolveHelp; Text = "Tests DNS resolution through the selected resolver path, useful when MagicDNS or split DNS behaves differently than Windows DNS." },
        @{ Control = $script:txtDnsResolveDomain; Text = "Enter the hostname or domain you want to resolve.`r`nExample: example.com" },
        @{ Control = $script:cmbDnsResolveResolver; Text = "Chooses which resolver path the lookup should use.`r`nDefault: Current" },
        @{ Control = $script:lblDnsResolveServerPreview; Text = "Shows the DNS server or resolver mode that will be used for the next lookup." },
        @{ Control = $script:txtDnsResolveOtherServer; Text = "Custom resolver used only when Resolver is set to Other.`r`nDefault: 1.1.1.1" },
        @{ Control = $script:radDnsResolveUseCache; Text = "Allows cached resolver answers when the underlying lookup supports it." },
        @{ Control = $script:radDnsResolveNoCache; Text = "Bypasses the Windows DNS client cache when possible, reducing confusion from stale cached results." },
        @{ Control = $script:lblPublicIpHelp; Text = "Checks which public IP the current route is using. Useful for validating exit nodes and DNS routing." },
        @{ Control = $script:radPublicIpFast; Text = "Uses fewer endpoints for a quicker public IP check." },
        @{ Control = $script:radPublicIpDetailed; Text = "Uses extra endpoints for more public IP and route context, but takes longer." },
        @{ Control = $script:btnActivityClearLog; Text = "Deletes the saved Activity log file content and clears the visible output." },
        @{ Control = $script:btnRefresh; Text = "Refreshes current Tailscale state, devices, DNS, MTU, maintenance status, and UI indicators." },
        @{ Control = $script:btnToggleConnect; Text = "Runs the connect/disconnect action for the local device, equivalent to switching Tailscale up or down." },
        @{ Control = $script:btnToggleExit; Text = "Enables the preferred exit node if none is active, or clears the current exit node if one is active." },
        @{ Control = $script:btnToggleDns; Text = "Switches whether this device accepts DNS settings from Tailscale." },
        @{ Control = $script:btnToggleSubnets; Text = "Switches whether this device accepts subnet routes advertised by other machines in the tailnet." },
        @{ Control = $script:btnToggleIncoming; Text = "Switches Tailscale Shields Up mode, which controls whether incoming Tailscale connections are allowed." }
    )
    foreach ($item in $tooltipItems) {
        Set-AppToolTip -Control $item.Control -Text $item.Text
    }
    foreach ($item in @(
        @{ Control = $lblOpacity; Text = "Changes how transparent the overlay is. Higher values make it more visible.`r`nDefault: 90%" },
        @{ Control = $overlayOpacityPanel; Text = "Changes how transparent the overlay is. Higher values make it more visible.`r`nDefault: 90%" },
        @{ Control = $lblOverlay; Text = "Changes how long the overlay remains visible after an action finishes.`r`nDefault: 1.5 seconds" },
        @{ Control = $overlaySecondsPanel; Text = "Changes how long the overlay remains visible after an action finishes.`r`nDefault: 1.5 seconds" },
        @{ Control = $lblRefresh; Text = "Sets how often the UI refreshes Tailscale status, devices, DNS, and MTU information.`r`nDefault: 10 seconds" },
        @{ Control = $refreshPanel; Text = "Sets how often the UI refreshes Tailscale status, devices, DNS, and MTU information.`r`nDefault: 10 seconds" },
        @{ Control = $lblToggleSoundVolume; Text = "Controls the volume used by the optional toggle sound.`r`nDefault: 20%" },
        @{ Control = $toggleSoundVolumePanel; Text = "Controls the volume used by the optional toggle sound.`r`nDefault: 20%" },
        @{ Control = $lblExitPref; Text = "Selects which exit node the Toggle Exit Node action should prefer. Empty means the app will not force a preferred node.`r`nDefault: Empty" },
        @{ Control = $dnsResolveDomainLabel; Text = "Enter the hostname or domain you want to resolve.`r`nExample: example.com" },
        @{ Control = $dnsResolveResolverLabel; Text = "Chooses which resolver path the lookup should use.`r`nDefault: Current" },
        @{ Control = $script:lblDnsResolveOther; Text = "Custom resolver used only when Resolver is set to Other.`r`nDefault: 1.1.1.1" },
        @{ Control = $dnsResolveCacheLabel; Text = "Chooses whether the DNS lookup may use cached answers or should try to bypass cache." }
    )) {
        Set-AppToolTip -Control $item.Control -Text $item.Text
    }
    Update-HotkeyToolTips
}

function Update-HotkeyToolTips {
    $defaults = @{
        ToggleConnect = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+T' }
        ToggleExitNode = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+E' }
        ToggleDns = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+D' }
        ToggleSubnets = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+S' }
        ToggleIncoming = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+I' }
        ShowSettings = @{ Enabled = $true; Shortcut = 'Ctrl+Alt+O' }
    }
    foreach ($name in $script:HotkeyNames) {
        try {
            if (-not $script:HotkeyControls.ContainsKey($name)) { continue }
            $controls = $script:HotkeyControls[$name]
            if ($null -eq $controls) { continue }
            $default = $defaults[$name]
            $quickIndex = Get-QuickAccountSwitchIndex -Name $name
            if ($null -eq $default -and $quickIndex -le 0) { continue }
            if ($quickIndex -gt 0) {
                if (-not [bool]$script:QuickAccountSwitchAvailable) {
                    $disabledText = 'Disabled until at least two Tailscale accounts are logged in.'
                    if ($null -ne $controls.Enabled) { Set-AppToolTip -Control $controls.Enabled -Text $disabledText }
                    if ($null -ne $controls.Capture) { Set-AppToolTip -Control $controls.Capture -Text $disabledText }
                    $accountControl = Get-ObjectPropertyOrDefault $controls 'Account' $null
                    if ($null -ne $accountControl) { Set-AppToolTip -Control $accountControl -Text $disabledText }
                    continue
                }
                $defaultState = 'Off'
                $defaultShortcut = if ($quickIndex -le 9) { 'Ctrl+Alt+Shift+' + [string]$quickIndex } elseif ($quickIndex -le 12) { 'Ctrl+Alt+Shift+F' + [string]($quickIndex - 9) } else { 'Empty' }
                if ($null -ne $controls.Enabled) { Set-AppToolTip -Control $controls.Enabled -Text ('Default: ' + $defaultState) }
                if ($null -ne $controls.Capture) { Set-AppToolTip -Control $controls.Capture -Text ('Default: ' + $defaultShortcut) }
                $accountControl = Get-ObjectPropertyOrDefault $controls 'Account' $null
                if ($null -ne $accountControl) { Set-AppToolTip -Control $accountControl -Text '' }
            }
            else {
                $defaultState = if ([bool]$default.Enabled) { 'On' } else { 'Off' }
                $defaultShortcut = [string]$default.Shortcut
                if ($null -ne $controls.Enabled) {
                    Set-AppToolTip -Control $controls.Enabled -Text ('Default: ' + $defaultState)
                }
                if ($null -ne $controls.Capture) {
                    Set-AppToolTip -Control $controls.Capture -Text ('Default: ' + $defaultShortcut)
                }
            }
            foreach ($control in @($controls.Label,$controls.Modifiers,$controls.Key)) {
                try { if ($null -ne $control -and $null -ne $script:ToolTip) { $script:ToolTip.SetToolTip($control, '') } } catch { }
            }
        } catch { }
    }
}
$trayTextRendererCode = @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;

public sealed class TailscaleControlTrayTextRendererV3 : ToolStripProfessionalRenderer {
    private static readonly Dictionary<string, Image> IconCache = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);

    private static Image GetMenuIcon(string path) {
        try {
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path)) { return null; }
            if (IconCache.ContainsKey(path)) { return IconCache[path]; }
            using (Icon icon = new Icon(path, 16, 16)) {
                Bitmap bitmap = icon.ToBitmap();
                Bitmap scaled = new Bitmap(16, 16);
                using (Graphics g = Graphics.FromImage(scaled)) {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.DrawImage(bitmap, new Rectangle(0, 0, 16, 16));
                }
                bitmap.Dispose();
                IconCache[path] = scaled;
                return scaled;
            }
        }
        catch {
            try {
                Image source = Image.FromFile(path);
                Bitmap scaled = new Bitmap(16, 16);
                using (Graphics g = Graphics.FromImage(scaled)) {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.DrawImage(source, new Rectangle(0, 0, 16, 16));
                }
                source.Dispose();
                IconCache[path] = scaled;
                return scaled;
            }
            catch { return null; }
        }
    }

    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {
        try {
            using (SolidBrush brush = new SolidBrush(SystemColors.Menu)) {
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
            }
            return;
        }
        catch { }
        base.OnRenderImageMargin(e);
    }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {
        try {
            if (e.Item != null && e.Item.Selected && e.Item.Enabled) {
                Rectangle rect = new Rectangle(1, 1, Math.Max(1, e.Item.Width - 2), Math.Max(1, e.Item.Height - 2));
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(229, 241, 251))) {
                    e.Graphics.FillRectangle(brush, rect);
                }
                return;
            }
        }
        catch { }
        base.OnRenderMenuItemBackground(e);
    }

    protected override void OnRenderItemImage(ToolStripItemImageRenderEventArgs e) {
        try {
            Image image = null;
            try {
                if (e.Item != null && !String.IsNullOrWhiteSpace(e.Item.AccessibleName)) {
                    image = GetMenuIcon(e.Item.AccessibleName);
                }
            }
            catch { image = null; }
            if (image == null) { image = e.Image; }
            if (image != null) {
                int size = 16;
                int imageWidth = Math.Max(size, e.ImageRectangle.Width);
                int x = e.ImageRectangle.X + Math.Max(0, (imageWidth - size) / 2);
                int y = Math.Max(0, (e.Item.Height - size) / 2);
                Rectangle rect = new Rectangle(x, y, size, size);
                e.Graphics.DrawImage(image, rect);
                return;
            }
        }
        catch { }
        base.OnRenderItemImage(e);
    }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) {
        e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        Rectangle rect = new Rectangle(e.TextRectangle.X, 0, e.TextRectangle.Width, e.Item.Height);
        TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
        if (e.Item.RightToLeft == RightToLeft.Yes) {
            flags = TextFormatFlags.Right | TextFormatFlags.RightToLeft | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
        }
        Color textColor = e.Item.Enabled ? SystemColors.MenuText : SystemColors.GrayText;
        if (e.Item.Enabled && !e.Item.ForeColor.IsEmpty) {
            textColor = e.Item.ForeColor;
        }
        TextRenderer.DrawText(e.Graphics, e.Text, e.TextFont, rect, textColor, flags);
    }

    protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e) {
        try {
            ToolStripMenuItem item = e.Item as ToolStripMenuItem;
            if (item != null && item.DropDownItems.Count > 0) {
                string marker = item.AccessibleDescription ?? String.Empty;
                if (marker.IndexOf("ManualTraySubmenuArrowGlyph", StringComparison.OrdinalIgnoreCase) >= 0) {
                    return;
                }
                Color arrowColor = e.Item.Enabled ? SystemColors.MenuText : SystemColors.GrayText;
                Rectangle arrowRect = e.ArrowRectangle;
                int x = arrowRect.Left + Math.Max(0, (arrowRect.Width - 5) / 2);
                int y = Math.Max(0, (e.Item.Height - 8) / 2);
                if (x <= 0) { x = Math.Max(8, e.Item.Width - 16); }
                Point[] points = new Point[] {
                    new Point(x, y),
                    new Point(x, y + 8),
                    new Point(x + 5, y + 4)
                };
                using (SolidBrush brush = new SolidBrush(arrowColor)) {
                    e.Graphics.FillPolygon(brush, points);
                }
                return;
            }
        }
        catch { }
        base.OnRenderArrow(e);
    }
}
"@
if (-not ('TailscaleControlTrayTextRendererV3' -as [type])) {
    try { Add-Type -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' -TypeDefinition $trayTextRendererCode } catch { }
}

$consoleCode = @"
using System;
using System.Runtime.InteropServices;
public static class TailscaleControlConsole {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
if (-not ('TailscaleControlConsole' -as [type])) {
    Add-Type $consoleCode
}

$taskbarCode = @"
using System;
using System.Runtime.InteropServices;

public static class TailscaleControlTaskbar {
    public const int WM_SETICON = 0x0080;
    public const int ICON_SMALL = 0;
    public const int ICON_BIG = 1;
    private const ushort VT_LPWSTR = 31;
    private const uint GPS_READWRITE = 0x00000002;

    [DllImport("shell32.dll", CharSet=CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string AppID);

    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000138-0000-0000-C000-000000000046")]
    private interface IPropertyStore {
        [PreserveSig] int GetCount(out uint cProps);
        [PreserveSig] int GetAt(uint iProp, out PROPERTYKEY pkey);
        [PreserveSig] int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        [PreserveSig] int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        [PreserveSig] int Commit();
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct PROPERTYKEY {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROPVARIANT {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public IntPtr p;
        public int p2;
    }

    [DllImport("shell32.dll")]
    private static extern int SHGetPropertyStoreForWindow(IntPtr hwnd, ref Guid riid, out IPropertyStore propertyStore);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetPropertyStoreFromParsingName(string pszPath, IntPtr pbc, uint flags, ref Guid riid, out IPropertyStore propertyStore);

    [DllImport("ole32.dll")]
    private static extern int PropVariantClear(ref PROPVARIANT pvar);

    private static PROPERTYKEY AppIdKey() {
        PROPERTYKEY key = new PROPERTYKEY();
        key.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
        key.pid = 5;
        return key;
    }

    private static PROPVARIANT StringPropVariant(string value) {
        PROPVARIANT pv = new PROPVARIANT();
        pv.vt = VT_LPWSTR;
        pv.p = Marshal.StringToCoTaskMemUni(value);
        return pv;
    }

    private static void SetAppId(IPropertyStore store, string appId) {
        if (store == null || String.IsNullOrWhiteSpace(appId)) return;
        PROPERTYKEY key = AppIdKey();
        PROPVARIANT pv = StringPropVariant(appId);
        try {
            store.SetValue(ref key, ref pv);
            store.Commit();
        }
        finally {
            try { PropVariantClear(ref pv); } catch { }
        }
    }

    public static void SetWindowAppUserModelID(IntPtr hwnd, string appId) {
        if (hwnd == IntPtr.Zero || String.IsNullOrWhiteSpace(appId)) return;
        Guid iid = new Guid("00000138-0000-0000-C000-000000000046");
        IPropertyStore store = null;
        try {
            int hr = SHGetPropertyStoreForWindow(hwnd, ref iid, out store);
            if (hr == 0 && store != null) SetAppId(store, appId);
        }
        finally {
            if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
        }
    }

    public static void SetShortcutAppUserModelID(string path, string appId) {
        if (String.IsNullOrWhiteSpace(path) || String.IsNullOrWhiteSpace(appId)) return;
        Guid iid = new Guid("00000138-0000-0000-C000-000000000046");
        IPropertyStore store = null;
        try {
            int hr = SHGetPropertyStoreFromParsingName(path, IntPtr.Zero, GPS_READWRITE, ref iid, out store);
            if (hr == 0 && store != null) SetAppId(store, appId);
        }
        finally {
            if (store != null) { try { Marshal.ReleaseComObject(store); } catch { } }
        }
    }
}
"@
if (-not ('TailscaleControlTaskbar' -as [type])) {
    Add-Type $taskbarCode
}

$hotkeyCode = @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class HotKeyEventArgs : EventArgs {
    public int Id { get; private set; }
    public HotKeyEventArgs(int id) { Id = id; }
}
public class TailscaleControlHotkeys : NativeWindow, IDisposable {
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    private const int WM_HOTKEY = 0x0312;
    public event EventHandler<HotKeyEventArgs> HotKeyPressed;
    public IntPtr WindowHandle { get { return this.Handle; } }
    public TailscaleControlHotkeys() { CreateHandle(new CreateParams()); }
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            var handler = HotKeyPressed;
            if (handler != null) handler(this, new HotKeyEventArgs(m.WParam.ToInt32()));
        }
        base.WndProc(ref m);
    }
    public void Dispose() { DestroyHandle(); }
}
"@
if (-not ('TailscaleControlHotkeys' -as [type])) {
    Add-Type -ReferencedAssemblies 'System.Windows.Forms' -TypeDefinition $hotkeyCode
}

$keyboardCode = @"
using System;
using System.Runtime.InteropServices;
public static class TailscaleControlKeyboard {
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
}
"@
if (-not ('TailscaleControlKeyboard' -as [type])) {
    Add-Type $keyboardCode
}

$trackBarCode = @"
using System.Windows.Forms;
public class TailscaleControlNoWheelTrackBar : TrackBar {
    private const int WM_MOUSEWHEEL = 0x020A;
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_MOUSEWHEEL) {
            return;
        }
        base.WndProc(ref m);
    }
    protected override void OnMouseWheel(MouseEventArgs e) {
    }
}
"@
if (-not ('TailscaleControlNoWheelTrackBar' -as [type])) {
    Add-Type -ReferencedAssemblies 'System.Windows.Forms' -TypeDefinition $trackBarCode
}

function Get-ToggleSoundPath {
    param([bool]$Enabled)
    $fileName = if ($Enabled) { 'Speech On.wav' } else { 'Speech Off.wav' }
    $candidate = Join-Path $env:WINDIR ('Media\' + $fileName)
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    return $null
}

function Get-ScaledWaveBytes {
    param(
        [string]$Path,
        [int]$Volume
    )
    [byte[]]$source = [System.IO.File]::ReadAllBytes($Path)
    if ($Volume -ge 100) { return ,$source }
    if ($source.Length -lt 44) { return ,$source }
    if ([System.Text.Encoding]::ASCII.GetString($source, 0, 4) -ne 'RIFF') { return ,$source }
    if ([System.Text.Encoding]::ASCII.GetString($source, 8, 4) -ne 'WAVE') { return ,$source }
    $audioFormat = 0
    $bitsPerSample = 0
    $dataOffset = -1
    $dataSize = 0
    $index = 12
    while ($index + 8 -le $source.Length) {
        $chunkId = [System.Text.Encoding]::ASCII.GetString($source, $index, 4)
        $chunkSize = [System.BitConverter]::ToInt32($source, $index + 4)
        if ($chunkSize -lt 0) { break }
        $chunkDataOffset = $index + 8
        if ($chunkDataOffset + $chunkSize -gt $source.Length) { break }
        if ($chunkId -eq 'fmt ' -and $chunkSize -ge 16) {
            $audioFormat = [System.BitConverter]::ToInt16($source, $chunkDataOffset)
            $bitsPerSample = [System.BitConverter]::ToInt16($source, $chunkDataOffset + 14)
        }
        elseif ($chunkId -eq 'data') {
            $dataOffset = $chunkDataOffset
            $dataSize = $chunkSize
        }
        $index = $chunkDataOffset + $chunkSize
        if (($chunkSize % 2) -eq 1) { $index++ }
    }
    if ($dataOffset -lt 0 -or $dataSize -le 0) { return ,$source }
    $factor = [double]$Volume / 100.0
    [byte[]]$output = New-Object byte[] ($source.Length)
    [System.Array]::Copy($source, $output, $source.Length)
    switch ($audioFormat) {
        1 {
            switch ($bitsPerSample) {
                8 {
                    for ($pos = $dataOffset; $pos -lt ($dataOffset + $dataSize); $pos++) {
                        $sample = [int]$output[$pos] - 128
                        $scaled = [int][math]::Round($sample * $factor)
                        if ($scaled -lt -128) { $scaled = -128 }
                        if ($scaled -gt 127) { $scaled = 127 }
                        $output[$pos] = [byte]($scaled + 128)
                    }
                }
                16 {
                    for ($pos = $dataOffset; $pos + 1 -lt ($dataOffset + $dataSize); $pos += 2) {
                        $sample = [System.BitConverter]::ToInt16($output, $pos)
                        $scaled = [int][math]::Round([double]$sample * $factor)
                        if ($scaled -lt -32768) { $scaled = -32768 }
                        if ($scaled -gt 32767) { $scaled = 32767 }
                        $bytes = [System.BitConverter]::GetBytes([int16]$scaled)
                        $output[$pos] = $bytes[0]
                        $output[$pos + 1] = $bytes[1]
                    }
                }
                24 {
                    for ($pos = $dataOffset; $pos + 2 -lt ($dataOffset + $dataSize); $pos += 3) {
                        $sample = [int]$output[$pos] -bor ([int]$output[$pos + 1] -shl 8) -bor ([int]$output[$pos + 2] -shl 16)
                        if ($sample -ge 0x800000) { $sample -= 0x1000000 }
                        $scaled = [int][math]::Round([double]$sample * $factor)
                        if ($scaled -lt -8388608) { $scaled = -8388608 }
                        if ($scaled -gt 8388607) { $scaled = 8388607 }
                        if ($scaled -lt 0) { $scaled += 0x1000000 }
                        $output[$pos] = [byte]($scaled -band 0xFF)
                        $output[$pos + 1] = [byte](($scaled -shr 8) -band 0xFF)
                        $output[$pos + 2] = [byte](($scaled -shr 16) -band 0xFF)
                    }
                }
                32 {
                    for ($pos = $dataOffset; $pos + 3 -lt ($dataOffset + $dataSize); $pos += 4) {
                        $sample = [System.BitConverter]::ToInt32($output, $pos)
                        $scaled = [long][math]::Round([double]$sample * $factor)
                        if ($scaled -lt -2147483648L) { $scaled = -2147483648L }
                        if ($scaled -gt 2147483647L) { $scaled = 2147483647L }
                        $bytes = [System.BitConverter]::GetBytes([int32]$scaled)
                        $output[$pos] = $bytes[0]
                        $output[$pos + 1] = $bytes[1]
                        $output[$pos + 2] = $bytes[2]
                        $output[$pos + 3] = $bytes[3]
                    }
                }
                default { return ,$source }
            }
        }
        3 {
            switch ($bitsPerSample) {
                32 {
                    for ($pos = $dataOffset; $pos + 3 -lt ($dataOffset + $dataSize); $pos += 4) {
                        $sample = [System.BitConverter]::ToSingle($output, $pos)
                        $scaled = [single]([double]$sample * $factor)
                        if ($scaled -lt -1.0) { $scaled = -1.0 }
                        if ($scaled -gt 1.0) { $scaled = 1.0 }
                        $bytes = [System.BitConverter]::GetBytes([single]$scaled)
                        $output[$pos] = $bytes[0]
                        $output[$pos + 1] = $bytes[1]
                        $output[$pos + 2] = $bytes[2]
                        $output[$pos + 3] = $bytes[3]
                    }
                }
                64 {
                    for ($pos = $dataOffset; $pos + 7 -lt ($dataOffset + $dataSize); $pos += 8) {
                        $sample = [System.BitConverter]::ToDouble($output, $pos)
                        $scaled = [double]$sample * $factor
                        if ($scaled -lt -1.0) { $scaled = -1.0 }
                        if ($scaled -gt 1.0) { $scaled = 1.0 }
                        $bytes = [System.BitConverter]::GetBytes([double]$scaled)
                        for ($i = 0; $i -lt 8; $i++) { $output[$pos + $i] = $bytes[$i] }
                    }
                }
                default { return ,$source }
            }
        }
        default { return ,$source }
    }
    return ,$output
}

function Set-BodySplitPreferredLayout {
    if ($null -eq $script:BodySplit -or $script:BodySplit.IsDisposed) { return }
    try {
        $bodySplit = $script:BodySplit
        $width = [int]$bodySplit.ClientSize.Width
        if ($width -le 0) { $width = [int]$bodySplit.Width }
        if ($width -le 0) { return }

        $panel1Min = $(if ($script:IsCompactLayout) { 360 } else { 390 })
        $panel2Floor = $(if ($script:IsCompactLayout) { 390 } else { 420 })
        $panel2Target = [Math]::Max($(if ($script:IsCompactLayout) { 430 } else { 470 }), [int]([math]::Floor($width * 0.445)))
        $availableForPanel2 = [Math]::Max(0, $width - $panel1Min - [int]$bodySplit.SplitterWidth)
        if ($availableForPanel2 -le 0) { return }

        $safePanel2 = [Math]::Min($panel2Target, $availableForPanel2)
        if ($safePanel2 -lt $panel2Floor) { $safePanel2 = [Math]::Min($panel2Floor, $availableForPanel2) }

        $desiredLeft = [Math]::Max($panel1Min, [int]([math]::Floor($width * 0.455)))
        $maxLeft = [Math]::Max($panel1Min, $width - $safePanel2 - [int]$bodySplit.SplitterWidth)
        if ($desiredLeft -gt $maxLeft) { $desiredLeft = $maxLeft }

        $bodySplit.SuspendLayout()
        try {
            $bodySplit.Panel1MinSize = $panel1Min
            if ($safePanel2 -gt 0) { $bodySplit.Panel2MinSize = $safePanel2 }
            if ($desiredLeft -ge $panel1Min) { $bodySplit.SplitterDistance = $desiredLeft }
        }
        finally {
            $bodySplit.ResumeLayout()
        }
    }
    catch {
        Write-LogException -Context 'Set body split preferred layout' -ErrorRecord $_
    }
}

function Hide-MainFormToBackground {
    if ($null -eq $script:MainForm) { return }
    if ($script:MainForm.InvokeRequired) {
        $null = $script:MainForm.BeginInvoke([Action]{ Hide-MainFormToBackground })
        return
    }
    try { $script:MainForm.Opacity = 0.0 } catch { }
    try { $script:MainForm.ShowInTaskbar = $false } catch { }
    try { $script:MainForm.WindowState = 'Normal' } catch { }
    try { $script:MainForm.Hide() } catch { }
    $script:StartupHidePending = $false
}

function Start-HideMainFormToBackground {
    if ($null -eq $script:MainForm -or $script:MainForm.IsDisposed) { return }
    $script:MainFormVisibilityToken = [int]$script:MainFormVisibilityToken + 1
    $queuedToken = [int]$script:MainFormVisibilityToken
    try {
        $action = [Action]{
            if ([int]$script:MainFormVisibilityToken -ne $queuedToken) { return }
            Hide-MainFormToBackground
        }.GetNewClosure()
        if ($script:MainForm.IsHandleCreated) {
            $null = $script:MainForm.BeginInvoke($action)
        }
        else {
            if ([int]$script:MainFormVisibilityToken -eq $queuedToken) { Hide-MainFormToBackground }
        }
    }
    catch {
        if ([int]$script:MainFormVisibilityToken -eq $queuedToken) { Hide-MainFormToBackground }
    }
}

function Switch-MainFormVisibility {
    if ($null -eq $script:MainForm -or $script:MainForm.IsDisposed) { return }
    if ($script:MainForm.InvokeRequired) {
        $null = $script:MainForm.BeginInvoke([Action]{ Switch-MainFormVisibility })
        return
    }
    $visible = $false
    try {
        $visible = [bool]$script:MainForm.Visible -and
            $script:MainForm.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized -and
            [double]$script:MainForm.Opacity -gt 0.05
    } catch { $visible = $false }
    if ($visible) {
        Hide-MainFormToBackground
        return
    }
    Show-MainForm
}

function Update-PreferenceSliderLabels {
    try {
        if ($null -ne $script:lblOverlayOpacityValue -and $null -ne $script:trkOverlayOpacity) { $script:lblOverlayOpacityValue.Text = ([int]$script:trkOverlayOpacity.Value).ToString() + '%' }
        if ($null -ne $script:lblOverlayValue -and $null -ne $script:trkOverlay) { $script:lblOverlayValue.Text = ('{0:N1}s' -f ([double]$script:trkOverlay.Value / 10.0)) }
        if ($null -ne $script:lblRefreshValue -and $null -ne $script:trkRefresh) { $script:lblRefreshValue.Text = ([int]$script:trkRefresh.Value).ToString() + 's' }
        if ($null -ne $script:lblToggleSoundVolumeValue -and $null -ne $script:trkToggleSoundVolume) { $script:lblToggleSoundVolumeValue.Text = ([int]$script:trkToggleSoundVolume.Value).ToString() + '%' }
    }
    catch { Write-LogException -Context 'Update preference slider labels' -ErrorRecord $_ }
}

function Invoke-ToggleFeedbackSound {
    param([bool]$Enabled)
    try {
        $cfg = Get-Config
        if (-not [bool](Get-ObjectPropertyOrDefault $cfg 'play_toggle_sounds' $false)) { return }
        $volume = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $cfg 'toggle_sound_volume' 20) 100
        if ($volume -lt 0) { $volume = 0 }
        if ($volume -gt 100) { $volume = 100 }
        if ($volume -eq 0) { return }
        $soundPath = Get-ToggleSoundPath -Enabled:$Enabled
        if (-not [string]::IsNullOrWhiteSpace([string]$soundPath) -and (Test-Path -LiteralPath $soundPath)) {
            for ($i = $script:ToggleSoundPlayers.Count - 1; $i -ge 0; $i--) {
                $entry = $script:ToggleSoundPlayers[$i]
                if ($null -eq $entry) { [void]$script:ToggleSoundPlayers.RemoveAt($i); continue }
                if (((Get-Date) - [datetime]$entry.CreatedAt).TotalSeconds -gt 15) {
                    try {
                        if ($null -ne $entry.Player -and $entry.Player.PSObject.Methods['Stop']) { $entry.Player.Stop() }
                        if ($null -ne $entry.Stream) { $entry.Stream.Dispose() }
                    }
                    catch { }
                    [void]$script:ToggleSoundPlayers.RemoveAt($i)
                }
            }
            try {
                [byte[]]$waveBytes = Get-ScaledWaveBytes -Path $soundPath -Volume $volume
                $memoryStream = New-Object System.IO.MemoryStream(,$waveBytes)
                $player = New-Object System.Media.SoundPlayer $memoryStream
                $player.Load()
                $player.Play()
                [void]$script:ToggleSoundPlayers.Add([pscustomobject]@{
                    Player = $player
                    Stream = $memoryStream
                    Backend = 'ScaledSoundPlayer'
                    CreatedAt = Get-Date
                })
                return
            }
            catch { }
            try {
                $player = New-Object System.Media.SoundPlayer $soundPath
                $player.Load()
                $player.Play()
                [void]$script:ToggleSoundPlayers.Add([pscustomobject]@{
                    Player = $player
                    Stream = $null
                    Backend = 'SoundPlayer'
                    CreatedAt = Get-Date
                })
                return
            }
            catch { }
        }
        if ($Enabled) { [System.Media.SystemSounds]::Asterisk.Play() } else { [System.Media.SystemSounds]::Exclamation.Play() }
    }
    catch { Write-LogException -Context 'Play toggle feedback sound' -ErrorRecord $_ }
}

$script:HotkeyExecutionLock = $false
$script:HotkeyExecutionName = $null
$script:HotkeyExecutionTimer = $null
$script:LastHotkeyAt = @{}
$script:HotkeyNames = @($script:StaticHotkeyNames + $script:QuickAccountSwitchNames)
$script:ConfigCache = $null
$script:IconCache = @{}
$script:TrayMenuItemIconFiles = @{}
$script:ToggleSoundPlayers = New-Object System.Collections.ArrayList
$script:SuppressToggleOverlay = $false
$script:SuppressNextAsyncToggleOverlay = $false
$script:InstanceActivateEvent = $null
$script:InstanceActivateTimer = $null
$script:trkOverlay = $null
$script:trkOverlayOpacity = $null
$script:trkRefresh = $null
$script:trkToggleSoundVolume = $null
$script:lblOverlayValue = $null
$script:lblOverlayOpacityValue = $null
$script:lblRefreshValue = $null
$script:lblToggleSoundVolumeValue = $null
$script:chkToggleSounds = $null
$script:btnDiagClear = $null
$script:TrayMenuShow = $null
$script:TrayMenuAdminPanel = $null
$script:TrayMenuAdminSwitchSeparator = $null
$script:TrayMenuSelectAccount = $null
$script:TrayMenuNetworkDevices = $null
$script:TrayMenuInfoTopSeparator = $null
$script:TrayMenuInfoBottomSeparator = $null
$script:TrayMenuMtu = $null
$script:TrayMenuInfoTailnet = $null
$script:TrayMenuInfoStatus = $null
$script:TrayMenuInfoUser = $null
$script:TrayMenuInfoAccountEmail = $null
$script:TrayMenuInfoDevice = $null
$script:TrayMenuInfoTailscaleVersion = $null
$script:TrayMenuInfoMagicDns = $null
$script:TrayMenuInfoIPv4 = $null
$script:TrayMenuInfoIPv6 = $null
$script:TrayMenuInfoDns = $null
$script:TrayMenuToggleConnect = $null
$script:TrayMenuToggleExitNode = $null
$script:TrayMenuToggleSubnets = $null
$script:TrayMenuToggleDns = $null
$script:TrayMenuToggleIncoming = $null
$script:TrayMenuChooseExitNode = $null
$script:TrayMenuAllowLanAccess = $null
$script:TrayMenuExit = $null
$script:BodySplit = $null
$script:StartupHidePending = $false
$script:MainFormVisibilityToken = 0

$script:SlowSnapshotCache = $null
$script:SlowSnapshotCachedAt = [datetime]::MinValue
$script:SlowSnapshotCacheExe = ''
$script:SlowSnapshotIntervalSeconds = 12

function Reset-SlowSnapshotCache {
    $script:SlowSnapshotCache = $null
    $script:SlowSnapshotCachedAt = [datetime]::MinValue
    $script:SlowSnapshotCacheExe = ''
}

function Get-SlowSnapshotData {
    param([string]$Exe,[switch]$Force)
    if ([string]::IsNullOrWhiteSpace([string]$Exe)) { return $null }

    $now = Get-Date
    $cacheAgeSeconds = if ($script:SlowSnapshotCachedAt -eq [datetime]::MinValue) { [double]::PositiveInfinity } else { ($now - $script:SlowSnapshotCachedAt).TotalSeconds }
    $cacheValid = (-not $Force) -and $null -ne $script:SlowSnapshotCache -and ([string]$script:SlowSnapshotCacheExe -eq [string]$Exe) -and ($cacheAgeSeconds -lt [double]$script:SlowSnapshotIntervalSeconds)
    if ($cacheValid) { return $script:SlowSnapshotCache }

    $data = [pscustomobject]@{
        Version = ConvertTo-PlainVersion ((Invoke-External -FilePath $Exe -Arguments @('version')).Output)
        Prefs = Get-PrefsObject -Exe $Exe
        ExitNodes = Get-ExitNodes -Exe $Exe
        DnsInfo = Get-TailscaleDnsInfo -Exe $Exe
        MtuInfo = Get-TailscaleMtuInfo
        MtuApp = Get-TailscaleMtuAppInfo
    }
    $script:SlowSnapshotCache = $data
    $script:SlowSnapshotCachedAt = $now
    $script:SlowSnapshotCacheExe = [string]$Exe
    return $data
}

function Initialize-AppRoot {
    if (-not (Test-Path -LiteralPath $script:AppRoot)) {
        New-Item -ItemType Directory -Path $script:AppRoot -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $script:InstalledIconsDir)) {
        New-Item -ItemType Directory -Path $script:InstalledIconsDir -Force | Out-Null
    }
}

function Remove-StaleTailscaleLoginScripts {
    param([int]$OlderThanMinutes = 15)
    try {
        Initialize-AppRoot
        $minutes = [math]::Max(1, [math]::Abs([int]$OlderThanMinutes))
        $cutoff = (Get-Date).AddMinutes(-$minutes)
        Get-ChildItem -LiteralPath $script:AppRoot -Filter 'tailscale-login-*.ps1' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
    catch { }
}

function Get-IconPathByFileName {
    param([string]$FileName)
    $candidates = @()
    if ([string]::IsNullOrWhiteSpace([string]$FileName)) { return $null }
    $candidates += (Join-Path $script:InstalledIconsDir $FileName)
    $candidates += (Join-Path $script:SourceIconsDir $FileName)
    $candidates += (Join-Path $script:AppRoot $FileName)
    $candidates += (Join-Path $script:SourceRoot $FileName)
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace([string]$candidate)) { continue }
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

function Initialize-IconInstalledByFileName {
    param([string]$FileName)
    Initialize-AppRoot
    if ([string]::IsNullOrWhiteSpace([string]$FileName)) { return $null }
    $installedPath = Join-Path $script:InstalledIconsDir $FileName
    if (Test-Path -LiteralPath $installedPath) { return $installedPath }
    foreach ($sourcePath in @((Join-Path $script:SourceIconsDir $FileName),(Join-Path $script:AppRoot $FileName),(Join-Path $script:SourceRoot $FileName))) {
        if ([string]::IsNullOrWhiteSpace([string]$sourcePath)) { continue }
        if (Test-Path -LiteralPath $sourcePath) {
            Copy-Item -LiteralPath $sourcePath -Destination $installedPath -Force
            return $installedPath
        }
    }
    return $null
}

function Get-AppIconPath { return (Get-IconPathByFileName -FileName 'tailscale-control.ico') }

function Initialize-AppIconInstalled { return (Initialize-IconInstalledByFileName -FileName 'tailscale-control.ico') }

function Get-IconCacheKey {
    param([string]$FileName,[string]$Kind,[int]$Size)
    return (($FileName + '|' + $Kind + '|' + [string]$Size).ToLowerInvariant())
}

function Get-CachedIconObject {
    param([string]$FileName,[int]$Size = 0)
    try {
        [void](Initialize-IconInstalledByFileName -FileName $FileName)
        $path = Get-IconPathByFileName -FileName $FileName
        $key = Get-IconCacheKey -FileName $FileName -Kind 'icon' -Size $Size
        if ($null -ne $script:IconCache -and $script:IconCache.ContainsKey($key)) { return $script:IconCache[$key] }
        $icon = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
            if ($Size -gt 0) {
                try {
                    $icon = New-Object System.Drawing.Icon -ArgumentList $path, $Size, $Size
                }
                catch {
                    try { $icon = New-Object System.Drawing.Icon -ArgumentList $path }
                    catch { $icon = $null }
                }
            }
            else {
                try { $icon = New-Object System.Drawing.Icon -ArgumentList $path }
                catch { $icon = $null }
            }
        }
        if ($null -eq $icon -and $FileName -eq 'tailscale-control.ico') { $icon = [System.Drawing.SystemIcons]::Application }
        if ($null -ne $icon) { $script:IconCache[$key] = $icon }
        return $icon
    }
    catch {
        if ($FileName -eq 'tailscale-control.ico') { return [System.Drawing.SystemIcons]::Application }
        return $null
    }
}

function Get-CachedIconBitmap {
    param([string]$FileName,[int]$Size = 32)
    try {
        $key = Get-IconCacheKey -FileName $FileName -Kind 'bitmap' -Size $Size
        if ($null -ne $script:IconCache -and $script:IconCache.ContainsKey($key)) { return $script:IconCache[$key] }
        $icon = Get-CachedIconObject -FileName $FileName -Size $Size
        if ($null -eq $icon) { return $null }
        $sourceBitmap = $icon.ToBitmap()
        $bitmap = New-Object System.Drawing.Bitmap $Size,$Size
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.DrawImage($sourceBitmap, 0, 0, $Size, $Size)
        }
        finally {
            $graphics.Dispose()
            $sourceBitmap.Dispose()
        }
        $script:IconCache[$key] = $bitmap
        return $bitmap
    }
    catch { return $null }
}

function Get-AppIcon { return (Get-CachedIconObject -FileName 'tailscale-control.ico' -Size 0) }

function Get-AppIcon32 { return (Get-CachedIconObject -FileName 'tailscale-control.ico' -Size 32) }

function Get-AppNotifyIcon {
    try {
        $key = Get-IconCacheKey -FileName 'tailscale-control.ico' -Kind 'notifyicon' -Size 32
        if ($null -ne $script:IconCache -and $script:IconCache.ContainsKey($key)) { return $script:IconCache[$key] }
        $sourceBitmap = Get-CachedIconBitmap -FileName 'tailscale-control.ico' -Size 32
        if ($null -eq $sourceBitmap) { return (Get-AppIcon32) }
        $bitmap = New-Object System.Drawing.Bitmap 32,32
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.Clear([System.Drawing.Color]::Transparent)
            $graphics.DrawImage($sourceBitmap, -1, -1, 34, 34)
            $handle = $bitmap.GetHicon()
            $icon = [System.Drawing.Icon]::FromHandle($handle)
            if ($null -ne $script:IconCache) { $script:IconCache[$key] = $icon }
            return $icon
        }
        finally {
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }
    catch {
        return (Get-AppIcon32)
    }
}

function Get-TrayMenuIconBitmap {
    param([string]$FileName)
    return (Get-CachedIconBitmap -FileName $FileName -Size 16)
}

function Enable-TrayMenuManualIconPaint {
    param($Item)
    try {
        if ($null -eq $Item) { return }
        $marker = [string]$Item.AccessibleDescription
        if ($marker -like '*ManualTrayIconPaint*') { return }
        if ([string]::IsNullOrWhiteSpace($marker)) { $Item.AccessibleDescription = 'ManualTrayIconPaint' } else { $Item.AccessibleDescription = ($marker + ';ManualTrayIconPaint') }
        $Item.add_Paint({
            param($paintSender,$paintEvent)
            try {
                if ($null -eq $paintSender -or $null -eq $paintEvent) { return }
                $key = [string]$paintSender.GetHashCode()
                if ($null -eq $script:TrayMenuItemIconFiles -or -not $script:TrayMenuItemIconFiles.ContainsKey($key)) { return }
                $fileName = [string]$script:TrayMenuItemIconFiles[$key]
                if ([string]::IsNullOrWhiteSpace($fileName)) { return }
                $image = Get-TrayMenuIconBitmap -FileName $fileName
                if ($null -eq $image) { return }
                $x = 7
                $y = [Math]::Max(0, [int](($paintSender.Height - 16) / 2))
                $paintEvent.Graphics.DrawImage($image, $x, $y, 16, 16)
            }
            catch { }
        }.GetNewClosure())
    }
    catch { }
}

function Set-TrayMenuItemIcon {
    param($Item,[string]$FileName)
    try {
        if ($null -eq $Item) { return }
        $path = Initialize-IconInstalledByFileName -FileName $FileName
        if ([string]::IsNullOrWhiteSpace([string]$path)) { $path = Get-IconPathByFileName -FileName $FileName }
        if ($null -ne $script:TrayMenuItemIconFiles) { $script:TrayMenuItemIconFiles[[string]$Item.GetHashCode()] = [string]$FileName }
        if (-not [string]::IsNullOrWhiteSpace([string]$path)) { $Item.AccessibleName = [string]$path }
        Enable-TrayMenuManualIconPaint -Item $Item
        $image = Get-TrayMenuIconBitmap -FileName $FileName
        if ($null -eq $image) {
            try { Write-DebugLog ('Tray menu icon not found or could not be loaded: ' + $FileName + ' | path=' + [string]$path) } catch { }
            $image = New-Object System.Drawing.Bitmap 16,16
        }
        $Item.Image = $image
        $Item.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $Item.ImageScaling = [System.Windows.Forms.ToolStripItemImageScaling]::None
        $Item.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
        try { $Item.Invalidate() } catch { }
    }
    catch { Write-LogException -Context ('Set tray menu icon ' + $FileName) -ErrorRecord $_ }
}

function Set-TrayTitleIcon { Set-TrayMenuItemIcon -Item $script:TrayMenuShow -FileName 'tailscale-control.ico' }

function Set-TrayTailscaleIcon { Set-TrayMenuItemIcon -Item $script:TrayMenuAdminPanel -FileName 'tailscale.ico' }

function Set-TrayMtuIcon { Set-TrayMenuItemIcon -Item $script:TrayMenuMtu -FileName 'tailscale-mtu.ico' }

function Set-MainFormAppIcon {
    try {
        if ($null -eq $script:MainForm -or $script:MainForm.IsDisposed) { return }
        [void](Initialize-AppIconInstalled)
        $smallIcon = Get-AppIcon32
        $bigIcon = Get-AppIcon
        $formIcon = if ($null -ne $smallIcon) { $smallIcon } else { $bigIcon }
        if ($null -ne $formIcon) {
            $script:MainForm.Icon = $formIcon
            $script:MainForm.ShowIcon = $true
            if ($script:MainForm.IsHandleCreated) {
                try { if ($null -ne $bigIcon) { [void][TailscaleControlTaskbar]::SendMessage($script:MainForm.Handle, [TailscaleControlTaskbar]::WM_SETICON, [IntPtr][TailscaleControlTaskbar]::ICON_BIG, $bigIcon.Handle) } } catch { }
                try { if ($null -ne $smallIcon) { [void][TailscaleControlTaskbar]::SendMessage($script:MainForm.Handle, [TailscaleControlTaskbar]::WM_SETICON, [IntPtr][TailscaleControlTaskbar]::ICON_SMALL, $smallIcon.Handle) } } catch { }
                try { [TailscaleControlTaskbar]::SetWindowAppUserModelID($script:MainForm.Handle, $script:AppUserModelId) } catch { }
            }
        }
    }
    catch { Write-LogException -Context 'Set main form icon' -ErrorRecord $_ }
}

function Optimize-LogFileIfNeeded {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return }
    try {
        $lineCount = 0
        foreach ($chunk in (Get-Content -LiteralPath $script:LogPath -ReadCount 500 -Encoding UTF8 -ErrorAction Stop)) {
            $lineCount += @($chunk).Count
            if ($lineCount -gt [int]$script:LogMaxLineCount) { break }
        }
        if ($lineCount -le [int]$script:LogMaxLineCount) { return }
        $keep = Get-Content -LiteralPath $script:LogPath -Tail ([int]$script:LogTrimToLineCount) -Encoding UTF8 -ErrorAction Stop
        Set-Content -LiteralPath $script:LogPath -Value $keep -Encoding UTF8
    }
    catch { }
}

function Write-Log {
    param([string]$Message)
    Initialize-AppRoot
    $line = ('{0} | {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message)
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    try {
        $script:LogWriteCountSinceOptimize = [int]$script:LogWriteCountSinceOptimize + 1
        $elapsedSeconds = ([datetime]::UtcNow - [datetime]$script:LastLogOptimizeAt).TotalSeconds
        if ($script:LogWriteCountSinceOptimize -ge 25 -or $elapsedSeconds -ge 120) {
            $script:LogWriteCountSinceOptimize = 0
            $script:LastLogOptimizeAt = [datetime]::UtcNow
            Optimize-LogFileIfNeeded
        }
    } catch { }
    try { Write-Host $line } catch { }
    try {
        if ($null -ne $script:txtLog -and $null -ne $script:MainForm -and -not $script:MainForm.IsDisposed) {
            $refreshActivity = [Action]{ try { Update-ActivityView -Text (Get-ActivityLogTail) } catch { } }
            if ($script:txtLog.InvokeRequired) { $null = $script:txtLog.BeginInvoke($refreshActivity) }
            else { & $refreshActivity }
        }
    }
    catch { }
}

function Test-DebugLogEnabled {
    try {
        $cfg = $script:ConfigCache
        if ($null -eq $cfg -and (Test-Path -LiteralPath $script:ConfigPath)) {
            $cfg = ConvertFrom-JsonSafe -Text (Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8) -Context 'Read log level'
        }
        $level = [string](Get-ObjectPropertyOrDefault $cfg 'log_level' 'INFO')
        return ($level.ToUpperInvariant() -eq 'DEBUG')
    }
    catch { return $false }
}

function Write-DebugLog {
    param([string]$Message)
    if (Test-DebugLogEnabled) { Write-Log $Message }
}

function Write-LogException {
    param(
        [string]$Context,
        $ErrorRecord
    )
    $message = ''
    try {
        if ($null -ne $ErrorRecord -and $null -ne $ErrorRecord.Exception -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.Exception.Message)) {
            $message = [string]$ErrorRecord.Exception.Message
        }
        elseif ($null -ne $ErrorRecord) {
            $message = [string]$ErrorRecord
        }
    }
    catch {
        $message = ''
    }
    if ([string]::IsNullOrWhiteSpace($message)) { $message = 'Unknown error' }
    Write-Log ($Context + ' failed: ' + $message)
}

function Get-CurrentSnapshot {
    if ($null -ne $script:Snapshot) { return $script:Snapshot }
    return Get-TailscaleSnapshot
}

function Write-ActivityCommandBlock {
    param([string]$Title,[string]$CommandText,[int]$ExitCode = 0,[string]$Output = '',[double]$DurationMs = 0)
    $cleanOutput = ConvertTo-DiagnosticText -Text ([string]$Output).TrimEnd()
    $lines = @(
        ([string]$Title),
        ('Command: ' + $CommandText),
        ('Exit code: ' + [string]$ExitCode)
    )
    if ([double]$DurationMs -gt 0) {
        $lines += ('Duration: ' + ('{0:N0} ms' -f [double]$DurationMs))
    }
    if (-not [string]::IsNullOrWhiteSpace($cleanOutput)) {
        $lines += ''
        $lines += $cleanOutput
    }
    $lines += $script:ActivitySeparator
    $lines += ''
    Write-Log (($lines -join [Environment]::NewLine))
}

function Write-ActivityFailureBlock {
    param([string]$Title,[string]$CommandText = '',[string]$Message = '',[double]$DurationMs = 0)
    $resolvedTitle = if ([string]::IsNullOrWhiteSpace([string]$Title)) { 'Action failed' } else { [string]$Title }
    $resolvedCommand = if ([string]::IsNullOrWhiteSpace([string]$CommandText)) { [string]$resolvedTitle } else { [string]$CommandText }
    $resolvedMessage = if ([string]::IsNullOrWhiteSpace([string]$Message)) { 'Unknown error.' } else { [string]$Message }
    Write-ActivityCommandBlock -Title $resolvedTitle -CommandText $resolvedCommand -ExitCode 1 -Output $resolvedMessage -DurationMs $DurationMs
}

function Get-CommandActivityOutput {
    param([string]$Summary,[string]$Output = '')
    $text = [string]$Summary
    if (-not [string]::IsNullOrWhiteSpace([string]$Output)) {
        $text += [Environment]::NewLine + [Environment]::NewLine + ([string]$Output).Trim()
    }
    return $text
}

function Get-VersionCheckActivityText {
    param($Check)
    $currentVersion = '-'
    $latestVersion = '-'
    $status = '-'
    try { if ($null -ne $Check -and $Check.PSObject.Properties['CurrentVersion']) { $currentVersion = [string]$Check.CurrentVersion } } catch { }
    try { if ($null -ne $Check -and $Check.PSObject.Properties['LatestVersion']) { $latestVersion = [string]$Check.LatestVersion } } catch { }
    try { if ($null -ne $Check -and $Check.PSObject.Properties['Status']) { $status = [string]$Check.Status } } catch { }
    if ([string]::IsNullOrWhiteSpace([string]$currentVersion)) { $currentVersion = '-' }
    if ([string]::IsNullOrWhiteSpace([string]$latestVersion)) { $latestVersion = '-' }
    if ([string]::IsNullOrWhiteSpace([string]$status)) { $status = '-' }
    $text = 'Current version: ' + $currentVersion + [Environment]::NewLine + 'Latest version: ' + $latestVersion + [Environment]::NewLine + 'Status: ' + $status
    $commandOutput = ''
    try { if ($null -ne $Check -and $Check.PSObject.Properties['Output']) { $commandOutput = [string]$Check.Output } } catch { }
    if (-not [string]::IsNullOrWhiteSpace([string]$commandOutput)) {
        $text += [Environment]::NewLine + [Environment]::NewLine + 'Tailscale output:' + [Environment]::NewLine + $commandOutput.Trim()
    }
    return $text
}

function Write-VersionCheckActivity {
    param([string]$Title,[string]$CommandText,$Check)
    $durationMs = 0
    try { if ($null -ne $Check -and $Check.PSObject.Properties['DurationMs']) { $durationMs = [double]$Check.DurationMs } } catch { }
    Write-ActivityCommandBlock -Title $Title -CommandText $CommandText -ExitCode 0 -Output (Get-VersionCheckActivityText -Check $Check) -DurationMs $durationMs
}

function Invoke-LoggedUrlOpen {
    param(
        [string]$Url,
        [string]$Title,
        [string]$CommandText,
        [string]$SuccessMessage,
        [string]$FailureTitle = '',
        [string]$FailureOverlayTitle = 'Open failed',
        [switch]$SkipActivity
    )
    if ([string]::IsNullOrWhiteSpace([string]$FailureTitle)) { $FailureTitle = $Title + ' failed' }
    try {
        Start-Process $Url | Out-Null
        if (-not $SkipActivity) {
            Write-ActivityCommandBlock -Title $Title -CommandText $CommandText -ExitCode 0 -Output $SuccessMessage
        }
    }
    catch {
        if (-not $SkipActivity) {
            Write-ActivityFailureBlock -Title $FailureTitle -CommandText $CommandText -Message $_.Exception.Message
        }
        Show-Overlay -Title $FailureOverlayTitle -Message $_.Exception.Message -ErrorStyle
    }
}

function Start-TailscaleProcess {
    param([string]$Exe,[string[]]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = ConvertTo-ProcessArgumentString -Arguments $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    try { $psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    try { $psi.StandardErrorEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    return [pscustomobject]@{ Process = $proc; Arguments = $Arguments; StartedUtc = [datetime]::UtcNow }
}

function Set-ControlFocusSafe {
    param($Control)
    try {
        if ($null -ne $Control) {
            $Control.Enabled = $true
            $Control.Select()
            [void]$Control.Focus()
            if ($null -ne $script:MainForm) { $script:MainForm.ActiveControl = $Control }
        }
    }
    catch { }
}

function Clear-UiFocusSafe {
    try {
        if ($null -eq $script:MainForm) { return }
        $script:MainForm.ActiveControl = $null
        [void]$script:MainForm.Focus()
    }
    catch { }
}

function Clear-UiFocusSoon {
    try {
        if ($null -eq $script:MainForm) { return }
        $action = [System.Action]{ Clear-UiFocusSafe }
        [void]$script:MainForm.BeginInvoke($action)
    }
    catch { Clear-UiFocusSafe }
}

function Get-NetworkDeviceActionButtons {
    $items = @(
        $script:btnDiagStatus,
        $script:btnDiagNetcheck,
        $script:btnDiagDns,
        $script:btnDiagIPs,
        $script:btnDiagMetrics,
        $script:btnDiagClear,
        $script:btnCmdPingAll,
        $script:btnCmdPingDns,
        $script:btnCmdPingIPv4,
        $script:btnCmdPingIPv6,
        $script:btnCmdWhois
    )
    return @($items | Where-Object { $null -ne $_ })
}

function Set-NetworkDeviceActionsBusyState {
    param([bool]$Busy)
    try {
        if ($Busy) {
            $script:DiagnosticsBusyButtons = @(Get-NetworkDeviceActionButtons)
            foreach ($btn in @($script:DiagnosticsBusyButtons)) {
                try { $btn.Enabled = $false } catch { }
            }
            if ($null -ne $script:btnAdminPanel) { $script:btnAdminPanel.Enabled = $true }
            Clear-UiFocusSoon
        }
        else {
            foreach ($btn in @($script:btnDiagStatus,$script:btnDiagNetcheck,$script:btnDiagDns,$script:btnDiagIPs,$script:btnDiagMetrics,$script:btnDiagClear)) {
                try { if ($null -ne $btn) { $btn.Enabled = $true } } catch { }
            }
            if ($null -ne $script:btnAdminPanel) { $script:btnAdminPanel.Enabled = $true }
            $script:DiagnosticsBusyButtons = @()
            $script:DiagnosticsBusyFocusedButton = $null
            try { Update-SelectedDeviceActionButtons } catch { }
        }
    }
    catch { Write-LogException -Context 'Set network/device actions busy state' -ErrorRecord $_ }
}

function Set-DiagnosticsCommandBusyState {
    param([bool]$Busy,$Button = $null,[string]$BusyText = 'Running...')
    try {
        if ($Busy) {
            $script:DiagnosticsBusyFocusedButton = $Button
            Set-NetworkDeviceActionsBusyState -Busy $true
        }
        else {
            Set-NetworkDeviceActionsBusyState -Busy $false
        }
    }
    catch { Write-LogException -Context 'Set diagnostics command busy state' -ErrorRecord $_ }
}

function Format-NetcheckOutputText {
    param([string]$Raw)
    $raw = ConvertTo-DiagnosticText -Text ([string]$Raw)
    $rawLines = @($raw -split "`r?`n")
    $reportIndex = -1
    for ($i = 0; $i -lt $rawLines.Count; $i++) {
        if (([string]$rawLines[$i]).Trim() -eq 'Report:') { $reportIndex = $i; break }
    }
    $lines = @('NETCHECK','')
    if ($reportIndex -ge 0) {
        $eventLines = @()
        for ($i = 0; $i -lt $reportIndex; $i++) {
            $line = [string]$rawLines[$i]
            if (-not [string]::IsNullOrWhiteSpace($line)) { $eventLines += $line }
        }
        $overview = @()
        $derp = @()
        $inDerp = $false
        for ($i = $reportIndex + 1; $i -lt $rawLines.Count; $i++) {
            $trim = ([string]$rawLines[$i]).Trim()
            if ([string]::IsNullOrWhiteSpace($trim)) { continue }
            if ($trim -eq '* DERP latency:' -or $trim -eq 'DERP latency:') { $inDerp = $true; continue }
            if (-not $inDerp -and $trim -match '^\*\s+([^:]+):\s*(.*)$') {
                $overview += ('  {0,-24} {1}' -f ($Matches[1].Trim() + ':'), $Matches[2].Trim())
                continue
            }
            if ($inDerp) {
                if ($trim -match '^\-\s*([a-z0-9]+):\s*(.*?)\s*\((.*?)\)\s*$') {
                    $lat = $Matches[2].Trim(); if ([string]::IsNullOrWhiteSpace($lat)) { $lat = '-' }
                    $derp += ('  {0,-8} {1,-10} {2}' -f ($Matches[1] + ':'), $lat, (ConvertTo-DiagnosticText -Text $Matches[3].Trim()))
                    continue
                }
                if ($trim -match '^\-\s*(.+)$') { $derp += ('  ' + (ConvertTo-DiagnosticText -Text $Matches[1].Trim())) }
            }
        }
        if ($overview.Count -gt 0) { $lines += 'OVERVIEW'; $lines += $overview; $lines += '' }
        if ($derp.Count -gt 0) { $lines += 'DERP LATENCY'; $lines += $derp; $lines += '' }
        if ($eventLines.Count -gt 0) { $lines += 'PORTMAP / ROUTER EVENTS'; $lines += ($eventLines | ForEach-Object { '  ' + $_ }); $lines += '' }
    }
    if ($lines.Count -le 2) { $lines += $raw.TrimEnd() }
    return (ConvertTo-DiagnosticText -Text ((($lines | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)))
}

function Format-DnsOutputText {
    param([string]$Raw)
    $rawText = ConvertTo-DiagnosticText -Text ([string]$Raw)
    $rawLines = @($rawText -split "`r?`n")

    $tailscaleEnabled = ''
    $magicDns = ''
    $deviceDnsName = ''
    $tailscaleResolvers = New-Object System.Collections.Generic.List[string]
    $splitDns = New-Object System.Collections.Generic.List[string]
    $fallbackResolvers = New-Object System.Collections.Generic.List[string]
    $searchDomains = New-Object System.Collections.Generic.List[string]
    $systemResolvers = New-Object System.Collections.Generic.List[string]

    $section = ''
    $subsection = ''
    foreach ($line in $rawLines) {
        $trim = ([string]$line).Trim()
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        if ($trim -like '===*Use Tailscale DNS*===') { $section = 'Use'; $subsection = ''; continue }
        if ($trim -like '===*MagicDNS configuration*===') { $section = 'Magic'; $subsection = ''; continue }
        if ($trim -like '===*System DNS configuration*===') { $section = 'System'; $subsection = ''; continue }
        if ($trim -match '^\[this is a preliminary version') { continue }
        if ($trim -match '^(Tailscale may use this configuration|This is the DNS configuration|Tailscale is configured to|Run .+ to revert)') { continue }

        if ($section -eq 'Use') {
            if ($trim -match '^Tailscale DNS:\s*(.+)$') { $tailscaleEnabled = [string]$Matches[1].Trim(); continue }
        }
        elseif ($section -eq 'Magic') {
            if ($trim -match '^MagicDNS:\s*(.+)$') { $magicDns = [string]$Matches[1].Trim(); continue }
            if ($trim -match '^Other devices in your tailnet can reach this device at\s+(.+)\.$') { $deviceDnsName = [string]$Matches[1].Trim(); continue }
            if ($trim -match '^Resolvers \(in preference order\):$') { $subsection = 'Resolvers'; continue }
            if ($trim -match '^Split DNS Routes:$') { $subsection = 'Split'; continue }
            if ($trim -match '^Fallback Resolvers:$') { $subsection = 'Fallback'; continue }
            if ($trim -match '^Search Domains:$') { $subsection = 'Search'; continue }
            if ($trim -match '^Nameservers IP Addresses:$') { $subsection = 'Nameservers'; continue }
            if ($trim -match '^Certificate Domains:$') { $subsection = 'Certificates'; continue }
            if ($trim -match '^Additional DNS Records:$') { $subsection = 'Additional'; continue }
            if ($trim -match '^Filtered suffixes when forwarding DNS queries as an exit node:$') { $subsection = 'Filtered'; continue }
            if ($trim -match '^\-\s+(.+)$') {
                $value = [string]$Matches[1].Trim()
                switch ($subsection) {
                    'Resolvers' { [void]$tailscaleResolvers.Add($value) }
                    'Split' { [void]$splitDns.Add($value) }
                    'Fallback' { if ($value -notmatch '^\(no ') { [void]$fallbackResolvers.Add($value) } }
                    'Search' { if ($value -notmatch '^\(no ') { [void]$searchDomains.Add($value) } }
                }
                continue
            }
        }
        elseif ($section -eq 'System') {
            if ($trim -match '^Nameservers:$') { $subsection = 'Nameservers'; continue }
            if ($trim -match '^Search domains:$') { $subsection = 'Search'; continue }
            if ($trim -match '^Match domains:$') { $subsection = 'Match'; continue }
            if ($trim -match '^\-\s+(.+)$') {
                $value = [string]$Matches[1].Trim()
                if ($subsection -eq 'Nameservers' -and $value -notmatch '^\(no ') { [void]$systemResolvers.Add($value) }
                continue
            }
        }
    }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('DNS')
    [void]$lines.Add('')
    [void]$lines.Add('OVERVIEW')
    [void]$lines.Add('  Tailscale DNS: ' + $(if ([string]::IsNullOrWhiteSpace($tailscaleEnabled)) { '-' } else { $tailscaleEnabled }))
    [void]$lines.Add('  MagicDNS: ' + $(if ([string]::IsNullOrWhiteSpace($magicDns)) { '-' } else { $magicDns }))
    [void]$lines.Add('  This device DNS: ' + $(if ([string]::IsNullOrWhiteSpace($deviceDnsName)) { '-' } else { $deviceDnsName }))
    [void]$lines.Add('')
    if (@($tailscaleResolvers).Count -gt 0) {
        [void]$lines.Add('TAILSCALE RESOLVERS')
        foreach ($item in $tailscaleResolvers) { [void]$lines.Add('  ' + [string]$item) }
        [void]$lines.Add('')
    }
    if (@($splitDns).Count -gt 0) {
        [void]$lines.Add('SPLIT DNS ROUTES')
        foreach ($item in $splitDns) { [void]$lines.Add('  ' + [string]$item) }
        [void]$lines.Add('')
    }
    if (@($searchDomains).Count -gt 0) {
        [void]$lines.Add('SEARCH DOMAINS')
        foreach ($item in $searchDomains) { [void]$lines.Add('  ' + [string]$item) }
        [void]$lines.Add('')
    }
    if (@($systemResolvers).Count -gt 0) {
        [void]$lines.Add('SYSTEM DNS SERVERS')
        foreach ($item in $systemResolvers) { [void]$lines.Add('  ' + [string]$item) }
        [void]$lines.Add('')
    }
    if (@($fallbackResolvers).Count -gt 0) {
        [void]$lines.Add('FALLBACK RESOLVERS')
        foreach ($item in $fallbackResolvers) { [void]$lines.Add('  ' + [string]$item) }
        [void]$lines.Add('')
    }

    $text = [string]($lines -join [Environment]::NewLine)
    $nonEmpty = ($text -replace 'DNS|OVERVIEW|Tailscale DNS: -|MagicDNS: -|This device DNS: -','').Trim()
    if ([string]::IsNullOrWhiteSpace($nonEmpty)) {
        return ('DNS' + [Environment]::NewLine + [Environment]::NewLine + $rawText.TrimEnd())
    }
    return $text.TrimEnd()
}

function Format-DiagnosticsCommandOutput {
    param([string]$Kind,[string]$Raw,[int]$ExitCode)
    $clean = ConvertTo-DiagnosticText -Text ([string]$Raw).TrimEnd()
    switch ($Kind) {
        'Netcheck' { return (Format-NetcheckOutputText -Raw $clean) }
        'DNS' { return (Format-DnsOutputText -Raw $clean) }
        'IPs' { return ('IPs' + [Environment]::NewLine + [Environment]::NewLine + $clean) }
        'Metrics' { return ('METRICS' + [Environment]::NewLine + [Environment]::NewLine + $clean) }
        'Whois' { return ('WHOIS' + [Environment]::NewLine + [Environment]::NewLine + $clean) }
        'Status' { return ('STATUS' + [Environment]::NewLine + [Environment]::NewLine + $clean) }
        default { return $clean }
    }
}

function Complete-DiagnosticsCommandTask {
    $taskRef = $script:DiagnosticsCommandTask
    if ($null -eq $taskRef) { return }
    try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
    try {
        $proc = $taskRef.Process
        $stdout = ''
        $stderr = ''
        try { $stdout = [string]$proc.StandardOutput.ReadToEnd() } catch { }
        try { $stderr = [string]$proc.StandardError.ReadToEnd() } catch { }
        $stdout = Remove-ChildProcessNoise -Text $stdout
        $stderr = Remove-ChildProcessNoise -Text $stderr
        $output = (($stdout,$stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
        $exit = 1
        try { $exit = [int]$proc.ExitCode } catch { }
        try { $proc.Dispose() } catch { }
        $elapsed = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
        $text = Format-DiagnosticsCommandOutput -Kind ([string]$taskRef.Kind) -Raw $output -ExitCode $exit
        Set-DiagnosticsOutput -Text $text -Mode 'Command'
        Write-ActivityCommandBlock -Title ([string]$taskRef.Title) -CommandText ([string]$taskRef.CommandText) -ExitCode $exit -Output $text -DurationMs ([double]$elapsed)
    }
    catch {
        Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix (([string]$taskRef.Title) + ' failed'))
        Write-ActivityFailureBlock -Title (([string]$taskRef.Title) + ' failed') -CommandText ([string]$taskRef.CommandText) -Message $_.Exception.Message
    }
    finally {
        $script:IsDiagnosticsCommandTaskRunning = $false
        $script:DiagnosticsCommandTask = $null
        Set-DiagnosticsCommandBusyState -Busy $false
    }
}

function Start-TailscaleDiagnosticsCommandAsync {
    param(
        [string]$Kind,
        [string]$Title,
        [string[]]$Arguments,
        $Button = $null,
        [string]$BusyText = 'Running...'
    )
    if ($script:IsDiagnosticsCommandTaskRunning -or $script:IsPingDiagnosticsTaskRunning) { return }
    $exe = ''
    try { if ($null -ne $script:Snapshot -and -not [string]::IsNullOrWhiteSpace([string]$script:Snapshot.Exe)) { $exe = [string]$script:Snapshot.Exe } } catch { }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { $exe = Find-TailscaleExe }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { throw 'tailscale.exe not found.' }
    try {
        $started = Start-TailscaleProcess -Exe $exe -Arguments $Arguments
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $script:IsDiagnosticsCommandTaskRunning = $true
        $script:DiagnosticsCommandTask = [pscustomobject]@{
            Kind = [string]$Kind
            Title = [string]$Title
            Process = $started.Process
            Arguments = $started.Arguments
            CommandText = 'tailscale ' + (($Arguments | ForEach-Object { [string]$_ }) -join ' ')
            Timer = $timer
            StartedUtc = [datetime]$started.StartedUtc
            Button = $Button
        }
        Set-DiagnosticsCommandBusyState -Busy $true -Button $Button -BusyText $BusyText
        $timer.add_Tick({
            try {
                $taskRef = $script:DiagnosticsCommandTask
                if ($null -eq $taskRef) { return }
                if ($null -ne $taskRef.Process -and -not $taskRef.Process.HasExited) { return }
                Complete-DiagnosticsCommandTask
            }
            catch {
                Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Diagnostics failed')
                $script:IsDiagnosticsCommandTaskRunning = $false
                $script:DiagnosticsCommandTask = $null
                Set-DiagnosticsCommandBusyState -Busy $false
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsDiagnosticsCommandTaskRunning = $false
        $script:DiagnosticsCommandTask = $null
        Set-DiagnosticsCommandBusyState -Busy $false
        throw
    }
}

function Register-TailscaleDiagnosticsButton {
    param([System.Windows.Forms.Control]$Button,[string]$Kind,[string]$Title,[string[]]$Arguments,[string]$BusyText = 'Running...')
    if ($null -eq $Button) { return }
    $handler = {
        if ($script:IsDiagnosticsCommandTaskRunning -or $script:IsPingDiagnosticsTaskRunning) { return }
        try { Start-TailscaleDiagnosticsCommandAsync -Kind $Kind -Title $Title -Arguments $Arguments -Button $this -BusyText $BusyText }
        catch { Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix ([string]$Title + ' failed')) }
    }.GetNewClosure()
    $Button.add_Click($handler)
}

function Register-PingButton {
    param([System.Windows.Forms.Control]$Button,[ValidateSet('All','DNS','IPv4','IPv6')] [string]$Kind)
    if ($null -eq $Button) { return }
    $handler = {
        if ($script:IsPingDiagnosticsTaskRunning -or $script:IsDiagnosticsCommandTaskRunning) { return }
        Set-ControlFocusSafe -Control $this
        if ($null -ne $this.Tag -and $this.Tag.PSObject.Properties['AllowPing'] -and -not [bool]$this.Tag.AllowPing) { return }
        try { Show-SelectedPingDiagnostics -Kind $Kind -SourceButton $this }
        catch { Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Ping failed') }
    }.GetNewClosure()
    $Button.add_Click($handler)
}

function Limit-UiText {
    param(
        [string]$Text,
        [int]$MaxChars = $script:UiTextMaxChars,
        [int]$MaxLines = $script:UiTextMaxLines
    )
    try {
        $value = [string]$Text
        if ([string]::IsNullOrEmpty($value)) { return '' }
        if ($MaxLines -gt 0) {
            $lines = $value -split "`r?`n"
            if ($lines.Count -gt $MaxLines) {
                $lines = @($lines | Select-Object -Last $MaxLines)
                $value = ('[output trimmed to last ' + $MaxLines + ' lines]' + [Environment]::NewLine + ($lines -join [Environment]::NewLine))
            }
        }
        if ($MaxChars -gt 0 -and $value.Length -gt $MaxChars) {
            $keep = [Math]::Max(0, $MaxChars - 96)
            $value = ('[output trimmed to last ' + $MaxChars + ' characters]' + [Environment]::NewLine + $value.Substring($value.Length - $keep))
        }
        return $value
    }
    catch { return [string]$Text }
}

function Invoke-MemoryCleanupThrottled {
    param([switch]$Force)
    try {
        $now = Get-Date
        $dueByTime = (($now - $script:LastGcAt).TotalMinutes -ge 5)
        $dueByCount = ([int]$script:RefreshCountSinceGc -ge 20)
        if (-not $Force -and -not $dueByTime -and -not $dueByCount) { return }
        $script:RefreshCountSinceGc = 0
        $script:LastGcAt = $now
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
    catch { }
}

function Update-ActivityView {
    param([string]$Text)
    if ($null -eq $script:txtLog) { return }
    $rawText = Limit-UiText -Text ([string](ConvertTo-DiagnosticText -Text $Text)) -MaxChars ([int]$script:ActivityTextMaxChars) -MaxLines ([int]$script:LogTailLineCount)
    try {
        $currentText = ''
        try { $currentText = [string]$script:txtLog.Text } catch { $currentText = '' }
        if ($currentText -ne $rawText) {
            $script:txtLog.SuspendLayout()
            $script:txtLog.Text = $rawText
            try {
                $baseColor = if ($null -ne $script:Palette) { $script:Palette.Text } else { [System.Drawing.Color]::White }
                $stampColor = [System.Drawing.Color]::FromArgb(45,120,255)
                $separatorColor = if ($null -ne $script:Palette) { $script:Palette.MutedText } else { [System.Drawing.Color]::Silver }

                $baseFont = $script:txtLog.Font
                $script:txtLog.SelectAll()
                $script:txtLog.SelectionColor = $baseColor
                $script:txtLog.SelectionFont = $baseFont

                $timestampMatches = [regex]::Matches($rawText, '(?m)^\[(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})\]')
                foreach ($m in $timestampMatches) {
                    if ($null -eq $m) { continue }
                    if ($m.Groups.Count -ge 3) {
                        $dateGroup = $m.Groups[1]
                        $timeGroup = $m.Groups[2]
                        if ($dateGroup.Length -gt 0) {
                            $script:txtLog.Select($dateGroup.Index, $dateGroup.Length)
                            $script:txtLog.SelectionColor = $stampColor
                            $script:txtLog.SelectionFont = $baseFont
                        }
                        if ($timeGroup.Length -gt 0) {
                            $script:txtLog.Select($timeGroup.Index, $timeGroup.Length)
                            $script:txtLog.SelectionColor = $stampColor
                            $script:txtLog.SelectionFont = $baseFont
                        }
                    }
                }

                $separatorMatches = [regex]::Matches($rawText, '(?m)^' + [regex]::Escape([string]$script:ActivitySeparator) + '$')
                foreach ($m in $separatorMatches) {
                    if ($null -eq $m) { continue }
                    $script:txtLog.Select($m.Index, $m.Length)
                    $script:txtLog.SelectionColor = $separatorColor
                    $script:txtLog.SelectionFont = $baseFont
                }

                $script:txtLog.Select($script:txtLog.TextLength, 0)
                $script:txtLog.SelectionColor = $baseColor
            }
            finally {
                $script:txtLog.ResumeLayout()
            }
        }
        $script:txtLog.SelectionStart = $script:txtLog.TextLength
        $script:txtLog.SelectionLength = 0
        $script:txtLog.ScrollToCaret()
    }
    catch { Write-LogException -Context 'Update activity view' -ErrorRecord $_ }
}

function Clear-ActivityOutputView {
    try {
        Initialize-AppRoot
        $script:ActivityOutputClearLineCount = [int](Get-LogLineCount)
        if ($null -ne $script:txtLog) {
            $script:txtLog.Clear()
            $script:txtLog.SelectionStart = 0
            $script:txtLog.SelectionLength = 0
        }
        if ($null -ne $script:toolStatusLabel) { $script:toolStatusLabel.Text = 'Activity output cleared.' }
    }
    catch { Write-LogException -Context 'Clear Activity output' -ErrorRecord $_ }
}

function Clear-ActivityLogFile {
    try {
        Initialize-AppRoot
        $script:ActivityOutputClearLineCount = 0
        Set-Content -LiteralPath $script:LogPath -Value '' -Encoding UTF8
        if ($null -ne $script:txtLog) {
            $script:txtLog.Clear()
            $script:txtLog.SelectionStart = 0
            $script:txtLog.SelectionLength = 0
        }
        if ($null -ne $script:toolStatusLabel) { $script:toolStatusLabel.Text = 'Activity log cleared.' }
    }
    catch { Write-LogException -Context 'Clear Activity log' -ErrorRecord $_ }
}

function ConvertTo-DiagnosticText {
    param([object]$Text)
    if ($null -eq $Text) { return '' }
    $t = [string]$Text

    $scoreText = {
        param([string]$Value)
        if ([string]::IsNullOrEmpty($Value)) { return [int]::MaxValue }
        $score = 0
        foreach ($token in @([string][char]0x00C3,[string][char]0x00C2,[string][char]0x00E2,[string][char]0xFFFD)) {
            if ($Value.Contains($token)) { $score += 10 }
        }
        $score += ([regex]::Matches($Value,'[\u2500-\u257F]').Count * 4)
        return $score
    }

    $needsRepair = $false
    foreach ($token in @([string][char]0x00C3,[string][char]0x00C2,[string][char]0x00E2,[string][char]0xFFFD)) {
        if ($t.Contains($token)) { $needsRepair = $true; break }
    }

    if ($needsRepair) {
        $encodings = @()
        try { $encodings += [System.Text.Encoding]::GetEncoding(1252) } catch { }
        try { $encodings += [System.Text.Encoding]::GetEncoding(28591) } catch { }
        try { $encodings += [System.Text.Encoding]::GetEncoding(437) } catch { }
        try { $encodings += [System.Text.Encoding]::GetEncoding(850) } catch { }
        $best = $t
        $bestScore = & $scoreText $best
        foreach ($enc in $encodings) {
            try { $candidate = [System.Text.Encoding]::UTF8.GetString($enc.GetBytes($t)) }
            catch { continue }
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $candidateScore = & $scoreText $candidate
            if ($candidateScore -lt $bestScore) { $best = $candidate; $bestScore = $candidateScore }
        }
        $t = $best
    }

    $t = $t.Replace([string]([char]0x2501), '-')
    try { $t = $t.Normalize([System.Text.NormalizationForm]::FormC) } catch { }
    return $t
}

function Write-AppLauncherVbs {
    param([string]$ScriptPath,[switch]$BackgroundMode)
    Initialize-AppRoot
    $extra = ''
    if ($BackgroundMode) { $extra = ' -Background' }
    $content = @"
Set oShell = CreateObject("WScript.Shell")
cmd = Chr(34) & "$($script:PowerShellExe)" & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & "$ScriptPath" & Chr(34) & "$extra"
oShell.Run cmd, 0, False
"@
    Set-Content -LiteralPath $script:LauncherVbsPath -Value $content -Encoding ASCII
}

function Initialize-StartMenuShortcut {
    Initialize-AppRoot
    Write-AppLauncherVbs -ScriptPath $script:InstalledScriptPath
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($script:StartMenuShortcutPath)
    $shortcut.TargetPath = $script:WScriptExe
    $shortcut.Arguments = '"' + $script:LauncherVbsPath + '"'
    $shortcut.WorkingDirectory = $script:AppRoot
    $iconPath = Initialize-AppIconInstalled
    if (-not [string]::IsNullOrWhiteSpace([string]$iconPath)) {
        $shortcut.IconLocation = $iconPath
    }
    else {
        $shortcut.IconLocation = "$env:SystemRoot\System32\SHELL32.dll,220"
    }
    $shortcut.Description = 'Open Tailscale Control'
    $shortcut.Save()
    try { [TailscaleControlTaskbar]::SetShortcutAppUserModelID($script:StartMenuShortcutPath, $script:AppUserModelId) } catch { }
}

function Install-TailscaleControl {
    Initialize-AppRoot
    if (Test-Path -LiteralPath $script:LogPath) {
        Remove-Item -LiteralPath $script:LogPath -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $script:ScriptPath -Destination $script:InstalledScriptPath -Force
    [void](Initialize-AppIconInstalled)
    Initialize-StartMenuShortcut
    Write-Log ('Install/update completed. Version ' + $script:AppVersion + '. Script path: ' + $script:InstalledScriptPath)
    Write-Log 'Start Menu shortcut recreated with hidden launcher and app icon.'
    return $script:InstalledScriptPath
}

function Read-ConfirmationText {
    param([string]$Prompt,[string]$Title)
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title
        $form.StartPosition = 'CenterParent'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MinimizeBox = $false
        $form.MaximizeBox = $false
        $form.ShowInTaskbar = $false
        $form.Width = 520
        $form.Height = 230
        $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Prompt
        $label.Left = 16
        $label.Top = 16
        $label.Width = 470
        $label.Height = 82
        $label.AutoEllipsis = $true

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Left = 16
        $textBox.Top = 106
        $textBox.Width = 470

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'OK'
        $ok.Width = 90
        $ok.Height = 30
        $ok.Left = 300
        $ok.Top = 145
        $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $cancel = New-Object System.Windows.Forms.Button
        $cancel.Text = 'Cancel'
        $cancel.Width = 90
        $cancel.Height = 30
        $cancel.Left = 396
        $cancel.Top = 145
        $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        [void]$form.Controls.Add($label)
        [void]$form.Controls.Add($textBox)
        [void]$form.Controls.Add($ok)
        [void]$form.Controls.Add($cancel)
        $form.AcceptButton = $ok
        $form.CancelButton = $cancel

        $owner = if ($null -ne $script:MainForm -and -not $script:MainForm.IsDisposed) { $script:MainForm } else { $null }
        $result = if ($null -ne $owner) { $form.ShowDialog($owner) } else { $form.ShowDialog() }
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return [string]$textBox.Text }
        return ''
    }
    catch {
        return ''
    }
    finally {
        try { if ($null -ne $form) { $form.Dispose() } } catch { }
    }
}

function New-TailscaleControlAppUninstallScript {
    param([int]$ProcessId)
    $payload = [pscustomobject]@{
        ProcessId = [int]$ProcessId
        AppRoot = [string]$script:AppRoot
        StartupVbsPath = [string]$script:StartupVbsPath
        StartMenuShortcutPath = [string]$script:StartMenuShortcutPath
        ProgramDataRoot = [string]$script:ProgramDataRoot
        RunnerPath = [string]$script:TailscaleClientElevatedRunnerPath
        LauncherPath = [string]$script:TailscaleClientElevatedLauncherPath
        CheckResultPath = [string]$script:TailscaleClientCheckResultPath
        UpdateResultPath = [string]$script:TailscaleClientUpdateResultPath
        TaskPath = [string]$script:TailscaleClientTaskPath
        CheckTaskName = [string]$script:TailscaleClientCheckTaskName
        UpdateTaskName = [string]$script:TailscaleClientUpdateTaskName
        PowerShellExe = [string]$script:PowerShellExe
    }
    $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes(($payload | ConvertTo-Json -Depth 8 -Compress)))
@"
param([switch]`$Elevated)
`$ErrorActionPreference = 'SilentlyContinue'
`$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
`$payload = `$payloadJson | ConvertFrom-Json
function Test-TailscaleControlUninstallerAdmin {
    try {
        `$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        `$principal = New-Object Security.Principal.WindowsPrincipal(`$identity)
        return `$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { return `$false }
}
function Test-TailscaleControlUninstallerNeedsElevation {
    try {
        foreach (`$taskName in @([string]`$payload.CheckTaskName, [string]`$payload.UpdateTaskName)) {
            if (-not [string]::IsNullOrWhiteSpace(`$taskName)) {
                `$task = Get-ScheduledTask -TaskPath ([string]`$payload.TaskPath) -TaskName `$taskName -ErrorAction SilentlyContinue
                if (`$null -ne `$task) { return `$true }
            }
        }
    } catch { }
    foreach (`$path in @([string]`$payload.ProgramDataRoot,[string]`$payload.RunnerPath,[string]`$payload.LauncherPath,[string]`$payload.CheckResultPath,[string]`$payload.UpdateResultPath)) {
        try { if (-not [string]::IsNullOrWhiteSpace(`$path) -and (Test-Path -LiteralPath `$path)) { return `$true } } catch { }
    }
    return `$false
}
function Remove-TailscaleControlElevatedArtifacts {
    foreach (`$taskName in @([string]`$payload.CheckTaskName, [string]`$payload.UpdateTaskName)) {
        try {
            if (-not [string]::IsNullOrWhiteSpace(`$taskName)) {
                `$task = Get-ScheduledTask -TaskPath ([string]`$payload.TaskPath) -TaskName `$taskName -ErrorAction SilentlyContinue
                if (`$null -ne `$task) { Unregister-ScheduledTask -TaskPath ([string]`$payload.TaskPath) -TaskName `$taskName -Confirm:`$false -ErrorAction SilentlyContinue }
            }
        } catch { }
    }
    foreach (`$path in @([string]`$payload.RunnerPath,[string]`$payload.LauncherPath,[string]`$payload.CheckResultPath,[string]`$payload.UpdateResultPath)) {
        try { if (-not [string]::IsNullOrWhiteSpace(`$path) -and (Test-Path -LiteralPath `$path)) { Remove-Item -LiteralPath `$path -Force -ErrorAction SilentlyContinue } } catch { }
    }
    try {
        `$root = [string]`$payload.ProgramDataRoot
        if (-not [string]::IsNullOrWhiteSpace(`$root) -and (Test-Path -LiteralPath `$root)) {
            `$children = @(Get-ChildItem -LiteralPath `$root -Force -ErrorAction SilentlyContinue)
            if (`$children.Count -eq 0) { Remove-Item -LiteralPath `$root -Force -ErrorAction SilentlyContinue }
        }
    } catch { }
    try {
        `$folderName = ([string]`$payload.TaskPath).Trim('\')
        if (-not [string]::IsNullOrWhiteSpace(`$folderName)) {
            `$service = New-Object -ComObject Schedule.Service
            `$service.Connect()
            `$folder = `$service.GetFolder('\' + `$folderName)
            if (`$folder.GetTasks(0).Count -eq 0 -and `$folder.GetFolders(0).Count -eq 0) {
                `$rootFolder = `$service.GetFolder('\')
                `$rootFolder.DeleteFolder(`$folderName, 0)
            }
        }
    } catch { }
}
function Remove-TailscaleControlUserArtifacts {
    foreach (`$path in @([string]`$payload.StartupVbsPath,[string]`$payload.StartMenuShortcutPath)) {
        try { if (-not [string]::IsNullOrWhiteSpace(`$path) -and (Test-Path -LiteralPath `$path)) { Remove-Item -LiteralPath `$path -Force -ErrorAction SilentlyContinue } } catch { }
    }
    `$appRoot = [string]`$payload.AppRoot
    if (-not [string]::IsNullOrWhiteSpace(`$appRoot)) {
        for (`$i = 0; `$i -lt 45; `$i++) {
            try {
                if (Test-Path -LiteralPath `$appRoot) { Remove-Item -LiteralPath `$appRoot -Recurse -Force -ErrorAction Stop }
                if (-not (Test-Path -LiteralPath `$appRoot)) { break }
            } catch { Start-Sleep -Milliseconds 700 }
        }
    }
}
function Show-TailscaleControlUninstallMessage {
    param([string]`$Message,[string]`$Icon)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(`$Message, 'Tailscale Control', 'OK', `$Icon) | Out-Null
    } catch { }
}
try {
    try {
        `$pidValue = [int]`$payload.ProcessId
        if (`$pidValue -gt 0 -and `$pidValue -ne `$PID) {
            `$proc = Get-Process -Id `$pidValue -ErrorAction SilentlyContinue
            if (`$null -ne `$proc) { [void]`$proc.WaitForExit(45000) }
        }
    } catch { }
    Start-Sleep -Milliseconds 800
    if ((Test-TailscaleControlUninstallerNeedsElevation) -and -not `$Elevated -and -not (Test-TailscaleControlUninstallerAdmin)) {
        `$ps = [string]`$payload.PowerShellExe
        if ([string]::IsNullOrWhiteSpace(`$ps) -or -not (Test-Path -LiteralPath `$ps)) { `$ps = Join-Path `$env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe' }
        `$args = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + `$PSCommandPath + '" -Elevated'
        Start-Process -FilePath `$ps -ArgumentList `$args -Verb RunAs -WindowStyle Hidden | Out-Null
        exit 0
    }
    Remove-TailscaleControlElevatedArtifacts
    Remove-TailscaleControlUserArtifacts
    Show-TailscaleControlUninstallMessage -Message 'Tailscale Control was uninstalled.' -Icon 'Information'
}
catch {
    Show-TailscaleControlUninstallMessage -Message ('Tailscale Control uninstall failed: ' + `$_.Exception.Message) -Icon 'Error'
}
finally {
    try {
        `$self = [string]`$PSCommandPath
        if (-not [string]::IsNullOrWhiteSpace(`$self) -and (Test-Path -LiteralPath `$self)) {
            `$cmd = 'ping 127.0.0.1 -n 3 >nul & del /f /q "' + `$self + '"'
            Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', `$cmd) -WindowStyle Hidden | Out-Null
        }
    } catch { }
}
"@
}

function Start-TailscaleControlAppUninstallProcess {
    $id = [guid]::NewGuid().ToString('N')
    $uninstallScriptPath = Join-Path $env:TEMP ('tailscale-control-uninstall-' + $id + '.ps1')
    Set-Content -LiteralPath $uninstallScriptPath -Value (New-TailscaleControlAppUninstallScript -ProcessId $PID) -Encoding UTF8 -Force
    $args = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $uninstallScriptPath + '"'
    Start-Process -FilePath $script:PowerShellExe -ArgumentList $args -WindowStyle Hidden -ErrorAction Stop | Out-Null
}

function Stop-TailscaleControlForUninstall {
    $script:Exiting = $true
    try { Clear-HotkeyExecutionLock } catch { Write-LogException -Context 'Release hotkey execution lock during uninstall' -ErrorRecord $_ }
    try { Unregister-Hotkeys | Out-Null } catch { Write-LogException -Context 'Unregister hotkeys during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:UiTimer) { $script:UiTimer.Stop(); $script:UiTimer.Dispose(); $script:UiTimer = $null } } catch { Write-LogException -Context 'Dispose UI timer during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:HotkeyPollTimer) { $script:HotkeyPollTimer.Stop(); $script:HotkeyPollTimer.Dispose(); $script:HotkeyPollTimer = $null } } catch { Write-LogException -Context 'Dispose hotkey fallback timer during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:InstanceActivateTimer) { $script:InstanceActivateTimer.Stop(); $script:InstanceActivateTimer.Dispose(); $script:InstanceActivateTimer = $null } } catch { Write-LogException -Context 'Dispose instance activation timer during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:InstanceActivateEvent) { $script:InstanceActivateEvent.Dispose(); $script:InstanceActivateEvent = $null } } catch { Write-LogException -Context 'Dispose instance activation event during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:NotifyIcon) { $script:NotifyIcon.Visible = $false; $script:NotifyIcon.Dispose(); $script:NotifyIcon = $null } } catch { Write-LogException -Context 'Dispose notify icon during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:TrayContextMenu) { $script:TrayContextMenu.Dispose(); $script:TrayContextMenu = $null } } catch { Write-LogException -Context 'Dispose tray menu during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:HotkeyWindow) { $script:HotkeyWindow.Dispose(); $script:HotkeyWindow = $null } } catch { Write-LogException -Context 'Dispose hotkey window during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:MainForm -and -not $script:MainForm.IsDisposed) { $script:MainForm.Hide(); $script:MainForm.Close() } } catch { Write-LogException -Context 'Close main form during uninstall' -ErrorRecord $_ }
    try { [System.Windows.Forms.Application]::Exit() } catch { }
}

function Invoke-UninstallApp {
    $prompt = "Type 'uninstall' to confirm full removal.`nThe app will close immediately, then remove the startup launcher, Start Menu shortcut, Tailscale client Auto Update Task, and all files in:`n" + $script:AppRoot
    $answer = Read-ConfirmationText -Prompt $prompt -Title 'Uninstall Tailscale Control'
    if (($null -eq $answer) -or ($answer.Trim().ToLowerInvariant() -ne 'uninstall')) {
        Show-Overlay -Title 'Uninstall canceled' -Message 'The confirmation text did not match.' -Indicator 'Info'
        return
    }
    try {
        Start-TailscaleControlAppUninstallProcess
        Write-Log 'Uninstall confirmed. External uninstaller started. Closing Tailscale Control.'
        Stop-TailscaleControlForUninstall
    }
    catch {
        Write-LogException -Context 'Start external app uninstall process' -ErrorRecord $_
        Show-Overlay -Title 'Uninstall failed' -Message $_.Exception.Message -ErrorStyle
    }
}

function Get-LogTail {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return '' }
    try { return ((Get-Content -LiteralPath $script:LogPath -Tail ([int]$script:LogTailLineCount) -Encoding UTF8 -ErrorAction Stop) -join [Environment]::NewLine) } catch { return '' }
}

function Get-LogLineCount {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return 0 }
    try {
        $lineCount = 0
        foreach ($chunk in (Get-Content -LiteralPath $script:LogPath -ReadCount 500 -Encoding UTF8 -ErrorAction Stop)) {
            $lineCount += @($chunk).Count
        }
        return [int]$lineCount
    }
    catch { return 0 }
}

function Get-ActivityLogTail {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return '' }
    $clearLine = 0
    try { $clearLine = [int]$script:ActivityOutputClearLineCount } catch { $clearLine = 0 }
    if ($clearLine -le 0) { return Get-LogTail }
    try {
        $lines = @(Get-Content -LiteralPath $script:LogPath -Encoding UTF8 -ErrorAction Stop)
        if ($lines.Count -le $clearLine) { return '' }
        $visible = @($lines | Select-Object -Skip $clearLine)
        if ($visible.Count -gt [int]$script:LogTailLineCount) { $visible = @($visible | Select-Object -Last ([int]$script:LogTailLineCount)) }
        return ($visible -join [Environment]::NewLine)
    }
    catch { return '' }
}

function Get-DefaultMachineColumnLayout {
    [ordered]@{
        Machine = 72
        Owner = 78
        IPv4 = 88
        IPv6 = 132
        OS = 54
        Connection = 118
        DNSName = 150
        LastSeen = 101
    }
}

function Get-MachineColumnMinWidthLayout {
    [ordered]@{
        Machine = 64
        Owner = 64
        IPv4 = 70
        IPv6 = 96
        OS = 48
        Connection = 84
        DNSName = 80
        LastSeen = 78
    }
}

function Get-NormalizedMachineColumnLayout {
    param($Candidate)
    $defaults = Get-DefaultMachineColumnLayout
    $mins = Get-MachineColumnMinWidthLayout
    $normalized = [ordered]@{}
    foreach ($name in $defaults.Keys) {
        $fallback = [int]$defaults[$name]
        $value = if ($null -ne $Candidate) { Convert-ToSafeInt (Get-ObjectPropertyOrDefault $Candidate $name $fallback) $fallback } else { $fallback }
        $minValue = [int]$mins[$name]
        if ($value -lt $minValue) { $value = $minValue }
        $normalized[$name] = [int]$value
    }
    return [pscustomobject]$normalized
}

function Get-QuickAccountSwitchName {
    param([int]$Index)
    return ('SwitchAccount' + [string]$Index)
}

function Get-QuickAccountSwitchIndex {
    param([string]$Name)
    $match = [regex]::Match([string]$Name, '^SwitchAccount(\d+)$')
    if ($match.Success) { return [int]$match.Groups[1].Value }
    return 0
}

function Get-QuickAccountSwitchCountFromHotkeys {
    param($Hotkeys)
    $count = [int]$script:QuickAccountSwitchMinimumRows
    try {
        if ($null -ne $Hotkeys) {
            foreach ($prop in @($Hotkeys.PSObject.Properties)) {
                $idx = Get-QuickAccountSwitchIndex -Name ([string]$prop.Name)
                if ($idx -gt $count) { $count = $idx }
            }
        }
    } catch { }
    if ($count -lt [int]$script:QuickAccountSwitchMinimumRows) { $count = [int]$script:QuickAccountSwitchMinimumRows }
    if ($count -gt [int]$script:QuickAccountSwitchMaximumRows) { $count = [int]$script:QuickAccountSwitchMaximumRows }
    return $count
}

function Ensure-QuickAccountSwitchDefinitions {
    param([int]$Count)
    if ($Count -lt [int]$script:QuickAccountSwitchMinimumRows) { $Count = [int]$script:QuickAccountSwitchMinimumRows }
    if ($Count -gt [int]$script:QuickAccountSwitchMaximumRows) { $Count = [int]$script:QuickAccountSwitchMaximumRows }
    $names = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $Count; $i++) {
        $name = Get-QuickAccountSwitchName -Index $i
        [void]$names.Add($name)
        if (-not $script:HotkeyIds.ContainsKey($name)) { $script:HotkeyIds[$name] = 100 + $i }
        $script:ActionLabels[$name] = 'Account ' + [string]$i
    }
    $script:QuickAccountSwitchNames = @($names.ToArray())
    $script:HotkeyNames = @($script:StaticHotkeyNames + $script:QuickAccountSwitchNames)
}

function Get-DefaultHotkeyEntry {
    param([string]$Name)
    switch ($Name) {
        'ToggleConnect' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'T' } }
        'ToggleExitNode' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'E' } }
        'ToggleDns' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'D' } }
        'ToggleSubnets' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'S' } }
        'ToggleIncoming' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'I' } }
        'ShowSettings' { return [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'O' } }
        default {
            $idx = Get-QuickAccountSwitchIndex -Name $Name
            if ($idx -gt 0) {
                $key = if ($idx -le 9) { [string]$idx } elseif ($idx -le 12) { 'F' + [string]($idx - 9) } else { '' }
                return [pscustomobject]@{ enabled = $false; modifiers = 'Ctrl+Alt+Shift'; key = $key; account_identifier = ''; account_display = '' }
            }
        }
    }
    return [pscustomobject]@{ enabled = $false; modifiers = 'Ctrl+Alt'; key = '' }
}

function Ensure-ConfigHotkeyEntries {
    param($Config,[int]$Count)
    if ($null -eq $Config) { return }
    Ensure-QuickAccountSwitchDefinitions -Count $Count
    if ($null -eq $Config.hotkeys) {
        Add-Member -InputObject $Config -MemberType NoteProperty -Name 'hotkeys' -Value ([pscustomobject]@{}) -Force
    }
    foreach ($name in $script:HotkeyNames) {
        $fallback = Get-DefaultHotkeyEntry -Name $name
        $prop = $Config.hotkeys.PSObject.Properties[$name]
        if ($null -eq $prop -or $null -eq $prop.Value) {
            Add-Member -InputObject $Config.hotkeys -MemberType NoteProperty -Name $name -Value $fallback -Force
            continue
        }
        $entry = $prop.Value
        if ($null -eq $entry.PSObject -or $null -eq $entry.PSObject.Properties['enabled']) { $entry = $fallback }
        if ($null -eq $entry.PSObject.Properties['enabled']) { Add-Member -InputObject $entry -MemberType NoteProperty -Name 'enabled' -Value ([bool]$fallback.enabled) -Force }
        if ($null -eq $entry.PSObject.Properties['modifiers']) { Add-Member -InputObject $entry -MemberType NoteProperty -Name 'modifiers' -Value ([string]$fallback.modifiers) -Force }
        if ($null -eq $entry.PSObject.Properties['key']) { Add-Member -InputObject $entry -MemberType NoteProperty -Name 'key' -Value ([string]$fallback.key) -Force }
        if ((Get-QuickAccountSwitchIndex -Name $name) -gt 0) {
            if ($null -eq $entry.PSObject.Properties['account_identifier']) { Add-Member -InputObject $entry -MemberType NoteProperty -Name 'account_identifier' -Value '' -Force }
            if ($null -eq $entry.PSObject.Properties['account_display']) { Add-Member -InputObject $entry -MemberType NoteProperty -Name 'account_display' -Value '' -Force }
        }
        Add-Member -InputObject $Config.hotkeys -MemberType NoteProperty -Name $name -Value $entry -Force
    }
}

function Get-DefaultConfig {
    [pscustomobject]@{
        start_with_windows = $true
        start_minimized = $true
        close_to_background = $true
        show_tray_icon = $true
        allow_lan_on_exit = $false
        overlay_seconds = 1.5
        overlay_opacity = 90
        show_toggle_popups = $true
        show_current_device_info_in_tray = $false
        log_refresh_activity = $false
        play_toggle_sounds = $false
        toggle_sound_volume = 20
        refresh_seconds = 10
        log_level = 'INFO'
        check_update_every_enabled = $false
        check_update_every_hours = 24
        last_client_update_check_utc = ''
        last_client_update_latest_version = ''
        last_client_update_status = ''
        control_check_update_every_enabled = $false
        control_check_update_every_hours = 24
        last_control_update_check_utc = ''
        last_control_update_latest_version = ''
        last_control_update_status = ''
        preferred_exit_node = ''
        preferred_exit_label = ''
        settings_version = 1
        update_url = $script:UpdateInstallerUrl
        mtu_install_url = 'https://raw.githubusercontent.com/luizbizzio/tailscale-mtu/refs/heads/main/windows/windows-setup.ps1'
        machine_columns = [pscustomobject](Get-DefaultMachineColumnLayout)
        hotkeys = [pscustomobject]@{
            ToggleConnect = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'T' }
            ToggleExitNode = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'E' }
            ToggleDns = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'D' }
            ToggleSubnets = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'S' }
            ToggleIncoming = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'I' }
            ShowSettings = [pscustomobject]@{ enabled = $true; modifiers = 'Ctrl+Alt'; key = 'O' }
        }
    }
}

function Get-ObjectPropertyOrDefault {
    param($Object,[string]$Name,$DefaultValue = $null)
    if ($null -eq $Object) { return $DefaultValue }
    try {
        $prop = $Object.PSObject.Properties[$Name]
        if ($null -eq $prop) { return $DefaultValue }
        if ($null -eq $prop.Value) { return $DefaultValue }
        return $prop.Value
    }
    catch { return $DefaultValue }
}

function Convert-ToSafeInt {
    param($Value,[int]$DefaultValue)
    if ($null -eq $Value) { return $DefaultValue }
    $parsed = 0
    if ([int]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

function Convert-ToSafeDouble {
    param($Value,[double]$DefaultValue)
    if ($null -eq $Value) { return $DefaultValue }
    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { return $parsed }
    if ([double]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

function Get-Config {
    Initialize-AppRoot
    if ($null -ne $script:ConfigCache) { return $script:ConfigCache }

    $default = Get-DefaultConfig
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        Ensure-ConfigHotkeyEntries -Config $default -Count $script:QuickAccountSwitchMinimumRows
        Save-Config -Config $default
        return $script:ConfigCache
    }

    try {
        $configText = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8
        $loaded = ConvertFrom-JsonSafe -Text $configText -Context 'Load config'
        $configText = $null

        $merged = [pscustomobject]@{
            start_with_windows = [bool](Get-ObjectPropertyOrDefault $loaded 'start_with_windows' $default.start_with_windows)
            start_minimized = [bool](Get-ObjectPropertyOrDefault $loaded 'start_minimized' $default.start_minimized)
            close_to_background = [bool](Get-ObjectPropertyOrDefault $loaded 'close_to_background' $default.close_to_background)
            show_tray_icon = [bool](Get-ObjectPropertyOrDefault $loaded 'show_tray_icon' $default.show_tray_icon)
            allow_lan_on_exit = [bool](Get-ObjectPropertyOrDefault $loaded 'allow_lan_on_exit' $default.allow_lan_on_exit)
            overlay_seconds = Convert-ToSafeDouble (Get-ObjectPropertyOrDefault $loaded 'overlay_seconds' $default.overlay_seconds) $default.overlay_seconds
            overlay_opacity = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'overlay_opacity' $default.overlay_opacity) $default.overlay_opacity
            show_toggle_popups = [bool](Get-ObjectPropertyOrDefault $loaded 'show_toggle_popups' $default.show_toggle_popups)
            show_current_device_info_in_tray = [bool](Get-ObjectPropertyOrDefault $loaded 'show_current_device_info_in_tray' $default.show_current_device_info_in_tray)
            log_refresh_activity = [bool](Get-ObjectPropertyOrDefault $loaded 'log_refresh_activity' $default.log_refresh_activity)
            play_toggle_sounds = [bool](Get-ObjectPropertyOrDefault $loaded 'play_toggle_sounds' $default.play_toggle_sounds)
            toggle_sound_volume = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'toggle_sound_volume' $default.toggle_sound_volume) $default.toggle_sound_volume
            refresh_seconds = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'refresh_seconds' $default.refresh_seconds) $default.refresh_seconds
            log_level = [string](Get-ObjectPropertyOrDefault $loaded 'log_level' $default.log_level)
            check_update_every_enabled = [bool](Get-ObjectPropertyOrDefault $loaded 'check_update_every_enabled' $default.check_update_every_enabled)
            check_update_every_hours = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'check_update_every_hours' $default.check_update_every_hours) $default.check_update_every_hours
            last_client_update_check_utc = [string](Get-ObjectPropertyOrDefault $loaded 'last_client_update_check_utc' $default.last_client_update_check_utc)
            last_client_update_latest_version = [string](Get-ObjectPropertyOrDefault $loaded 'last_client_update_latest_version' $default.last_client_update_latest_version)
            last_client_update_status = [string](Get-ObjectPropertyOrDefault $loaded 'last_client_update_status' $default.last_client_update_status)
            control_check_update_every_enabled = [bool](Get-ObjectPropertyOrDefault $loaded 'control_check_update_every_enabled' $default.control_check_update_every_enabled)
            control_check_update_every_hours = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'control_check_update_every_hours' $default.control_check_update_every_hours) $default.control_check_update_every_hours
            last_control_update_check_utc = [string](Get-ObjectPropertyOrDefault $loaded 'last_control_update_check_utc' $default.last_control_update_check_utc)
            last_control_update_latest_version = [string](Get-ObjectPropertyOrDefault $loaded 'last_control_update_latest_version' $default.last_control_update_latest_version)
            last_control_update_status = [string](Get-ObjectPropertyOrDefault $loaded 'last_control_update_status' $default.last_control_update_status)
            preferred_exit_node = [string](Get-ObjectPropertyOrDefault $loaded 'preferred_exit_node' $default.preferred_exit_node)
            preferred_exit_label = [string](Get-ObjectPropertyOrDefault $loaded 'preferred_exit_label' $default.preferred_exit_label)
            settings_version = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $loaded 'settings_version' $default.settings_version) $default.settings_version
            update_url = [string](Get-ObjectPropertyOrDefault $loaded 'update_url' $default.update_url)
            mtu_install_url = [string](Get-ObjectPropertyOrDefault $loaded 'mtu_install_url' $default.mtu_install_url)
            machine_columns = Get-NormalizedMachineColumnLayout (Get-ObjectPropertyOrDefault $loaded 'machine_columns' $null)
            hotkeys = [pscustomobject]@{}
        }
        if ($merged.overlay_seconds -lt 0.5) { $merged.overlay_seconds = 0.5 }
        if ($merged.overlay_seconds -gt 5.0) { $merged.overlay_seconds = 5.0 }
        $merged.overlay_seconds = [math]::Round([double]$merged.overlay_seconds, 1)
        if ($merged.overlay_opacity -lt 0) { $merged.overlay_opacity = 0 }
        if ($merged.overlay_opacity -gt 100) { $merged.overlay_opacity = 100 }
        if ($merged.toggle_sound_volume -lt 0) { $merged.toggle_sound_volume = 0 }
        if ($merged.toggle_sound_volume -gt 100) { $merged.toggle_sound_volume = 100 }
        if ($merged.refresh_seconds -lt 3) { $merged.refresh_seconds = 3 }
        if ($merged.refresh_seconds -gt 60) { $merged.refresh_seconds = 60 }
        if ($merged.check_update_every_hours -lt 1) { $merged.check_update_every_hours = 1 }
        if ($merged.check_update_every_hours -gt 168) { $merged.check_update_every_hours = 168 }
        if ($merged.control_check_update_every_hours -lt 1) { $merged.control_check_update_every_hours = 1 }
        if ($merged.control_check_update_every_hours -gt 168) { $merged.control_check_update_every_hours = 168 }
        if ([string]::IsNullOrWhiteSpace([string]$merged.log_level)) { $merged.log_level = [string]$default.log_level }
        $merged.log_level = $merged.log_level.ToUpperInvariant()
        if ($merged.log_level -notin @('INFO','DEBUG')) { $merged.log_level = 'INFO' }
        if (
            [string]::IsNullOrWhiteSpace([string]$merged.update_url) -or
            [string]$merged.update_url -eq 'https://raw.githubusercontent.com/luizbizzio/tailscale-control/main/tailscale-control.ps1' -or
            [string]$merged.update_url -eq 'https://raw.githubusercontent.com/luizbizzio/tailscale-control/main/install.ps1' -or
            [string]$merged.update_url -match '(?i)raw\.githubusercontent\.com/luizbizzio/tailscale-control/(main|master)/'
        ) { $merged.update_url = [string]$default.update_url }

        $loadedHotkeys = Get-ObjectPropertyOrDefault $loaded 'hotkeys' $null
        $quickCount = Get-QuickAccountSwitchCountFromHotkeys -Hotkeys $loadedHotkeys
        Ensure-QuickAccountSwitchDefinitions -Count $quickCount

        $hotkeyMap = [ordered]@{}
        foreach ($name in $script:HotkeyNames) {
            $fallback = Get-DefaultHotkeyEntry -Name $name
            $source = if ($null -ne $loadedHotkeys) { Get-ObjectPropertyOrDefault $loadedHotkeys $name $null } else { $null }
            if ((Get-QuickAccountSwitchIndex -Name $name) -gt 0) {
                $hotkeyMap[$name] = [pscustomobject]@{
                    enabled = [bool](Get-ObjectPropertyOrDefault $source 'enabled' $fallback.enabled)
                    modifiers = [string](Get-ObjectPropertyOrDefault $source 'modifiers' $fallback.modifiers)
                    key = [string](Get-ObjectPropertyOrDefault $source 'key' $fallback.key)
                    account_identifier = [string](Get-ObjectPropertyOrDefault $source 'account_identifier' (Get-ObjectPropertyOrDefault $fallback 'account_identifier' ''))
                    account_display = [string](Get-ObjectPropertyOrDefault $source 'account_display' (Get-ObjectPropertyOrDefault $fallback 'account_display' ''))
                }
            }
            else {
                $hotkeyMap[$name] = [pscustomobject]@{
                    enabled = [bool](Get-ObjectPropertyOrDefault $source 'enabled' $fallback.enabled)
                    modifiers = [string](Get-ObjectPropertyOrDefault $source 'modifiers' $fallback.modifiers)
                    key = [string](Get-ObjectPropertyOrDefault $source 'key' $fallback.key)
                }
            }
        }
        $merged.hotkeys = [pscustomobject]$hotkeyMap

        $script:ConfigCache = $merged
        return $script:ConfigCache
    }
    catch {
        Write-LogException -Context 'Load config' -ErrorRecord $_
        Ensure-ConfigHotkeyEntries -Config $default -Count $script:QuickAccountSwitchMinimumRows
        Save-Config -Config $default
        return $script:ConfigCache
    }
}

function Save-Config {
    param($Config)
    Initialize-AppRoot
    $script:ConfigCache = $Config
    $Config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Get-MachineColumnOrder {
    @('Machine','Owner','IPv4','IPv6','OS','Connection','DNSName','LastSeen')
}

function Save-MachineColumnLayoutFromGrid {
    if ($script:IsApplyingMachineColumnLayout -or $null -eq $script:gridMachines) { return }
    try {
        $cfg = Get-Config
        $layout = [ordered]@{}
        foreach ($name in Get-MachineColumnOrder) {
            if ($script:gridMachines.Columns.Contains($name)) {
                $col = $script:gridMachines.Columns[$name]
                $value = if ($col.AutoSizeMode -eq [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill -and [double]$col.FillWeight -gt 0) {
                    [int][math]::Round([double]$col.FillWeight)
                }
                else {
                    [int]$col.Width
                }
                if ($value -lt 1) { $value = 1 }
                $layout[$name] = $value
            }
        }
        $cfg.machine_columns = [pscustomobject]$layout
        Save-Config -Config $cfg
    }
    catch {
        Write-Log ('Machine column layout save failed: ' + $_.Exception.Message)
    }
}

function Set-MachineColumnLayout {
    param([switch]$UseConfig,[switch]$PreserveCurrent)
    if ($null -eq $script:gridMachines -or $script:gridMachines.Columns.Count -le 0) { return }

    $mins = Get-MachineColumnMinWidthLayout
    $defaults = Get-DefaultMachineColumnLayout
    $source = if ($PreserveCurrent) {
        $layout = [ordered]@{}
        foreach ($name in Get-MachineColumnOrder) {
            if ($script:gridMachines.Columns.Contains($name)) {
                $col = $script:gridMachines.Columns[$name]
                $layout[$name] = if ($col.AutoSizeMode -eq [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill -and [double]$col.FillWeight -gt 0) {
                    [int][math]::Round([double]$col.FillWeight)
                }
                else {
                    [int]$col.Width
                }
            }
            else {
                $layout[$name] = [int](Get-ObjectPropertyOrDefault $defaults $name 100)
            }
        }
        [pscustomobject]$layout
    }
    elseif ($UseConfig) { (Get-Config).machine_columns }
    else { [pscustomobject](Get-DefaultMachineColumnLayout) }

    $weights = [ordered]@{}
    foreach ($name in Get-MachineColumnOrder) {
        $value = [int](Get-ObjectPropertyOrDefault $source $name (Get-ObjectPropertyOrDefault $defaults $name 100))
        if ($value -lt 1) { $value = 1 }
        $weights[$name] = $value
    }

    $script:IsApplyingMachineColumnLayout = $true
    try {
        $script:gridMachines.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        foreach ($name in Get-MachineColumnOrder) {
            if ($script:gridMachines.Columns.Contains($name)) {
                $col = $script:gridMachines.Columns[$name]
                $col.MinimumWidth = [int](Get-ObjectPropertyOrDefault $mins $name 72)
                $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
                $col.FillWeight = [single]([double]$weights[$name])
            }
        }
    }
    finally {
        $script:IsApplyingMachineColumnLayout = $false
    }
}

function ConvertTo-DnsName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace([string]$Name)) { return '' }
    return ([string]$Name).Trim().TrimEnd('.')
}

function Get-PropertyValue {
    param($Object,[string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        try {
            $prop = $Object.PSObject.Properties[$name]
            if ($null -ne $prop -and $null -ne $prop.Value) { return $prop.Value }
        }
        catch { Write-LogException -Context ('Read optional property ' + $name) -ErrorRecord $_ }
    }
    return $null
}

function Resolve-UserLabel {
    param($UserMap,$UserId)
    if ($null -eq $UserMap -or $null -eq $UserId) { return '' }
    try {
        $key = [string]$UserId
        $entry = $UserMap.PSObject.Properties[$key]
        if ($null -eq $entry -or $null -eq $entry.Value) { return '' }
        $userProfileEntry = $entry.Value
        $display = ConvertTo-DiagnosticText -Text ([string](Get-PropertyValue $userProfileEntry @('DisplayName','LoginName')))
        return $display
    }
    catch { return '' }
}

function Convert-ToNullableBool {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) {
        if ($Value.Count -le 0) { return $null }
        $Value = $Value[0]
    }
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    switch ($text) {
        'true' { return $true }
        'false' { return $false }
        '1' { return $true }
        '0' { return $false }
        'yes' { return $true }
        'no' { return $false }
        'on' { return $true }
        'off' { return $false }
    }
    try { return [System.Convert]::ToBoolean($Value) } catch { return $null }
}

function Convert-ToObjectArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @([string]$Value) }
    if ($Value -is [System.Array]) { return @($Value) }
    if ($Value -is [System.Collections.Generic.List[object]]) { return $Value.ToArray() }
    if ($Value -is [System.Collections.IEnumerable]) {
        $buffer = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$buffer.Add($item) }
        return $buffer.ToArray()
    }
    return @($Value)
}

function ConvertFrom-JsonSafe {
    param(
        [string]$Text,
        [string]$Context = 'JSON parse'
    )
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return $null }
    try {
        return ($Text | ConvertFrom-Json)
    }
    catch {
        Write-LogException -Context $Context -ErrorRecord $_
        return $null
    }
}

function Get-ThemePalette {
    return [ordered]@{
        Theme = 'Light'; FormBack = [System.Drawing.Color]::FromArgb(242,245,249); PanelBack = [System.Drawing.Color]::White; InputBack = [System.Drawing.Color]::White; Border = [System.Drawing.Color]::FromArgb(210,218,231); HeaderBack = [System.Drawing.Color]::FromArgb(38,80,170); HeaderText = [System.Drawing.Color]::White; Text = [System.Drawing.Color]::FromArgb(17,24,39); MutedText = [System.Drawing.Color]::FromArgb(88,102,126); ButtonBack = [System.Drawing.Color]::FromArgb(246,248,252); ButtonText = [System.Drawing.Color]::FromArgb(17,24,39); Accent = [System.Drawing.Color]::FromArgb(59,130,246); SuccessBack = [System.Drawing.Color]::FromArgb(226,243,230); SuccessText = [System.Drawing.Color]::FromArgb(8,88,42); WarnBack = [System.Drawing.Color]::FromArgb(255,244,229); WarnText = [System.Drawing.Color]::FromArgb(146,64,14); OverlayBack = [System.Drawing.Color]::White; OverlayText = [System.Drawing.Color]::FromArgb(17,24,39)
    }
}

function Find-TailscaleExe {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Tailscale\tailscale.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Tailscale\tailscale.exe')
    )
    $cmd = Get-Command tailscale.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if (-not [string]::IsNullOrWhiteSpace([string]$cmd)) { $candidates = @($cmd) + $candidates }
    foreach ($path in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$path) -and (Test-Path -LiteralPath $path)) { return $path }
    }
    return $null
}

function ConvertTo-ProcessArgumentString {
    param([string[]]$Arguments)
    $quoted = New-Object System.Collections.Generic.List[string]
    foreach ($arg in @($Arguments)) {
        $value = [string]$arg
        if ($null -eq $arg) { $value = '' }
        if ($value -match '[\s"]') {
            $escaped = $value -replace '(\*)"', '$1$1\"'
            $escaped = $escaped -replace '(\+)$', '$1$1'
            [void]$quoted.Add('"' + $escaped + '"')
        }
        else {
            [void]$quoted.Add($value)
        }
    }
    return ($quoted -join ' ')
}

function Invoke-External {
    param([string]$FilePath,[string[]]$Arguments)
    $label = $FilePath + ' ' + ($Arguments -join ' ')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $FilePath
        $psi.Arguments = ConvertTo-ProcessArgumentString -Arguments $Arguments
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        try { if ($null -ne $script:Utf8NoBomEncoding) { $psi.StandardOutputEncoding = $script:Utf8NoBomEncoding } else { $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding -ArgumentList $false } } catch { }
        try { if ($null -ne $script:Utf8NoBomEncoding) { $psi.StandardErrorEncoding = $script:Utf8NoBomEncoding } else { $psi.StandardErrorEncoding = New-Object System.Text.UTF8Encoding -ArgumentList $false } } catch { }
        $psi.CreateNoWindow = $true
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()
        $exit = [int]$proc.ExitCode
        $parts = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrWhiteSpace($stdout)) { [void]$parts.Add($stdout.TrimEnd()) }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) { [void]$parts.Add($stderr.TrimEnd()) }
        $text = ConvertTo-DiagnosticText -Text (($parts -join [Environment]::NewLine).TrimEnd())
        $sw.Stop()
        return [pscustomobject]@{ ExitCode = $exit; Output = [string]$text; DurationMs = [double]$sw.Elapsed.TotalMilliseconds }
    }
    catch {
        try { $sw.Stop() } catch { }
        Write-Log ('Run failed: ' + $label + ' => ' + $_.Exception.Message)
        return [pscustomobject]@{ ExitCode = 1; Output = [string]$_.Exception.Message; DurationMs = [double]$sw.Elapsed.TotalMilliseconds }
    }
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-TailscaleCommandExe {
    param([string]$Exe)
    if (-not [string]::IsNullOrWhiteSpace([string]$Exe)) { return $Exe }
    $snapshot = Get-CurrentSnapshot
    if ($null -ne $snapshot -and -not [string]::IsNullOrWhiteSpace([string]$snapshot.Exe)) { return [string]$snapshot.Exe }
    return Find-TailscaleExe
}

function Invoke-TailscaleCommand {
    param([string]$Exe,[string[]]$Arguments)
    $resolvedExe = Get-TailscaleCommandExe -Exe $Exe
    if ([string]::IsNullOrWhiteSpace([string]$resolvedExe)) {
        return [pscustomobject]@{ ExitCode = 1; Output = 'Tailscale executable was not found.' }
    }
    return Invoke-External -FilePath $resolvedExe -Arguments $Arguments
}

function ConvertTo-PlainVersion {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return (($Text -split "`r?`n") | Select-Object -First 1).Trim()
}

function Get-ScriptVersionFromFile {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace([string]$Path) -or -not (Test-Path -LiteralPath $Path)) { return '' }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
        if ($raw -match '\$script:AppVersion\s*=\s*[''"\"]([^''"\"]+)[''"\"]') { return [string]$Matches[1] }
        if ($raw -match '\$AppVersion\s*=\s*[''"\"]([^''"\"]+)[''"\"]') { return [string]$Matches[1] }
        if ($raw -match '\$Version\s*=\s*[''"\"]([^''"\"]+)[''"\"]') { return [string]$Matches[1] }
        if ($raw -match '\$script:Version\s*=\s*[''"\"]([^''"\"]+)[''"\"]') { return [string]$Matches[1] }
    }
    catch { Write-LogException -Context ('Read script version from file: ' + $Path) -ErrorRecord $_ }
    return ''
}

function Read-JsonFile {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace([string]$Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        return (ConvertFrom-JsonSafe -Text (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) -Context ('Read JSON file ' + $Path))
    }
    catch {
        Write-LogException -Context ('Read JSON file: ' + $Path) -ErrorRecord $_
        return $null
    }
}

function Get-TailscaleMtuInfo {
    $result = [ordered]@{ Interface=''; IPv4=''; IPv6='' }
    try {
        $interfaces = @(Get-NetIPInterface -ErrorAction Stop | Where-Object {
            [string]$_.InterfaceAlias -match 'Tailscale'
        })
        foreach ($entry in $interfaces) {
            if ([string]::IsNullOrWhiteSpace([string]$result.Interface)) { $result.Interface = [string]$entry.InterfaceAlias }
            if ([string]$entry.AddressFamily -eq 'IPv4' -and [string]::IsNullOrWhiteSpace([string]$result.IPv4)) {
                $result.IPv4 = [string]$entry.NlMtu
            }
            elseif ([string]$entry.AddressFamily -eq 'IPv6' -and [string]::IsNullOrWhiteSpace([string]$result.IPv6)) {
                $result.IPv6 = [string]$entry.NlMtu
            }
        }
    }
    catch { Write-LogException -Context 'Read adapter MTU info' -ErrorRecord $_ }
    return [pscustomobject]$result
}

function Get-TailscaleMtuAppInfo {
    $programDataRoot = Join-Path $env:ProgramData 'TailscaleMTU'
    $result = [ordered]@{
        Installed = $false
        ScriptPath = ''
        ServiceName = 'TailscaleMTU'
        ServiceInstalled = $false
        ServiceState = 'Not detected'
        OpenPath = ''
        StatusText = 'Not installed'
        ConfigPath = Join-Path $programDataRoot 'config.json'
        StatePath = Join-Path $programDataRoot 'state.json'
        DesiredMtu = ''
        DesiredMtuIPv4 = ''
        DesiredMtuIPv6 = ''
        CheckInterval = ''
        InterfaceMatch = ''
        DetectedInterface = ''
        LastError = ''
        LastResult = ''
        Version = ''
    }
    $candidates = @(
        (Join-Path $env:ProgramData 'TailscaleMTU\Tailscale-MTU.ps1'),
        (Join-Path $env:ProgramData 'TailscaleMTU\Tailscale MTU.ps1')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate)) {
            $result.Installed = $true
            $result.ScriptPath = $candidate
            $result.OpenPath = $candidate
            $result.Version = Get-ScriptVersionFromFile -Path $candidate
            break
        }
    }
    $serviceCandidates = @('TailscaleMTU', 'TailscaleMTUService') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    $lastServiceError = $null
    foreach ($serviceName in $serviceCandidates) {
        try {
            $svc = Get-Service -Name $serviceName -ErrorAction Stop
            if ($null -ne $svc) {
                $result.ServiceInstalled = $true
                $result.ServiceName = [string]$svc.Name
                $result.ServiceState = [string]$svc.Status
                break
            }
        }
        catch {
            $lastServiceError = $_
            $svcMsg = ''
            try { $svcMsg = [string]$_.Exception.Message } catch { $svcMsg = '' }
            if ($svcMsg -notmatch 'Cannot find any service with service name') {
                Write-LogException -Context ('Read Tailscale MTU service state [' + [string]$serviceName + ']') -ErrorRecord $_
            }
        }
    }
    if (-not $result.ServiceInstalled -and $null -ne $lastServiceError) {
        $svcMsg = ''
        try { $svcMsg = [string]$lastServiceError.Exception.Message } catch { $svcMsg = '' }
        if ($svcMsg -notmatch 'Cannot find any service with service name') {
            Write-LogException -Context 'Read Tailscale MTU service state' -ErrorRecord $lastServiceError
        }
    }
    if (-not $result.Installed -and (Test-Path -LiteralPath $programDataRoot)) {
        $result.Installed = $true
        $result.OpenPath = $programDataRoot
    }

    $configJson = Read-JsonFile -Path $result.ConfigPath
    if ($null -ne $configJson) {
        $desiredLegacy = [string](Get-PropertyValue $configJson @('desired_mtu','DesiredMtu','desiredMtu','mtu'))
        $desiredV4 = [string](Get-PropertyValue $configJson @('desired_mtu_ipv4','DesiredMtuIPv4','desiredMtuIPv4','mtu_ipv4'))
        $desiredV6 = [string](Get-PropertyValue $configJson @('desired_mtu_ipv6','DesiredMtuIPv6','desiredMtuIPv6','mtu_ipv6'))

        if ([string]::IsNullOrWhiteSpace([string]$desiredV4) -and -not [string]::IsNullOrWhiteSpace([string]$desiredLegacy)) {
            $desiredV4 = [string]$desiredLegacy
        }
        if ([string]::IsNullOrWhiteSpace([string]$desiredV6) -and -not [string]::IsNullOrWhiteSpace([string]$desiredLegacy)) {
            $desiredV6 = [string]$desiredLegacy
        }

        $result.DesiredMtuIPv4 = [string]$desiredV4
        $result.DesiredMtuIPv6 = [string]$desiredV6

        if (-not [string]::IsNullOrWhiteSpace([string]$desiredLegacy)) {
            $result.DesiredMtu = $desiredLegacy
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredV4) -and -not [string]::IsNullOrWhiteSpace([string]$desiredV6)) {
            if ([string]$desiredV4 -eq [string]$desiredV6) {
                $result.DesiredMtu = [string]$desiredV4
            }
            else {
                $result.DesiredMtu = ('IPv4 ' + [string]$desiredV4 + ' | IPv6 ' + [string]$desiredV6)
            }
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredV4)) {
            $result.DesiredMtu = ('IPv4 ' + [string]$desiredV4)
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredV6)) {
            $result.DesiredMtu = ('IPv6 ' + [string]$desiredV6)
        }

        $result.CheckInterval = [string](Get-PropertyValue $configJson @('check_interval_seconds','CheckIntervalSeconds','checkIntervalSeconds','interval_seconds'))
        $result.InterfaceMatch = [string](Get-PropertyValue $configJson @('interface_match','InterfaceMatch','interface_alias','interfaceAlias'))
        if ([string]::IsNullOrWhiteSpace([string]$result.Version)) {
            $result.Version = [string](Get-PropertyValue $configJson @('version','Version','build','Build','app_version','AppVersion'))
        }
    }

    $stateJson = Read-JsonFile -Path $result.StatePath
    if ($null -ne $stateJson) {
        $result.DetectedInterface = [string](Get-PropertyValue $stateJson @('detected_interface','DetectedInterface','interface','Interface'))
        $result.LastError = [string](Get-PropertyValue $stateJson @('last_error','LastError','error'))
        $result.LastResult = [string](Get-PropertyValue $stateJson @('last_result','LastResult','status','Status'))

        if ([string]::IsNullOrWhiteSpace([string]$result.DesiredMtu)) {
            $desiredStateLegacy = [string](Get-PropertyValue $stateJson @('desired_mtu','DesiredMtu','desiredMtu','mtu'))
            $desiredStateV4 = [string](Get-PropertyValue $stateJson @('desired_mtu_ipv4','DesiredMtuIPv4','desiredMtuIPv4','mtu_ipv4'))
            $desiredStateV6 = [string](Get-PropertyValue $stateJson @('desired_mtu_ipv6','DesiredMtuIPv6','desiredMtuIPv6','mtu_ipv6'))

            if ([string]::IsNullOrWhiteSpace([string]$desiredStateV4) -and -not [string]::IsNullOrWhiteSpace([string]$desiredStateLegacy)) {
                $desiredStateV4 = [string]$desiredStateLegacy
            }
            if ([string]::IsNullOrWhiteSpace([string]$desiredStateV6) -and -not [string]::IsNullOrWhiteSpace([string]$desiredStateLegacy)) {
                $desiredStateV6 = [string]$desiredStateLegacy
            }

            if ([string]::IsNullOrWhiteSpace([string]$result.DesiredMtuIPv4) -and -not [string]::IsNullOrWhiteSpace([string]$desiredStateV4)) {
                $result.DesiredMtuIPv4 = [string]$desiredStateV4
            }
            if ([string]::IsNullOrWhiteSpace([string]$result.DesiredMtuIPv6) -and -not [string]::IsNullOrWhiteSpace([string]$desiredStateV6)) {
                $result.DesiredMtuIPv6 = [string]$desiredStateV6
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$desiredStateLegacy)) {
                $result.DesiredMtu = $desiredStateLegacy
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredStateV4) -and -not [string]::IsNullOrWhiteSpace([string]$desiredStateV6)) {
                if ([string]$desiredStateV4 -eq [string]$desiredStateV6) {
                    $result.DesiredMtu = [string]$desiredStateV4
                }
                else {
                    $result.DesiredMtu = ('IPv4 ' + [string]$desiredStateV4 + ' | IPv6 ' + [string]$desiredStateV6)
                }
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredStateV4)) {
                $result.DesiredMtu = ('IPv4 ' + [string]$desiredStateV4)
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$desiredStateV6)) {
                $result.DesiredMtu = ('IPv6 ' + [string]$desiredStateV6)
            }
        }

        if ([string]::IsNullOrWhiteSpace([string]$result.InterfaceMatch)) {
            $result.InterfaceMatch = [string](Get-PropertyValue $stateJson @('interface_match','InterfaceMatch'))
        }
        if ([string]::IsNullOrWhiteSpace([string]$result.Version)) {
            $result.Version = [string](Get-PropertyValue $stateJson @('version','Version','build','Build','app_version','AppVersion'))
        }
    }

    if ($result.Installed) {
        if ($result.ServiceInstalled) {
            $result.StatusText = 'Installed (' + [string]$result.ServiceState + ')'
        }
        else {
            $result.StatusText = 'Installed'
        }
    }
    return [pscustomobject]$result
}

function Show-ToggleOverlay {
    param([string]$Title,[string]$Message,[string]$Indicator='Info',[switch]$ErrorStyle)
    if ($script:SuppressToggleOverlay) { return }
    $cfg = Get-Config
    if ([bool](Get-ObjectPropertyOrDefault $cfg 'show_toggle_popups' $true)) {
        Show-Overlay -Title $Title -Message $Message -Indicator $Indicator -ErrorStyle:$ErrorStyle
    }
}

function Get-TailscaleDnsInfo {
    param([string]$Exe)
    $basic = Invoke-TailscaleCommand -Exe $Exe -Arguments @('dns','status')
    $advanced = Invoke-TailscaleCommand -Exe $Exe -Arguments @('dns','status','--all')
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($chunk in @($basic.Output,$advanced.Output)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$chunk)) { [void]$parts.Add([string]$chunk) }
    }
    $raw = ($parts -join [Environment]::NewLine)
    $rawLines = @($raw -split "`r?`n")
    $active = New-Object System.Collections.Generic.List[string]
    $ipPattern = '(?<![A-Za-z0-9])((?:\d{1,3}(?:\.\d{1,3}){3})|(?:[0-9A-Fa-f]{0,4}(?::[0-9A-Fa-f]{0,4}){2,}[0-9A-Fa-f:%\.]*))(?![A-Za-z0-9])'
    foreach ($line in $rawLines) {
        $trim = [string]$line
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        foreach ($m in [regex]::Matches($trim,$ipPattern)) {
            $val = ([string]$m.Groups[1].Value).Trim().Trim(',', ';', '|')
            if (-not [string]::IsNullOrWhiteSpace($val) -and -not $active.Contains($val)) {
                [void]$active.Add($val)
            }
        }
    }
    $summaryLines = @()
    foreach ($line in $rawLines) {
        $trim = [string]$line
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        $summaryLines += $trim
        if (@($summaryLines).Count -ge 4) { break }
    }
    return [pscustomobject]@{
        Summary = ($summaryLines -join ' | ')
        Nameservers = $(if (@($active).Count -gt 0) { $active -join ', ' } else { 'System default' })
        Raw = $raw
    }
}

function Test-IsTailscaleDnsAddress {
    param([string]$Address)
    $value = ([string]$Address).Trim().Trim(',', ';', '|')
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    $value = ($value -replace '%.*$','')
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($value, [ref]$ip)) { return $false }
    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        $bytes = $ip.GetAddressBytes()
        if ($bytes.Length -ne 4) { return $false }
        if ($bytes[0] -ne 100) { return $false }
        return ($bytes[1] -ge 64 -and $bytes[1] -le 127)
    }
    $ipv6 = $ip.ToString().ToLowerInvariant()
    return $ipv6.StartsWith('fd7a:115c:a1e0:')
}

function Get-ActiveDnsDisplayText {
    param(
        [string]$Nameservers,
        [object]$CorpDns,
        [switch]$PreferSystem
    )
    $systemDns = Get-SystemDnsNameservers
    if ($PreferSystem -and -not [string]::IsNullOrWhiteSpace([string]$systemDns)) {
        return [string]$systemDns
    }
    $parts = @([string]$Nameservers -split ',') | ForEach-Object { ([string]$_).Trim().Trim(',', ';', '|') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (@($parts).Count -le 0) {
        if ($PreferSystem) { return [string]$systemDns }
        return ''
    }
    $corpDnsEnabled = Convert-ToNullableBool $CorpDns
    if ($corpDnsEnabled) {
        $tailscaleOnly = @($parts | Where-Object { Test-IsTailscaleDnsAddress $_ })
        if (@($tailscaleOnly).Count -gt 0) { $parts = $tailscaleOnly }
        elseif ($PreferSystem) { return [string]$systemDns }
    }
    if (@($parts).Count -eq 1) { return [string]$parts[0] }
    return ($parts[0..([Math]::Min(1, @($parts).Count - 1))] -join ', ')
}

function Get-SystemDnsNameservers {
    $servers = New-Object System.Collections.Generic.List[string]
    try {
        $rows = Get-DnsClientServerAddress -AddressFamily IPv4,IPv6 -ErrorAction Stop
        foreach ($row in @($rows)) {
            foreach ($addr in @(Convert-ToObjectArray $row.ServerAddresses)) {
                $val = ([string]$addr).Trim().Trim(',', ';', '|')
                if ([string]::IsNullOrWhiteSpace($val)) { continue }
                if ($val -match '^(?i)fec0:0:0:ffff::[1-3](?:%.*)?$') { continue }
                try { if (Test-IsTailscaleVirtualDnsProxyAddress -Address $val) { continue } } catch { }
                if (-not $servers.Contains($val)) {
                    [void]$servers.Add($val)
                }
            }
        }
    }
    catch { Write-LogException -Context 'Parse DNS resolvers in use' -ErrorRecord $_ }
    if (@($servers).Count -le 0) { return 'System default' }
    return ($servers -join ', ')
}

function Resolve-ExitNodeDisplay {
    param($Snapshot)
    $current = ConvertTo-DnsName ([string]$Snapshot.CurrentExitNode)
    if ([string]::IsNullOrWhiteSpace($current)) { return 'None' }
    foreach ($m in @(Convert-ToObjectArray $Snapshot.Machines)) {
        if ($null -eq $m) { continue }
        $machineName = ConvertTo-DnsName ([string]$m.Machine)
        $dns = ConvertTo-DnsName ([string]$m.DNSName)
        $ipv4 = ConvertTo-DnsName ([string]$m.IPv4)
        $ipv6 = ConvertTo-DnsName ([string]$m.IPv6)
        $deviceId = ConvertTo-DnsName ([string]$m.DeviceId)
        if ($current -eq $machineName -or $current -eq $dns -or $current -eq $ipv4 -or $current -eq $ipv6 -or $current -eq $deviceId) {
            return [string]$m.Machine
        }
    }
    return $current
}

function Resolve-ExitNodeInfo {
    param($Snapshot)
    $current = ConvertTo-DnsName ([string]$Snapshot.CurrentExitNode)
    if ([string]::IsNullOrWhiteSpace($current)) {
        return [pscustomobject]@{ Found = $false; Name = ''; DNSName = ''; IPv4 = ''; IPv6 = ''; ID = ''; Label = '' }
    }
    foreach ($m in @(Convert-ToObjectArray $Snapshot.Machines)) {
        if ($null -eq $m) { continue }
        $machineName = ConvertTo-DnsName ([string]$m.Machine)
        $dns = ConvertTo-DnsName ([string]$m.DNSName)
        $ipv4 = ConvertTo-DnsName ([string]$m.IPv4)
        $ipv6 = ConvertTo-DnsName ([string]$m.IPv6)
        $deviceId = ConvertTo-DnsName ([string]$m.DeviceId)
        if ($current -ne $machineName -and $current -ne $dns -and $current -ne $ipv4 -and $current -ne $ipv6 -and $current -ne $deviceId) { continue }
        $label = $machineName
        if ([string]::IsNullOrWhiteSpace($label)) { $label = $dns }
        if ([string]::IsNullOrWhiteSpace($label)) { $label = $ipv4 }
        if ([string]::IsNullOrWhiteSpace($label)) { $label = $current }
        return [pscustomobject]@{ Found = $true; Name = $machineName; DNSName = $dns; IPv4 = $ipv4; IPv6 = $ipv6; ID = $deviceId; Label = $label }
    }
    return [pscustomobject]@{ Found = $false; Name = $current; DNSName = ''; IPv4 = ''; IPv6 = ''; ID = $current; Label = $current }
}

function Resolve-ExitNodeDetailedDisplay {
    param($Snapshot)
    $info = Resolve-ExitNodeInfo -Snapshot $Snapshot
    if ($null -eq $info -or [string]::IsNullOrWhiteSpace([string]$info.Label)) { return '' }
    $main = [string]$info.DNSName
    if ([string]::IsNullOrWhiteSpace($main)) { $main = [string]$info.Name }
    if ([string]::IsNullOrWhiteSpace($main)) { $main = [string]$info.IPv4 }
    if ([string]::IsNullOrWhiteSpace($main)) { $main = [string]$info.Label }
    $extra = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace([string]$info.IPv4) -and [string]$info.IPv4 -ne $main) { [void]$extra.Add([string]$info.IPv4) }
    if (-not [string]::IsNullOrWhiteSpace([string]$info.IPv6) -and [string]$info.IPv6 -ne $main) { [void]$extra.Add([string]$info.IPv6) }
    if (-not [string]::IsNullOrWhiteSpace([string]$info.ID) -and [string]$info.ID -ne $main) { [void]$extra.Add('id: ' + [string]$info.ID) }
    if ($extra.Count -gt 0) { return ($main + ' (' + ([string]::Join(', ', $extra.ToArray())) + ')') }
    return $main
}
function Set-LastClientUpdateCheckTimestamp {
    param([datetime]$Timestamp = ([datetime]::UtcNow))
    $cfg = Get-Config
    $cfg.last_client_update_check_utc = $Timestamp.ToString('o')
    Save-Config -Config $cfg
}

function Set-LastClientUpdateLatestVersion {
    param([string]$Version)
    $cfg = Get-Config
    $cfg.last_client_update_latest_version = [string]$Version
    Save-Config -Config $cfg
}

function Set-LastClientUpdateStatus {
    param([string]$Status)
    $cfg = Get-Config
    $cfg.last_client_update_status = [string]$Status
    Save-Config -Config $cfg
}

function Convert-ToComparableVersion {
    param([string]$VersionText)
    $clean = [string]$VersionText
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    if ($clean -match '(\d+(?:\.\d+){1,3})') { $clean = $Matches[1] }
    try { return [version]$clean } catch { return $null }
}

function Test-VersionNewer {
    param([string]$CurrentVersion,[string]$LatestVersion)
    $cur = Convert-ToComparableVersion $CurrentVersion
    $lat = Convert-ToComparableVersion $LatestVersion
    if ($null -eq $cur -or $null -eq $lat) { return $false }
    return ($lat -gt $cur)
}
function Update-TailscaleClientUpdateUi {
    param($Snapshot)
    $cfg = Get-Config
    $currentVersion = if ($null -ne $Snapshot) { ConvertTo-PlainVersion ([string]$Snapshot.Version) } else { '' }
    $latestVersion = [string](Get-ObjectPropertyOrDefault $cfg 'last_client_update_latest_version' '')
    $lastStatus = [string](Get-ObjectPropertyOrDefault $cfg 'last_client_update_status' '')
    $lastCheckUtc = [string](Get-ObjectPropertyOrDefault $cfg 'last_client_update_check_utc' '')
    $lastCheckDisplay = '-'
    if (-not [string]::IsNullOrWhiteSpace($lastCheckUtc)) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($lastCheckUtc, [ref]$parsed)) {
            $lastCheckDisplay = $parsed.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
        }
        else { $lastCheckDisplay = $lastCheckUtc }
    }
    if ($null -ne $script:lblMaintLocalVersion) { $script:lblMaintLocalVersion.Text = $(if ([string]::IsNullOrWhiteSpace($currentVersion)) { '-' } else { $currentVersion }) }
    if ($null -ne $script:lblMaintLatestVersion) { $script:lblMaintLatestVersion.Text = $(if ([string]::IsNullOrWhiteSpace($latestVersion)) { '-' } else { $latestVersion }) }
    if ($null -ne $script:lblMaintLastCheck) { $script:lblMaintLastCheck.Text = $lastCheckDisplay }
    $hasUpdate = Test-VersionNewer -CurrentVersion $currentVersion -LatestVersion $latestVersion
    if (-not $hasUpdate -and $lastStatus -match '(?i)new version available|update available') { $hasUpdate = $true }
    $taskReady = Test-TailscaleClientElevatedTasksReady
    if ($null -ne $script:btnRunClientUpdate) {
        $script:btnRunClientUpdate.Enabled = ($hasUpdate -and $taskReady -and -not $script:IsClientTaskSetupRunning -and -not $script:IsClientMaintenanceTaskRunning)
        $script:btnRunClientUpdate.Text = 'Update'
    }
    Update-TailscaleClientTaskSetupUi
    if ($null -ne $script:lblMaintUpdateStatus) {
        if ([string]::IsNullOrWhiteSpace([string]$latestVersion)) {
            if ([string]::IsNullOrWhiteSpace([string]$lastStatus)) { $script:lblMaintUpdateStatus.Text = 'Check for updates to detect the latest version.' }
            else { $script:lblMaintUpdateStatus.Text = [string]$lastStatus }
        }
        elseif ($hasUpdate) { $script:lblMaintUpdateStatus.Text = 'Update available.' }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$lastStatus)) { $script:lblMaintUpdateStatus.Text = [string]$lastStatus }
        else { $script:lblMaintUpdateStatus.Text = 'Already up to date.' }
    }
}

function Set-LastControlUpdateCheckTimestamp {
    param([datetime]$Timestamp = ([datetime]::UtcNow))
    $cfg = Get-Config
    $cfg.last_control_update_check_utc = $Timestamp.ToString('o')
    Save-Config -Config $cfg
}

function Set-LastControlUpdateLatestVersion {
    param([string]$Version)
    $cfg = Get-Config
    $cfg.last_control_update_latest_version = [string]$Version
    Save-Config -Config $cfg
}

function Set-LastControlUpdateStatus {
    param([string]$Status)
    $cfg = Get-Config
    $cfg.last_control_update_status = [string]$Status
    Save-Config -Config $cfg
}
function Update-ControlAppUpdateUi {
    $cfg = Get-Config
    $currentVersion = [string]$script:AppVersion
    $latestVersion = [string](Get-ObjectPropertyOrDefault $cfg 'last_control_update_latest_version' '')
    $lastCheckUtc = [string](Get-ObjectPropertyOrDefault $cfg 'last_control_update_check_utc' '')
    $lastStatus = [string](Get-ObjectPropertyOrDefault $cfg 'last_control_update_status' '')
    $autoEnabled = [bool](Get-ObjectPropertyOrDefault $cfg 'control_check_update_every_enabled' $false)
    $lastCheckDisplay = '-'
    if (-not [string]::IsNullOrWhiteSpace($lastCheckUtc)) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($lastCheckUtc, [ref]$parsed)) { $lastCheckDisplay = $parsed.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { $lastCheckDisplay = $lastCheckUtc }
    }
    if ($null -ne $script:lblControlVersion) { $script:lblControlVersion.Text = [string]$currentVersion }
    if ($null -ne $script:lblControlLatestVersion) { $script:lblControlLatestVersion.Text = $(if ([string]::IsNullOrWhiteSpace($latestVersion)) { '-' } else { [string]$latestVersion }) }
    if ($null -ne $script:lblControlLastCheck) { $script:lblControlLastCheck.Text = [string]$lastCheckDisplay }
    if ($null -ne $script:lblControlAutoUpdate) { $script:lblControlAutoUpdate.Text = $(if ($autoEnabled) { 'On' } else { 'Off' }) }
    if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = $(if ([string]::IsNullOrWhiteSpace($lastStatus)) { '-' } else { [string]$lastStatus }) }
    $hasUpdate = $false
    if (-not [string]::IsNullOrWhiteSpace([string]$latestVersion)) {
        $hasUpdate = Test-VersionNewer -CurrentVersion $currentVersion -LatestVersion $latestVersion
    }
    if ($null -ne $script:btnUpdate -and -not $script:IsControlMaintenanceTaskRunning) { $script:btnUpdate.Enabled = $hasUpdate }
    if ($null -ne $script:btnCheckControlUpdate -and -not $script:IsControlMaintenanceTaskRunning) {
        $script:btnCheckControlUpdate.Enabled = $true
        $script:btnCheckControlUpdate.Text = 'Check Update'
    }
}

function Set-TailscaleControlMaintenanceBusyState {
    param([bool]$Busy,[string]$Operation)
    try {
        if ($null -ne $script:btnCheckControlUpdate) {
            $script:btnCheckControlUpdate.Text = $(if ($Busy -and ($Operation -eq 'Check' -or $Operation -eq 'AutoUpdate')) { 'Checking...' } else { 'Check Update' })
            $script:btnCheckControlUpdate.Enabled = $(if ($Busy) { $false } else { $true })
            if ($Busy -and ($Operation -eq 'Check' -or $Operation -eq 'AutoUpdate')) { Clear-UiFocusSoon }
        }
        if ($null -ne $script:btnUpdate) {
            if ($Busy) {
                $script:btnUpdate.Enabled = $(if ($Operation -eq 'Update') { $true } else { $false })
                $script:btnUpdate.Text = $(if ($Operation -eq 'Update') { 'Updating...' } else { 'Update' })
                if ($Operation -eq 'Update') { Set-ControlFocusSafe -Control $script:btnUpdate }
            }
            else {
                $script:btnUpdate.Text = 'Update'
                $currentVersion = [string]$script:AppVersion
                $latestVersion = ''
                try { $latestVersion = [string](Get-ObjectPropertyOrDefault (Get-Config) 'last_control_update_latest_version' '') } catch { }
                $script:btnUpdate.Enabled = (Test-VersionNewer -CurrentVersion $currentVersion -LatestVersion $latestVersion)
            }
        }
    }
    catch {
        Write-LogException -Context 'Set Tailscale Control maintenance busy state' -ErrorRecord $_
    }
}

function Start-TailscaleControlMaintenanceTask {
    param(
        [ValidateSet('Check','Update','AutoUpdate')][string]$Operation,
        [switch]$Scheduled
    )
    if ($script:IsControlMaintenanceTaskRunning) { return }
    $script:IsControlMaintenanceTaskRunning = $true
    Set-TailscaleControlMaintenanceBusyState -Busy $true -Operation $Operation
    try {
        if ($Operation -eq 'Update') {
            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Updating...' }
        }
        elseif ($Operation -eq 'AutoUpdate') {
            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Checking for automatic update...' }
        }
        else {
            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Checking for updates...' }
        }
    }
    catch { }

    $taskId = [guid]::NewGuid().ToString('N')
    $resultPath = Join-Path $env:TEMP ('tailscale-control-self-maintenance-' + $taskId + '.json')
    $runnerPath = Join-Path $env:TEMP ('tailscale-control-self-maintenance-' + $taskId + '.ps1')
    $cfg = Get-Config
    $updateUrl = [string](Get-ObjectPropertyOrDefault $cfg 'update_url' '')
    $payload = [pscustomobject]@{
        Operation = [string]$Operation
        CurrentVersion = [string]$script:AppVersion
        UpdateUrl = [string]$updateUrl
        ResultPath = [string]$resultPath
        PowerShellExe = [string]$script:PowerShellExe
    }
    $payloadJson = $payload | ConvertTo-Json -Depth 5 -Compress
    $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payloadJson))

    $runner = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'
`$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
`$payload = `$payloadJson | ConvertFrom-Json

function Write-ResultFile {
    param(`$Object)
    `$json = `$Object | ConvertTo-Json -Depth 10 -Compress
    [IO.File]::WriteAllText([string]`$payload.ResultPath, `$json, [Text.Encoding]::UTF8)
}

function Convert-ToComparableVersionLocal {
    param([string]`$VersionText)
    `$clean = [string]`$VersionText
    if ([string]::IsNullOrWhiteSpace(`$clean)) { return `$null }
    if (`$clean -match '(\d+(?:\.\d+){1,3})') { `$clean = `$Matches[1] }
    try { return [version]`$clean } catch { return `$null }
}

function Test-VersionNewerLocal {
    param([string]`$CurrentVersion,[string]`$LatestVersion)
    `$cur = Convert-ToComparableVersionLocal `$CurrentVersion
    `$lat = Convert-ToComparableVersionLocal `$LatestVersion
    if (`$null -eq `$cur -or `$null -eq `$lat) { return `$false }
    return (`$lat -gt `$cur)
}

function Get-ControlLatestAvailableVersionLocal {
    param([string]`$Url)
    if ([string]::IsNullOrWhiteSpace(`$Url)) {
        return [pscustomobject]@{ Success = `$false; Version = ''; Status = 'Update URL is not configured.' }
    }
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
        `$req = [Net.HttpWebRequest]::Create(`$Url)
        `$req.Method = 'GET'
        `$req.Timeout = 12000
        `$req.ReadWriteTimeout = 12000
        `$req.UserAgent = 'TailscaleControl'
        `$resp = `$req.GetResponse()
        try {
            `$stream = `$resp.GetResponseStream()
            `$reader = New-Object IO.StreamReader(`$stream)
            `$raw = [string]`$reader.ReadToEnd()
        }
        finally {
            try { `$reader.Dispose() } catch { }
            try { `$resp.Dispose() } catch { }
        }
        if (`$raw -match '\`$script:AppVersion\s*=\s*[''\"]([^''\"]+)[''\"]') {
            return [pscustomobject]@{ Success = `$true; Version = [string]`$Matches[1]; Status = 'Version file read successfully.' }
        }
        if (`$raw -match '\`$AppVersion\s*=\s*[''\"]([^''\"]+)[''\"]') {
            return [pscustomobject]@{ Success = `$true; Version = [string]`$Matches[1]; Status = 'Version file read successfully.' }
        }
        return [pscustomobject]@{ Success = `$false; Version = ''; Status = 'AppVersion was not found in the remote file.' }
    }
    catch {
        return [pscustomobject]@{ Success = `$false; Version = ''; Status = ('Could not read the remote version file: ' + [string]`$_.Exception.Message) }
    }
}

function Get-ControlUpdateCheckLocal {
    param([string]`$CurrentVersion,[string]`$Url)
    `$sw = [Diagnostics.Stopwatch]::StartNew()
    `$remote = Get-ControlLatestAvailableVersionLocal -Url `$Url
    `$latest = ''
    `$hasUpdate = `$false
    `$status = 'Unable to determine latest version.'
    if (`$null -ne `$remote -and [bool]`$remote.Success) {
        `$latest = [string]`$remote.Version
        `$hasUpdate = Test-VersionNewerLocal -CurrentVersion `$CurrentVersion -LatestVersion `$latest
        if ([string]::IsNullOrWhiteSpace(`$latest)) { `$status = 'Unable to determine latest version.' }
        elseif (`$hasUpdate) { `$status = 'Update available.' }
        else { `$status = 'Already up to date.' }
    }
    elseif (`$null -ne `$remote) {
        `$status = [string]`$remote.Status
    }
    `$sw.Stop()
    return [pscustomobject]@{ CurrentVersion = [string]`$CurrentVersion; LatestVersion = [string]`$latest; LastCheckUtc = ([datetime]::UtcNow).ToString('o'); HasUpdate = [bool]`$hasUpdate; Status = [string]`$status; DurationMs = [double]`$sw.Elapsed.TotalMilliseconds; Command = 'Check latest Tailscale Control release' }
}

function Start-ControlUpdaterLocal {
    param([string]`$Url)
    `$sw = [Diagnostics.Stopwatch]::StartNew()
    if ([string]::IsNullOrWhiteSpace(`$Url)) { throw 'Update URL is not configured yet.' }
    `$tempPath = Join-Path `$env:TEMP ('tailscale-control-update-' + [guid]::NewGuid().ToString('N') + '.ps1')
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
        `$wc = New-Object Net.WebClient
        `$wc.Headers['User-Agent'] = 'TailscaleControl'
        `$wc.DownloadFile(`$Url, `$tempPath)
        `$launcher = [string]`$payload.PowerShellExe
        if ([string]::IsNullOrWhiteSpace(`$launcher)) { `$launcher = Join-Path `$env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe' }
        `$psi = New-Object Diagnostics.ProcessStartInfo
        `$psi.FileName = `$launcher
        `$psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (`$tempPath -replace '"','\"') + '"'
        `$psi.UseShellExecute = `$false
        `$psi.CreateNoWindow = `$true
        `$psi.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        [void][Diagnostics.Process]::Start(`$psi)
        `$sw.Stop()
        return [pscustomobject]@{ StatusText = 'Updater downloaded and launched.'; DurationMs = [double]`$sw.Elapsed.TotalMilliseconds; TempPath = [string]`$tempPath }
    }
    finally {
        try { if (`$null -ne `$wc) { `$wc.Dispose() } } catch { }
    }
}

try {
    `$op = [string]`$payload.Operation
    `$currentVersion = [string]`$payload.CurrentVersion
    `$url = [string]`$payload.UpdateUrl
    if (`$op -eq 'Check') {
        `$check = Get-ControlUpdateCheckLocal -CurrentVersion `$currentVersion -Url `$url
        Write-ResultFile ([pscustomobject]@{ Operation = 'Check'; Success = `$true; Check = `$check; ErrorMessage = '' })
    }
    elseif (`$op -eq 'AutoUpdate') {
        `$check = Get-ControlUpdateCheckLocal -CurrentVersion `$currentVersion -Url `$url
        `$update = `$null
        if (`$check.HasUpdate) { `$update = Start-ControlUpdaterLocal -Url `$url }
        Write-ResultFile ([pscustomobject]@{ Operation = 'AutoUpdate'; Success = `$true; Check = `$check; UpdateStarted = [bool]`$check.HasUpdate; UpdateResult = `$update; ErrorMessage = '' })
    }
    else {
        `$update = Start-ControlUpdaterLocal -Url `$url
        Write-ResultFile ([pscustomobject]@{ Operation = 'Update'; Success = `$true; UpdateResult = `$update; ErrorMessage = '' })
    }
}
catch {
    Write-ResultFile ([pscustomobject]@{ Operation = [string]`$payload.Operation; Success = `$false; ErrorMessage = [string]`$_.Exception.Message })
    exit 1
}
"@

    try {
        Set-Content -LiteralPath $runnerPath -Value $runner -Encoding UTF8 -Force
        $runnerPathArg = '"' + ([string]$runnerPath -replace '"','\"') + '"'
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $script:PowerShellExe
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ' + $runnerPathArg
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $process = [System.Diagnostics.Process]::Start($psi)
        if ($null -eq $process) { throw 'Could not start Tailscale Control maintenance worker process.' }

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 350
        $task = [pscustomobject]@{
            Operation = [string]$Operation
            Scheduled = [bool]$Scheduled
            Process = $process
            ResultPath = [string]$resultPath
            RunnerPath = [string]$runnerPath
            Timer = $timer
            StartedUtc = [datetime]::UtcNow
        }
        $script:ControlMaintenanceWorker = $task

        $timer.add_Tick({
            try {
                $taskRef = $script:ControlMaintenanceWorker
                if ($null -eq $taskRef) { return }
                $proc = $taskRef.Process
                if ($null -eq $proc -or -not $proc.HasExited) { return }
                try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
                $operationName = [string]$taskRef.Operation
                $result = $null
                try {
                    if (Test-Path -LiteralPath ([string]$taskRef.ResultPath)) {
                        $raw = Get-Content -LiteralPath ([string]$taskRef.ResultPath) -Raw -Encoding UTF8 -ErrorAction Stop
                        if (-not [string]::IsNullOrWhiteSpace([string]$raw)) { $result = $raw | ConvertFrom-Json }
                    }
                }
                catch {
                    $result = [pscustomobject]@{ Operation = $operationName; Success = $false; ErrorMessage = 'Could not read Tailscale Control maintenance result: ' + $_.Exception.Message }
                }
                if ($null -eq $result) {
                    $exitCode = $null
                    try { $exitCode = [int]$proc.ExitCode } catch { $exitCode = -1 }
                    $workerOutput = ''
                    try {
                        $outText = [string]$proc.StandardOutput.ReadToEnd()
                        $errText = [string]$proc.StandardError.ReadToEnd()
                        $workerOutput = (($outText, $errText) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
                    }
                    catch { }
                    $message = 'Tailscale Control maintenance process ended without a result file. Exit code: ' + $exitCode
                    if (-not [string]::IsNullOrWhiteSpace([string]$workerOutput)) { $message += [Environment]::NewLine + $workerOutput }
                    $result = [pscustomobject]@{ Operation = $operationName; Success = $false; ErrorMessage = $message }
                }
                try { $proc.Dispose() } catch { }
                try { if (Test-Path -LiteralPath ([string]$taskRef.ResultPath)) { Remove-Item -LiteralPath ([string]$taskRef.ResultPath) -Force -ErrorAction SilentlyContinue } } catch { }
                try { if (Test-Path -LiteralPath ([string]$taskRef.RunnerPath)) { Remove-Item -LiteralPath ([string]$taskRef.RunnerPath) -Force -ErrorAction SilentlyContinue } } catch { }

                try {
                    $operationName = [string](Get-ObjectPropertyOrDefault $result 'Operation' $operationName)
                    $success = [bool](Get-ObjectPropertyOrDefault $result 'Success' $false)
                    if (-not $success) {
                        $message = [string](Get-ObjectPropertyOrDefault $result 'ErrorMessage' 'Task failed.')
                        if ($operationName -eq 'Update') {
                            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Update failed: ' + $message }
                            Write-ActivityFailureBlock -Title 'Tailscale Control Update failed' -CommandText 'Download and run updater' -Message $message
                        }
                        else {
                            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Check failed: ' + $message }
                            Write-ActivityFailureBlock -Title 'Tailscale Control Check Update failed' -CommandText 'Check latest Tailscale Control release' -Message $message
                        }
                        return
                    }

                    if ($operationName -eq 'Check' -or $operationName -eq 'AutoUpdate') {
                        $check = Get-ObjectPropertyOrDefault $result 'Check' $null
                        if ($null -ne $check) {
                            $latest = [string](Get-ObjectPropertyOrDefault $check 'LatestVersion' '')
                            $lastCheck = [string](Get-ObjectPropertyOrDefault $check 'LastCheckUtc' '')
                            $parsedDate = [datetime]::MinValue
                            if ([datetime]::TryParse($lastCheck, [ref]$parsedDate)) { Set-LastControlUpdateCheckTimestamp -Timestamp $parsedDate.ToUniversalTime() }
                            else { Set-LastControlUpdateCheckTimestamp }
                            Set-LastControlUpdateLatestVersion -Version $latest
                            Set-LastControlUpdateStatus -Status ([string](Get-ObjectPropertyOrDefault $check 'Status' ''))
                            Write-VersionCheckActivity -Title 'Tailscale Control Check Update' -CommandText 'Check latest Tailscale Control version' -Check $check
                        }
                        if ($operationName -eq 'AutoUpdate' -and [bool](Get-ObjectPropertyOrDefault $result 'UpdateStarted' $false)) {
                            $updateResult = Get-ObjectPropertyOrDefault $result 'UpdateResult' $null
                            $statusText = if ($null -ne $updateResult) { [string](Get-ObjectPropertyOrDefault $updateResult 'StatusText' 'Updater downloaded and launched.') } else { 'Updater downloaded and launched.' }
                            $duration = if ($null -ne $updateResult) { [double](Get-ObjectPropertyOrDefault $updateResult 'DurationMs' 0) } else { 0 }
                            Set-LastControlUpdateStatus -Status $statusText
                            Write-ActivityCommandBlock -Title 'Tailscale Control Auto Update' -CommandText 'Download and run updater' -ExitCode 0 -Output $statusText -DurationMs $duration
                            if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = $statusText }
                        }
                        elseif ($null -ne $script:lblControlUpdateStatus -and $null -ne $check) {
                            $script:lblControlUpdateStatus.Text = [string](Get-ObjectPropertyOrDefault $check 'Status' '')
                        }
                        Update-ControlAppUpdateUi
                    }
                    else {
                        $updateResult = Get-ObjectPropertyOrDefault $result 'UpdateResult' $null
                        $statusText = if ($null -ne $updateResult) { [string](Get-ObjectPropertyOrDefault $updateResult 'StatusText' 'Updater downloaded and launched.') } else { 'Updater downloaded and launched.' }
                        $duration = if ($null -ne $updateResult) { [double](Get-ObjectPropertyOrDefault $updateResult 'DurationMs' 0) } else { 0 }
                        Set-LastControlUpdateStatus -Status $statusText
                        Write-ActivityCommandBlock -Title 'Tailscale Control Update' -CommandText 'Download and run updater' -ExitCode 0 -Output $statusText -DurationMs $duration
                        if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = $statusText }
                    }
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if ($operationName -eq 'Update') {
                        if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Update failed: ' + $message }
                        Write-ActivityFailureBlock -Title 'Tailscale Control Update failed' -CommandText 'Download and run updater' -Message $message
                    }
                    else {
                        if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Check failed: ' + $message }
                        Write-ActivityFailureBlock -Title 'Tailscale Control Check Update failed' -CommandText 'Check latest Tailscale Control version' -Message $message
                    }
                }
                finally {
                    $script:IsControlMaintenanceTaskRunning = $false
                    $script:ControlMaintenanceWorker = $null
                    Set-TailscaleControlMaintenanceBusyState -Busy $false -Operation $operationName
                }
            }
            catch {
                $operationName = 'Check'
                try { if ($null -ne $script:ControlMaintenanceWorker) { $operationName = [string]$script:ControlMaintenanceWorker.Operation } } catch { }
                $script:IsControlMaintenanceTaskRunning = $false
                $script:ControlMaintenanceWorker = $null
                Set-TailscaleControlMaintenanceBusyState -Busy $false -Operation $operationName
                if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Task failed: ' + $_.Exception.Message }
                Write-ActivityFailureBlock -Title ('Tailscale Control ' + $operationName + ' failed') -CommandText ('Tailscale Control ' + $operationName) -Message $_.Exception.Message
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsControlMaintenanceTaskRunning = $false
        $script:ControlMaintenanceWorker = $null
        Set-TailscaleControlMaintenanceBusyState -Busy $false -Operation $Operation
        try { if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue } } catch { }
        try { if (Test-Path -LiteralPath $runnerPath) { Remove-Item -LiteralPath $runnerPath -Force -ErrorAction SilentlyContinue } } catch { }
        if ($null -ne $script:lblControlUpdateStatus) { $script:lblControlUpdateStatus.Text = 'Task failed: ' + $_.Exception.Message }
        Write-ActivityFailureBlock -Title ('Tailscale Control ' + $Operation + ' failed') -CommandText ('Tailscale Control ' + $Operation) -Message $_.Exception.Message
    }
}

function Invoke-ScheduledControlUpdate {
    if ($script:IsBusy -or $script:IsRefreshing -or $script:IsControlMaintenanceTaskRunning) { return }
    if ($null -eq $script:chkControlCheckUpdateEvery -or $null -eq $script:numControlCheckUpdateHours) { return }
    if (-not [bool]$script:chkControlCheckUpdateEvery.Checked) { return }
    $cfg = Get-Config
    if (-not [bool](Get-ObjectPropertyOrDefault $cfg 'control_check_update_every_enabled' $false)) { return }
    $hours = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $cfg 'control_check_update_every_hours' 24) 24
    if ($hours -lt 1) { $hours = 1 }
    $lastText = [string](Get-ObjectPropertyOrDefault $cfg 'last_control_update_check_utc' '')
    $shouldRun = $true
    if (-not [string]::IsNullOrWhiteSpace([string]$lastText)) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($lastText, [ref]$parsed)) {
            $elapsed = ([datetime]::UtcNow - $parsed.ToUniversalTime()).TotalHours
            if ($elapsed -lt [double]$hours) { $shouldRun = $false }
        }
    }
    if (-not $shouldRun) { return }
    Start-TailscaleControlMaintenanceTask -Operation 'AutoUpdate' -Scheduled
}

function Invoke-ScheduledClientUpdate {
    if ($script:IsBusy -or $script:IsRefreshing -or $script:IsClientMaintenanceTaskRunning) { return }
    if ($null -eq $script:chkCheckUpdateEvery -or $null -eq $script:numCheckUpdateHours) { return }
    if (-not [bool]$script:chkCheckUpdateEvery.Checked) { return }
    $cfg = Get-Config
    if (-not [bool](Get-ObjectPropertyOrDefault $cfg 'check_update_every_enabled' $false)) { return }
    $hours = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $cfg 'check_update_every_hours' 24) 24
    if ($hours -lt 1) { $hours = 1 }
    $lastText = [string](Get-ObjectPropertyOrDefault $cfg 'last_client_update_check_utc' '')
    $shouldRun = $true
    if (-not [string]::IsNullOrWhiteSpace([string]$lastText)) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($lastText, [ref]$parsed)) {
            $elapsed = ([datetime]::UtcNow - $parsed.ToUniversalTime()).TotalHours
            if ($elapsed -lt [double]$hours) { $shouldRun = $false }
        }
    }
    if (-not $shouldRun) { return }
    if (-not (Test-TailscaleClientElevatedTasksReady)) {
        if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Auto update skipped: elevated task is not installed.' }
        return
    }
    Start-TailscaleClientMaintenanceTask -Operation 'AutoUpdate' -Scheduled
}

function Install-TailscaleMtu {
    param($Button = $null)
    if ($null -eq $Button) { $Button = $script:btnInstallMtu }
    $originalText = ''
    try { if ($null -ne $Button) { $originalText = [string]$Button.Text; $Button.Text = 'Installing...'; $Button.Enabled = $false } } catch { }
    $script:IsMtuInstallRunning = $true
    try {
        $cfg = Get-Config
        $defaultCfg = Get-DefaultConfig
        $url = [string](Get-ObjectPropertyOrDefault $cfg 'mtu_install_url' $defaultCfg.mtu_install_url)
        if ([string]::IsNullOrWhiteSpace($url)) { throw 'MTU install URL is not configured.' }
        $tempPath = Join-Path $env:TEMP 'tailscale-mtu-install.ps1'
        Invoke-WebRequest -Uri $url -OutFile $tempPath -UseBasicParsing
        $argLine = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $tempPath + '" -Install'
        $process = Start-Process -FilePath $script:PowerShellExe -ArgumentList $argLine -WindowStyle Hidden -PassThru
        Write-ActivityCommandBlock -Title 'Install Tailscale MTU' -CommandText 'Install Tailscale MTU' -ExitCode 0 -Output 'The installer was downloaded and launched.'
        Show-Overlay -Title 'Tailscale MTU install started' -Message 'The installer was downloaded and launched.' -Indicator 'Info'
        $state = [pscustomobject]@{ Process = $process; Button = $Button; OriginalText = $originalText; Attempts = 0 }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $timer.add_Tick({
            param($sender,$eventArgs)
            try {
                $state.Attempts = [int]$state.Attempts + 1
                $done = $false
                try { $done = ($null -eq $state.Process -or $state.Process.HasExited) } catch { $done = $true }
                if (-not $done -and [int]$state.Attempts -lt 240) { return }
                try { $sender.Stop(); $sender.Dispose() } catch { }
                try { if ($null -ne $state.Process) { $state.Process.Dispose() } } catch { }
                $script:IsMtuInstallRunning = $false
                try { Reset-SlowSnapshotCache } catch { }
                try { Update-Status } catch { Write-LogException -Context 'Refresh after MTU install' -ErrorRecord $_ }
                if ($null -ne $state.Button) {
                    try {
                        $info = Get-TailscaleMtuAppInfo
                        if (-not [bool]$info.Installed) {
                            $state.Button.Text = $(if ([string]::IsNullOrWhiteSpace([string]$state.OriginalText)) { 'Install Tailscale MTU' } else { [string]$state.OriginalText })
                            $state.Button.Enabled = $true
                        }
                    } catch { }
                }
            }
            catch {
                try { $sender.Stop(); $sender.Dispose() } catch { }
                $script:IsMtuInstallRunning = $false
                if ($null -ne $state.Button) {
                    try { $state.Button.Text = $(if ([string]::IsNullOrWhiteSpace([string]$state.OriginalText)) { 'Install Tailscale MTU' } else { [string]$state.OriginalText }); $state.Button.Enabled = $true } catch { }
                }
                Write-LogException -Context 'Monitor Tailscale MTU install' -ErrorRecord $_
            }
        }.GetNewClosure())
        $timer.Start()
        try { Reset-SlowSnapshotCache } catch { }
    }
    catch {
        $script:IsMtuInstallRunning = $false
        if ($null -ne $Button) {
            try { $Button.Text = $(if ([string]::IsNullOrWhiteSpace($originalText)) { 'Install Tailscale MTU' } else { $originalText }); $Button.Enabled = $true } catch { }
        }
        throw
    }
}

function Open-TailscaleMtu {
    $info = Get-TailscaleMtuAppInfo
    if (-not $info.Installed) { throw 'Tailscale MTU is not installed.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$info.ScriptPath) -and (Test-Path -LiteralPath $info.ScriptPath)) {
        $argLine = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $info.ScriptPath + '"'
        Start-Process -FilePath $script:PowerShellExe -ArgumentList $argLine -WindowStyle Hidden
        return
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$info.OpenPath) -and (Test-Path -LiteralPath $info.OpenPath)) {
        Start-Process explorer.exe -ArgumentList @($info.OpenPath)
        return
    }
    throw 'Tailscale MTU installation was not found.'
}

function Open-TailscaleControlPath {
    try {
        $path = ''
        if (-not [string]::IsNullOrWhiteSpace([string]$script:InstalledScriptPath) -and (Test-Path -LiteralPath $script:InstalledScriptPath)) {
            $path = Split-Path -Parent $script:InstalledScriptPath
        }
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { $path = $script:AppRoot }
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { throw 'Tailscale Control path was not found.' }
        Start-Process explorer.exe -ArgumentList @($path)
        Write-ActivityCommandBlock -Title 'Open Tailscale Control Path' -CommandText ('explorer.exe "' + $path + '"') -ExitCode 0 -Output ('Opened path: ' + $path) -DurationMs 0
    }
    catch {
        Write-ActivityFailureBlock -Title 'Open Tailscale Control Path failed' -CommandText 'explorer.exe' -Message $_.Exception.Message
        Show-Overlay -Title 'Open path failed' -Message $_.Exception.Message -Indicator 'Error'
    }
}

function Get-DiagnosticExportDirectory {
    Initialize-AppRoot
    $dir = Join-Path $script:AppRoot 'export'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function ConvertTo-RedactedDiagnosticText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return [string]$Text }
    $value = [string]$Text
    $value = $value -replace '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+.[A-Z]{2,}\b','[redacted-email]'
    $value = $value -replace '(?i)\b(nodekey|machinekey|privatekey|disco|token|secret|authkey)[:=]?[A-Za-z0-9_\-\/+=.]{12,}\b','$1=[redacted]'
    $value = $value -replace '(?i)\b(tskey|tskey-auth|tskey-client)-[A-Za-z0-9_\-]+\b','[redacted-key]'
    $value = $value -replace '(?i)\b[0-9a-f]{32,}\b','[redacted-id]'
    $userName = [System.Environment]::UserName
    if (-not [string]::IsNullOrWhiteSpace([string]$userName)) {
        $escapedUserName = [regex]::Escape($userName)
        $value = $value -replace $escapedUserName,'[redacted-user]'
    }
    $knownDnsPlaceholders = @{}
    $knownDnsIndex = 0
    foreach ($knownDns in @('1.1.1.1','1.0.0.1','8.8.8.8','8.8.4.4','9.9.9.9','149.112.112.112','208.67.222.222','208.67.220.220')) {
        $placeholder = '__TC_KNOWN_DNS_' + $knownDnsIndex + '__'
        $knownDnsPlaceholders[$placeholder] = $knownDns
        $value = $value -replace ('\b' + [regex]::Escape($knownDns) + '\b'), $placeholder
        $knownDnsIndex++
    }
    $value = $value -replace '(?i)\b([A-Za-z0-9][A-Za-z0-9-]*)\.([A-Za-z0-9][A-Za-z0-9-]*)\.ts\.net\b','$1.[redacted-tailnet].ts.net'
    $value = $value -replace '(?i)\b[A-Za-z0-9][A-Za-z0-9-]*\.ts\.net\b','[redacted-tailnet].ts.net'
    $value = $value -replace '(?i)\bfd7a:[0-9a-f:]+\b','[redacted-tailscale-ipv6]'
    $value = $value -replace '\b100\.(?:\d{1,3}\.){2}\d{1,3}(?::\d+)?\b','[redacted-tailscale-ipv4]'
    $value = $value -replace '\b(?:10\.(?:\d{1,3}\.){2}\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3})(?::\d+)?\b','[redacted-lan-ipv4]'
    $value = $value -replace '\b(?!(?:127|169\.254|0|255)\.)(?:\d{1,3}\.){3}\d{1,3}(?::\d+)?\b','[redacted-public-ipv4]'
    foreach ($entry in $knownDnsPlaceholders.GetEnumerator()) {
        $value = $value -replace [regex]::Escape([string]$entry.Key), [string]$entry.Value
    }
    return $value
}

function ConvertTo-RedactedExportObject {
    param($Value,[int]$Depth = 18)
    if ($null -eq $Value) { return $null }
    try {
        $json = $Value | ConvertTo-Json -Depth $Depth
        $redacted = ConvertTo-RedactedDiagnosticText -Text $json
        return ($redacted | ConvertFrom-Json)
    }
    catch {
        try { return (ConvertTo-RedactedDiagnosticText -Text ([string]$Value)) } catch { return $null }
    }
}

function ConvertTo-DiagnosticExportText {
    param([string]$Text,[bool]$RedactSensitive = $true)
    if ($RedactSensitive) { return (ConvertTo-RedactedDiagnosticText -Text ([string]$Text)) }
    return [string]$Text
}

function ConvertTo-DiagnosticExportObject {
    param($Value,[int]$Depth = 18,[bool]$RedactSensitive = $true)
    if ($null -eq $Value) { return $null }
    if ($RedactSensitive) { return (ConvertTo-RedactedExportObject -Value $Value -Depth $Depth) }
    return $Value
}

function Get-SafeCommandExport {
    param([string]$Exe,[string[]]$Arguments,[bool]$RedactSensitive = $true)
    try {
        if ([string]::IsNullOrWhiteSpace([string]$Exe)) { return [pscustomobject]@{ ExitCode = 1; DurationMs = 0; Output = 'Tailscale executable was not found.' } }
        $result = Invoke-TailscaleCommand -Exe $Exe -Arguments $Arguments
        return [pscustomobject]@{
            ExitCode = [int]$result.ExitCode
            DurationMs = [double]$result.DurationMs
            Output = (ConvertTo-DiagnosticExportText -Text ([string]$result.Output) -RedactSensitive $RedactSensitive)
        }
    }
    catch {
        return [pscustomobject]@{ ExitCode = 1; DurationMs = 0; Output = (ConvertTo-DiagnosticExportText -Text ([string]$_.Exception.Message) -RedactSensitive $RedactSensitive) }
    }
}

function Get-SafeCommandRawExport {
    param([string]$Exe,[string[]]$Arguments,[bool]$RedactSensitive = $true)
    try {
        if ([string]::IsNullOrWhiteSpace([string]$Exe)) { return [pscustomobject]@{ ExitCode = 1; DurationMs = 0; RawOutput = ''; Output = 'Tailscale executable was not found.' } }
        $result = Invoke-TailscaleCommand -Exe $Exe -Arguments $Arguments
        return [pscustomobject]@{
            ExitCode = [int]$result.ExitCode
            DurationMs = [double]$result.DurationMs
            RawOutput = [string]$result.Output
            Output = (ConvertTo-DiagnosticExportText -Text ([string]$result.Output) -RedactSensitive $RedactSensitive)
        }
    }
    catch {
        $message = [string]$_.Exception.Message
        return [pscustomobject]@{ ExitCode = 1; DurationMs = 0; RawOutput = $message; Output = (ConvertTo-DiagnosticExportText -Text $message -RedactSensitive $RedactSensitive) }
    }
}

function New-TailscaleSnapshotExport {
    param($Snapshot,[bool]$RedactSensitive = $true)
    if ($null -eq $Snapshot) { return $null }
    $machines = @()
    try { $machines = @(Get-ObjectPropertyOrDefault $Snapshot 'Machines' @()) } catch { $machines = @() }
    $obj = [ordered]@{
        found = [bool](Get-ObjectPropertyOrDefault $Snapshot 'Found' $false)
        backend_state = [string](Get-ObjectPropertyOrDefault $Snapshot 'BackendState' '')
        version = [string](Get-ObjectPropertyOrDefault $Snapshot 'Version' '')
        short_name = [string](Get-ObjectPropertyOrDefault $Snapshot 'ShortName' '')
        dns_name = [string](Get-ObjectPropertyOrDefault $Snapshot 'DNSName' '')
        ipv4 = [string](Get-ObjectPropertyOrDefault $Snapshot 'IPv4' '')
        ipv6 = [string](Get-ObjectPropertyOrDefault $Snapshot 'IPv6' '')
        mtu_ipv4 = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuIPv4' '')
        mtu_ipv6 = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuIPv6' '')
        mtu_interface = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuInterface' '')
        dns_summary = [string](Get-ObjectPropertyOrDefault $Snapshot 'DnsSummary' '')
        dns_nameservers = [string](Get-ObjectPropertyOrDefault $Snapshot 'DnsNameservers' '')
        users_count = [int](Get-ObjectPropertyOrDefault $Snapshot 'UsersCount' 0)
        auto_update = [bool](Get-ObjectPropertyOrDefault $Snapshot 'AutoUpdate' $false)
        mtu_installed = [bool](Get-ObjectPropertyOrDefault $Snapshot 'MtuInstalled' $false)
        mtu_status = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuStatus' '')
        corp_dns = [bool](Get-ObjectPropertyOrDefault $Snapshot 'CorpDNS' $false)
        route_all = [bool](Get-ObjectPropertyOrDefault $Snapshot 'RouteAll' $false)
        incoming_allowed = [bool](Get-ObjectPropertyOrDefault $Snapshot 'IncomingAllowed' $false)
        current_exit_node = [string](Get-ObjectPropertyOrDefault $Snapshot 'CurrentExitNode' '')
        machines_count = [int]$machines.Count
    }
    return (ConvertTo-DiagnosticExportObject -Value ([pscustomobject]$obj) -Depth 12 -RedactSensitive $RedactSensitive)
}

function New-TailscalePeerSummaryExport {
    param($Peer,[int]$Index)
    if ($null -eq $Peer) { return $null }
    return [pscustomobject][ordered]@{
        peer_index = [int]$Index
        host_name = [string](Get-ObjectPropertyOrDefault $Peer 'HostName' '')
        dns_name = [string](Get-ObjectPropertyOrDefault $Peer 'DNSName' '')
        os = [string](Get-ObjectPropertyOrDefault $Peer 'OS' '')
        online = [bool](Get-ObjectPropertyOrDefault $Peer 'Online' $false)
        active = [bool](Get-ObjectPropertyOrDefault $Peer 'Active' $false)
        relay = [string](Get-ObjectPropertyOrDefault $Peer 'Relay' '')
        peer_relay = [string](Get-ObjectPropertyOrDefault $Peer 'PeerRelay' '')
        exit_node = [bool](Get-ObjectPropertyOrDefault $Peer 'ExitNode' $false)
        exit_node_option = [bool](Get-ObjectPropertyOrDefault $Peer 'ExitNodeOption' $false)
        tailscale_ips = @(Get-ObjectPropertyOrDefault $Peer 'TailscaleIPs' @())
        allowed_ips = @(Get-ObjectPropertyOrDefault $Peer 'AllowedIPs' @())
        primary_routes = @(Get-ObjectPropertyOrDefault $Peer 'PrimaryRoutes' @())
        last_seen = [string](Get-ObjectPropertyOrDefault $Peer 'LastSeen' '')
        last_handshake = [string](Get-ObjectPropertyOrDefault $Peer 'LastHandshake' '')
    }
}

function New-TailscaleStatusSummaryExport {
    param($Status,[bool]$RedactSensitive = $true)
    if ($null -eq $Status) { return $null }
    $peers = @()
    try {
        $peerRoot = Get-ObjectPropertyOrDefault $Status 'Peer' $null
        if ($null -ne $peerRoot) {
            $index = 0
            foreach ($prop in $peerRoot.PSObject.Properties) {
                $index++
                $summary = New-TailscalePeerSummaryExport -Peer $prop.Value -Index $index
                if ($null -ne $summary) { $peers += $summary }
            }
        }
    }
    catch { }
    $self = Get-ObjectPropertyOrDefault $Status 'Self' $null
    $obj = [pscustomobject][ordered]@{
        version = [string](Get-ObjectPropertyOrDefault $Status 'Version' '')
        backend_state = [string](Get-ObjectPropertyOrDefault $Status 'BackendState' '')
        tun = [bool](Get-ObjectPropertyOrDefault $Status 'TUN' $false)
        have_node_key = [bool](Get-ObjectPropertyOrDefault $Status 'HaveNodeKey' $false)
        magic_dns_suffix = [string](Get-ObjectPropertyOrDefault $Status 'MagicDNSSuffix' '')
        self = New-TailscalePeerSummaryExport -Peer $self -Index 0
        peers = $peers
        peer_count = [int]$peers.Count
        health = @(Get-ObjectPropertyOrDefault $Status 'Health' @())
        cert_domains = @(Get-ObjectPropertyOrDefault $Status 'CertDomains' @())
    }
    return (ConvertTo-DiagnosticExportObject -Value $obj -Depth 22 -RedactSensitive $RedactSensitive)
}

function New-DiagnosticExportObject {
    param([bool]$RedactSensitive = $true)
    $generatedLocal = Get-Date
    $snapshot = Get-CurrentSnapshot
    $exe = ''
    try { $exe = [string](Get-ObjectPropertyOrDefault $snapshot 'Exe' '') } catch { }
    if ([string]::IsNullOrWhiteSpace($exe)) { $exe = Find-TailscaleExe }
    $statusResult = Get-SafeCommandRawExport -Exe $exe -Arguments @('status','--json') -RedactSensitive $RedactSensitive
    $statusParsed = $null
    $statusSummary = $null
    if ([int]$statusResult.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$statusResult.Output)) {
        try { $statusParsed = [string](Get-ObjectPropertyOrDefault $statusResult 'RawOutput' '') | ConvertFrom-Json } catch { $statusParsed = $null }
        try { $statusSummary = New-TailscaleStatusSummaryExport -Status $statusParsed -RedactSensitive $RedactSensitive } catch { $statusSummary = $null }
    }
    $netcheckResult = Get-SafeCommandExport -Exe $exe -Arguments @('netcheck') -RedactSensitive $RedactSensitive
    $dnsResult = Get-SafeCommandExport -Exe $exe -Arguments @('dns','status','--all') -RedactSensitive $RedactSensitive
    $ipResult = Get-SafeCommandExport -Exe $exe -Arguments @('ip') -RedactSensitive $RedactSensitive
    $mtuExportInfo = $null
    try { $mtuExportInfo = Get-TailscaleMtuAppInfo } catch { $mtuExportInfo = [pscustomobject]@{ Error = [string]$_.Exception.Message } }
    $configExportInfo = $null
    try { $configExportInfo = Get-Config } catch { $configExportInfo = [pscustomobject]@{ Error = [string]$_.Exception.Message } }
    $osInfo = $null
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $osInfo = [ordered]@{
            Caption = [string]$os.Caption
            Version = [string]$os.Version
            BuildNumber = [string]$os.BuildNumber
            OSArchitecture = [string]$os.OSArchitecture
            LastBootUpTime = [string]$os.LastBootUpTime
        }
    }
    catch { $osInfo = [ordered]@{ Error = [string]$_.Exception.Message } }
    $screens = @()
    try {
        foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
            $screens += [pscustomobject]@{
                DeviceName = [string]$screen.DeviceName
                Primary = [bool]$screen.Primary
                Bounds = [pscustomobject]@{ X = [int]$screen.Bounds.X; Y = [int]$screen.Bounds.Y; Width = [int]$screen.Bounds.Width; Height = [int]$screen.Bounds.Height }
                WorkingArea = [pscustomobject]@{ X = [int]$screen.WorkingArea.X; Y = [int]$screen.WorkingArea.Y; Width = [int]$screen.WorkingArea.Width; Height = [int]$screen.WorkingArea.Height }
            }
        }
    }
    catch { }
    $redactionMode = $(if ($RedactSensitive) { 'redacted' } else { 'non-redacted' })
    $redactionScope = [string[]]@()
    if ($RedactSensitive) {
        $redactionScope = [string[]]@('emails','tokens','auth keys','long hexadecimal identifiers','current Windows username','tailnet domains','Tailscale IPs','LAN IPs','public IPv4 addresses except well-known public DNS resolvers')
    }
    return [pscustomobject][ordered]@{
        diagnostic_schema_version = 1
        generated_at_local = $generatedLocal.ToString('yyyy-MM-dd HH:mm:ss')
        generated_at_utc = $generatedLocal.ToUniversalTime().ToString('o')
        app = [pscustomobject][ordered]@{
            name = [string]$script:AppName
            version = [string]$script:AppVersion
            release_tag = [string]$script:ReleaseTag
            app_user_model_id = [string]$script:AppUserModelId
            script_path = (ConvertTo-DiagnosticExportText -Text ([string]$script:ScriptPath) -RedactSensitive $RedactSensitive)
            app_root = (ConvertTo-DiagnosticExportText -Text ([string]$script:AppRoot) -RedactSensitive $RedactSensitive)
            installed_script_path = (ConvertTo-DiagnosticExportText -Text ([string]$script:InstalledScriptPath) -RedactSensitive $RedactSensitive)
            icon_path = (ConvertTo-DiagnosticExportText -Text ([string](Get-AppIconPath)) -RedactSensitive $RedactSensitive)
            icon_file_found = [bool](-not [string]::IsNullOrWhiteSpace([string](Get-AppIconPath)))
            running_background_mode = [bool]$Background
            is_admin = [bool](Test-IsAdministrator)
        }
        system = [pscustomobject][ordered]@{
            computer_name = (ConvertTo-DiagnosticExportText -Text ([string][System.Environment]::MachineName) -RedactSensitive $RedactSensitive)
            user_name = (ConvertTo-DiagnosticExportText -Text ([string][System.Environment]::UserName) -RedactSensitive $RedactSensitive)
            powershell_version = [string]$PSVersionTable.PSVersion
            dotnet_version = [string][System.Environment]::Version
            culture = [string][Globalization.CultureInfo]::CurrentCulture.Name
            ui_culture = [string][Globalization.CultureInfo]::CurrentUICulture.Name
            os = $osInfo
            screens = $screens
        }
        tailscale = [pscustomobject][ordered]@{
            executable = (ConvertTo-DiagnosticExportText -Text ([string]$exe) -RedactSensitive $RedactSensitive)
            snapshot_summary = (New-TailscaleSnapshotExport -Snapshot $snapshot -RedactSensitive $RedactSensitive)
            status_summary = $statusSummary
            commands = [pscustomobject][ordered]@{
                status_json_collected = [bool]([int]$statusResult.ExitCode -eq 0)
                status_json_duration_ms = [double](Get-ObjectPropertyOrDefault $statusResult 'DurationMs' 0)
                netcheck = $netcheckResult
                dns_status_all = $dnsResult
                ip = $ipResult
            }
        }
        tailscale_mtu = (ConvertTo-DiagnosticExportObject -Value $mtuExportInfo -Depth 16 -RedactSensitive $RedactSensitive)
        config = (ConvertTo-DiagnosticExportObject -Value $configExportInfo -Depth 18 -RedactSensitive $RedactSensitive)
        paths = [pscustomobject][ordered]@{
            config = (ConvertTo-DiagnosticExportText -Text ([string]$script:ConfigPath) -RedactSensitive $RedactSensitive)
            log = (ConvertTo-DiagnosticExportText -Text ([string]$script:LogPath) -RedactSensitive $RedactSensitive)
            startup_vbs = (ConvertTo-DiagnosticExportText -Text ([string]$script:StartupVbsPath) -RedactSensitive $RedactSensitive)
            launcher_vbs = (ConvertTo-DiagnosticExportText -Text ([string]$script:LauncherVbsPath) -RedactSensitive $RedactSensitive)
            start_menu_shortcut = (ConvertTo-DiagnosticExportText -Text ([string]$script:StartMenuShortcutPath) -RedactSensitive $RedactSensitive)
        }
        activity = [pscustomobject][ordered]@{
            log_tail = (ConvertTo-DiagnosticExportText -Text ([string](Get-LogTail)) -RedactSensitive $RedactSensitive)
        }
        redaction = [pscustomobject][ordered]@{
            enabled = [bool]$RedactSensitive
            mode = [string]$redactionMode
            scope = $redactionScope
        }
    }
}

function Export-TailscaleDiagnosticJson {
    param([bool]$RedactSensitive = $true)
    if (-not $RedactSensitive) {
        $answer = [System.Windows.Forms.MessageBox]::Show('This export may include private IPs, device names, tailnet names, usernames, DNS data, routes, and command history. Continue?', 'Export full diagnostic JSON', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $null }
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $dir = Get-DiagnosticExportDirectory
    $redactionLabel = $(if ($RedactSensitive) { 'redacted' } else { 'non-redacted' })
    $fileName = 'tailscale-control-diagnostics-' + $redactionLabel + '-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.json'
    $path = Join-Path $dir $fileName
    $payload = New-DiagnosticExportObject -RedactSensitive $RedactSensitive
    $json = $payload | ConvertTo-Json -Depth 28
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8 -Force
    $sw.Stop()
    Start-Process explorer.exe -ArgumentList @($dir)
    Write-ActivityCommandBlock -Title 'Export Diagnostic JSON' -CommandText ('Export to "' + $path + '"') -ExitCode 0 -Output ('Exported file: ' + $path) -DurationMs ([double]$sw.Elapsed.TotalMilliseconds)
    Show-Overlay -Title 'Diagnostic JSON exported' -Message $fileName -Indicator 'Info'
    return $path
}

function Convert-LastSeenText {
    param($Peer,[string]$Status)
    if ([string]$Status -eq 'Online') { return 'Connected' }
    if ([string]$Status -eq 'This device') { return 'Disconnected' }
    $value = Get-PropertyValue $Peer @('LastSeen','LastWrite')
    if ($null -eq $value) { return '' }
    try {
        $dt = [datetime]$value
        if ($dt.Year -le 2001) {
            if ([string]$Status -eq 'Offline') { return 'Offline' }
            return ''
        }
        return $dt.ToLocalTime().ToString('MMM d, HH:mm')
    }
    catch { return [string]$value }
}

function Convert-PeerToMachine {
    param($Peer,[switch]$ThisDevice,$UserMap,$SelfVersion)
    $dns = ConvertTo-DnsName ([string](Get-PropertyValue $Peer @('DNSName')))
    $name = ConvertTo-DnsName ([string](Get-PropertyValue $Peer @('HostName','ComputedName','Name')))
    if ([string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($dns)) {
        $name = ($dns -split '\.')[0]
    }
    $userId = Get-PropertyValue $Peer @('UserID')
    $owner = Resolve-UserLabel -UserMap $UserMap -UserId $userId
    if ([string]::IsNullOrWhiteSpace($owner)) {
        $owner = [string](Get-PropertyValue (Get-PropertyValue $Peer @('UserProfile')) @('DisplayName','LoginName'))
    }
    if ([string]::IsNullOrWhiteSpace($owner)) { $owner = [string](Get-PropertyValue $Peer @('User','UserName')) }
    if ([string]::IsNullOrWhiteSpace($owner) -and $ThisDevice) { $owner = [System.Environment]::UserName }
    $onlineBool = Convert-ToNullableBool (Get-PropertyValue $Peer @('Online'))
    $activeBool = Convert-ToNullableBool (Get-PropertyValue $Peer @('Active'))
    $exitNodeBool = Convert-ToNullableBool (Get-PropertyValue $Peer @('ExitNode'))
    $exitNodeOptionBool = Convert-ToNullableBool (Get-PropertyValue $Peer @('ExitNodeOption'))
    $status = if ($ThisDevice) {
        if ($onlineBool -eq $true) { 'Online' } else { 'This device' }
    } else {
        if ($onlineBool -eq $true) { 'Online' } elseif ($onlineBool -eq $false) { 'Offline' } else { '' }
    }
    $ips = Convert-ToObjectArray (Get-PropertyValue $Peer @('TailscaleIPs'))
    $ipv4 = ($ips | Where-Object { [string]$_ -match '^[0-9]+\.' } | Select-Object -First 1)
    $ipv6 = ($ips | Where-Object { [string]$_ -match ':' } | Select-Object -First 1)
    $hostinfo = Get-PropertyValue $Peer @('Hostinfo')
    $os = [string](Get-PropertyValue $hostinfo @('OS','os'))
    if ([string]::IsNullOrWhiteSpace($os)) { $os = [string](Get-PropertyValue $Peer @('OS')) }
    $tsVersion = [string](Get-PropertyValue $Peer @('TailscaleVersion'))
    if ([string]::IsNullOrWhiteSpace($tsVersion) -and $ThisDevice) { $tsVersion = [string]$SelfVersion }
    if ([string]::IsNullOrWhiteSpace($tsVersion)) { $tsVersion = '-' }
    $tags = New-Object System.Collections.Generic.List[string]
    if ($ThisDevice) { [void]$tags.Add('This device') }
    if ($exitNodeBool -eq $true) { [void]$tags.Add('Using exit node') }
    if ($exitNodeOptionBool -eq $true) { [void]$tags.Add('Exit node option') }
    $primaryRoutes = Convert-ToObjectArray (Get-PropertyValue $Peer @('PrimaryRoutes'))
    $allowedIPs = Convert-ToObjectArray (Get-PropertyValue $Peer @('AllowedIPs'))
    $primaryRoutesCount = @(Convert-ToObjectArray $primaryRoutes).Count
    $allowedIPsCount = @(Convert-ToObjectArray $allowedIPs).Count
    if ($primaryRoutesCount -gt 0 -or $allowedIPsCount -gt 2) { [void]$tags.Add('Subnets') }
    $lastSeen = Convert-LastSeenText -Peer $Peer -Status $status
    $addressText = ((@($ipv4,$ipv6) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ' | ')
    $deviceId = [string](Get-PropertyValue $Peer @('ID','StableID','NodeID'))
    $addrs = Convert-ToObjectArray (Get-PropertyValue $Peer @('Addrs'))
    $relay = [string](Get-PropertyValue $Peer @('Relay'))
    $curAddr = [string](Get-PropertyValue $Peer @('CurAddr'))
    $created = [string](Get-PropertyValue $Peer @('Created'))
    $lastWrite = [string](Get-PropertyValue $Peer @('LastWrite'))
    $lastHandshake = [string](Get-PropertyValue $Peer @('LastHandshake'))
    $peerApi = Convert-ToObjectArray (Get-PropertyValue $Peer @('PeerAPIURL'))
    $taildrop = [string](Get-PropertyValue $Peer @('TaildropTarget'))
    $publicKey = [string](Get-PropertyValue $Peer @('PublicKey'))
    $capabilities = Convert-ToObjectArray (Get-PropertyValue $Peer @('Capabilities'))
    $capabilityLabels = New-Object System.Collections.Generic.List[string]
    $capString = ((@($capabilities) | ForEach-Object { [string]$_ }) -join ' ')
    if ($capString -match '(?i)default-auto-update') { [void]$capabilityLabels.Add('Auto update') }
    if ($capString -match '(?i)(^|[#/])ssh($|\s)') { [void]$capabilityLabels.Add('SSH') }
    if ($capString -match '(?i)(^|[#/])funnel($|[-?/\s])') { [void]$capabilityLabels.Add('Funnel') }
    if ($capString -match '(?i)(^|[#/])https($|\s)') { [void]$capabilityLabels.Add('HTTPS') }
    if ($capString -match '(?i)file-sharing') { [void]$capabilityLabels.Add('File sharing') }
    if ($capString -match '(?i)is-admin') { [void]$capabilityLabels.Add('Admin') }
    if ($capString -match '(?i)is-owner') { [void]$capabilityLabels.Add('Owner') }
    if ($capString -match '(?i)tailnet-lock') { [void]$capabilityLabels.Add('Tailnet lock') }
    if ($capString -match '(?i)app-connectors') { [void]$capabilityLabels.Add('App connectors') }
    $noFileSharingReason = [string](Get-PropertyValue $Peer @('NoFileSharingReason'))
    $rxBytes = [double](Get-PropertyValue $Peer @('RxBytes'))
    $txBytes = [double](Get-PropertyValue $Peer @('TxBytes'))
    $inNetworkMap = Convert-ToNullableBool (Get-PropertyValue $Peer @('InNetworkMap'))
    $inMagicSock = Convert-ToNullableBool (Get-PropertyValue $Peer @('InMagicSock'))
    $inEngine = Convert-ToNullableBool (Get-PropertyValue $Peer @('InEngine'))
    $connection = ''
    if ($ThisDevice) {
        $connection = 'Local'
    }
    elseif ($onlineBool -eq $true) {
        if (-not [string]::IsNullOrWhiteSpace($curAddr)) { $connection = 'Direct' }
        elseif (-not [string]::IsNullOrWhiteSpace($relay)) { $connection = 'Relay' }
        else { $connection = 'Online' }
    }
    elseif ($ThisDevice -and [string]$status -eq 'This device') {
        $connection = 'Stopped'
    }
    elseif ($onlineBool -eq $false) {
        $connection = 'Offline'
    }
    $rawJson = ''
    try { $rawJson = ($Peer | ConvertTo-Json -Depth 8) } catch { Write-LogException -Context 'Serialize peer to JSON' -ErrorRecord $_; $rawJson = '' }
    return [pscustomobject]@{
        Machine = (ConvertTo-DiagnosticText -Text $name)
        Owner = (ConvertTo-DiagnosticText -Text $owner)
        IsLocal = [bool]$ThisDevice
        Addresses = $addressText
        Version = $tsVersion
        LastSeen = $lastSeen
        Tags = ($tags -join ', ')
        ShortName = (ConvertTo-DiagnosticText -Text $name)
        DNSName = $dns
        IPv4 = [string]$ipv4
        IPv6 = [string]$ipv6
        Status = $status
        OS = $os
        DeviceId = $deviceId
        UserID = [string]$userId
        PublicKey = $publicKey
        OnlineText = $(if ($onlineBool -eq $true) { 'True' } elseif ($onlineBool -eq $false) { 'False' } else { '' })
        ActiveText = $(if ($activeBool -eq $true) { 'True' } elseif ($activeBool -eq $false) { 'False' } else { '' })
        ExitNodeText = $(if ($exitNodeBool -eq $true) { 'True' } elseif ($exitNodeBool -eq $false) { 'False' } else { '' })
        ExitNodeOptionText = $(if ($exitNodeOptionBool -eq $true) { 'True' } elseif ($exitNodeOptionBool -eq $false) { 'False' } else { '' })
        Relay = $relay
        CurAddr = $curAddr
        Connection = $connection
        Addrs = ($addrs -join ', ')
        AllowedIPs = ($allowedIPs -join ', ')
        AllowedIPsCount = [string]$allowedIPsCount
        PrimaryRoutes = ($primaryRoutes -join ', ')
        PrimaryRoutesCount = [string]$primaryRoutesCount
        PeerAPIURL = ($peerApi -join ', ')
        TaildropTarget = $taildrop
        Created = $created
        LastWrite = $lastWrite
        LastHandshake = $lastHandshake
        NoFileSharingReason = $noFileSharingReason
        InNetworkMapText = $(if ($null -eq $inNetworkMap) { '' } else { [string][bool]$inNetworkMap })
        InMagicSockText = $(if ($null -eq $inMagicSock) { '' } else { [string][bool]$inMagicSock })
        InEngineText = $(if ($null -eq $inEngine) { '' } else { [string][bool]$inEngine })
        RawJson = $rawJson
        CapabilitiesCount = [string](@($capabilities).Count)
        CapabilitiesText = $(if ($capabilityLabels.Count -gt 0) { [string]::Join(', ', $capabilityLabels.ToArray()) } else { '-' })
        RxBytes = [string]$rxBytes
        TxBytes = [string]$txBytes
        RxText = (Convert-ToHumanBytes $rxBytes)
        TxText = (Convert-ToHumanBytes $txBytes)
    }
}

function Get-ExitNodes {
    param([string]$Exe)
    $list = @()
    $result = Invoke-TailscaleCommand -Exe $Exe -Arguments @('exit-node','list')
    if ($result.ExitCode -ne 0) { return $list }
    $lines = @($result.Output -split "`r?`n")
    foreach ($line in $lines) {
        $trim = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        if ($trim -match '^(IP|-|WARNING|No exit nodes)') { continue }
        if ($trim -match '^([0-9a-fA-F\.:]+)\s+([^\s]+)\s+([^\s]+)') {
            $ip = $Matches[1]
            $name = ConvertTo-DnsName $Matches[2]
            $status = $Matches[3]
            $list += [pscustomobject]@{ Node = $name; Name = $name; DNSName = $name; IPv4 = $ip; Status = $status }
        }
    }
    return $list
}

function Get-PrefsObject {
    param([string]$Exe)
    $result = Invoke-TailscaleCommand -Exe $Exe -Arguments @('debug','prefs')
    if ($result.ExitCode -ne 0 -or [string]::IsNullOrWhiteSpace($result.Output)) { return $null }
    return (ConvertFrom-JsonSafe -Text $result.Output -Context 'Parse tailscale debug prefs')
}

function Get-AutoUpdatePrefValue {
    param($Prefs)
    if ($null -eq $Prefs) { return $null }

    $direct = Convert-ToNullableBool (Get-PropertyValue $Prefs @('AutoUpdateEnabled','AutoUpdatesEnabled','AutomaticUpdatesEnabled'))
    if ($null -ne $direct) { return $direct }

    $auto = Get-PropertyValue $Prefs @('AutoUpdate','AutoUpdates','AutomaticUpdates')
    if ($null -eq $auto) { return $null }

    if ($auto -is [bool] -or $auto -is [string] -or $auto -is [int]) {
        return (Convert-ToNullableBool $auto)
    }

    $apply = Convert-ToNullableBool (Get-PropertyValue $auto @('Apply','apply','Enabled','enabled','Value','value'))
    if ($null -ne $apply) { return $apply }

    $check = Convert-ToNullableBool (Get-PropertyValue $auto @('Check','check'))
    if ($null -ne $check) { return $check }

    return $null
}

function Get-TailscaleSnapshot {
    $exe = Find-TailscaleExe
    if ([string]::IsNullOrWhiteSpace([string]$exe)) {
        return [pscustomobject]@{
            Found = $false
            Exe = ''
            BackendState = 'Not detected'
            Version = ''
            FullVersion = ''
            User = ''
            UserEmail = ''
            Tailnet = ''
            TailnetId = ''
            ShortName = ''
            DNSName = ''
            IPv4 = ''
            IPv6 = ''
            MtuIPv4 = ''
            MtuIPv6 = ''
            MtuInterface = ''
            DnsSummary = ''
            DnsNameservers = 'System default'
            DnsRaw = ''
            UsersJoined = ''
            UsersCount = 0
            AutoUpdate = $null
            MtuInstalled = $false
            MtuStatus = 'Not detected'
            MtuOpenPath = ''
            MtuScriptPath = ''
            MtuServiceName = ''
            MtuServiceState = 'Not detected'
            MtuConfigPath = ''
            MtuStatePath = ''
            MtuDesired = ''
            MtuDesiredIPv4 = ''
            MtuDesiredIPv6 = ''
            MtuCheckInterval = ''
            MtuInterfaceMatch = ''
            MtuDetectedInterface = ''
            MtuLastError = ''
            MtuLastResult = ''
            MtuVersion = ''
            CorpDNS = $null
            RouteAll = $null
            IncomingAllowed = $null
            CurrentExitNode = ''
            ExitNodes = @()
            Machines = @()
            LogTail = [string](Get-ActivityLogTail)
        }
    }
    $statusRaw = Invoke-TailscaleCommand -Exe $exe -Arguments @('status','--json')
    $status = $null
    if ($statusRaw.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($statusRaw.Output)) {
        $status = ConvertFrom-JsonSafe -Text $statusRaw.Output -Context 'Parse tailscale status --json'
    }
    $slowData = Get-SlowSnapshotData -Exe $exe
    $version = [string]$slowData.Version
    $prefs = $slowData.Prefs
    $self = if ($null -ne $status) { Get-PropertyValue $status @('Self') } else { $null }
    $userMap = if ($null -ne $status) { Get-PropertyValue $status @('User') } else { $null }
    $magic = ConvertTo-DnsName ([string](Get-PropertyValue $status @('MagicDNSSuffix')))
    $dnsName = ConvertTo-DnsName ([string](Get-PropertyValue $self @('DNSName')))
    $shortName = ConvertTo-DnsName ([string](Get-PropertyValue $self @('HostName','ComputedName','Name')))
    if ([string]::IsNullOrWhiteSpace($shortName) -and -not [string]::IsNullOrWhiteSpace($dnsName)) { $shortName = ($dnsName -split '\.')[0] }
    $ips = @()
    $selfIps = Get-PropertyValue $self @('TailscaleIPs')
    if ($selfIps -is [System.Array]) { $ips = @($selfIps) } elseif ($null -ne $selfIps) { $ips = @($selfIps) }
    $ipv4 = ($ips | Where-Object { $_ -match '^[0-9]+\.' } | Select-Object -First 1)
    $ipv6 = ($ips | Where-Object { $_ -match ':' } | Select-Object -First 1)
    $userProfile = Get-PropertyValue $self @('UserProfile')
    $selfUserId = Get-PropertyValue $self @('UserID')
    $resolvedProfile = $null
    if ($null -ne $selfUserId -and $null -ne $userMap) { $resolvedProfile = Get-PropertyValue $userMap @([string]$selfUserId) }
    $userDisplay = [string](Get-PropertyValue $userProfile @('DisplayName'))
    if ([string]::IsNullOrWhiteSpace($userDisplay)) { $userDisplay = [string](Get-PropertyValue $resolvedProfile @('DisplayName')) }
    $userEmail = [string](Get-PropertyValue $userProfile @('LoginName'))
    if ([string]::IsNullOrWhiteSpace($userEmail)) { $userEmail = [string](Get-PropertyValue $resolvedProfile @('LoginName')) }
    if ([string]::IsNullOrWhiteSpace($userEmail)) { $userEmail = [string](Get-PropertyValue (Get-PropertyValue $status @('CurrentTailnet')) @('Name')) }
    if ([string]::IsNullOrWhiteSpace($userDisplay)) { $userDisplay = if (-not [string]::IsNullOrWhiteSpace($userEmail)) { $userEmail } else { [System.Environment]::UserName } }
    $backend = [string](Get-PropertyValue $status @('BackendState'))
    $tailnet = if (-not [string]::IsNullOrWhiteSpace($magic)) { $magic } elseif (-not [string]::IsNullOrWhiteSpace($dnsName) -and $dnsName -match '^[^.]+\.(.+)$') { $Matches[1] } else { '' }
    $exitNodes = Convert-ToObjectArray $slowData.ExitNodes
    $currentExit = ''
    $exitIP = ConvertTo-DnsName ([string](Get-PropertyValue $prefs @('ExitNodeIP','ExitNodeIPString')))
    $exitID = ConvertTo-DnsName ([string](Get-PropertyValue $prefs @('ExitNodeID','ExitNode')))
    foreach ($node in @($exitNodes)) {
        $label = ConvertTo-DnsName ([string]$node.DNSName)
        $nodeIP = ConvertTo-DnsName ([string]$node.IPv4)
        $nodeName = ConvertTo-DnsName ([string]$node.Name)
        if ((-not [string]::IsNullOrWhiteSpace($exitIP) -and $nodeIP -eq $exitIP) -or (-not [string]::IsNullOrWhiteSpace($exitID) -and ($label -eq $exitID -or $nodeName -eq $exitID))) {
            $currentExit = $label
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($currentExit) -and -not [string]::IsNullOrWhiteSpace($exitID)) { $currentExit = $exitID }
    $machines = New-Object System.Collections.Generic.List[object]
    if ($null -ne $self) {
        $selfMachine = Convert-PeerToMachine -Peer $self -ThisDevice -UserMap $userMap -SelfVersion $version
        if ([string]::IsNullOrWhiteSpace([string]$selfMachine.Owner)) { $selfMachine.Owner = [string]$userDisplay }
        if ([string]::IsNullOrWhiteSpace([string]$selfMachine.Version)) { $selfMachine.Version = [string]$version }
        [void]$machines.Add($selfMachine)
    }
    $peers = Get-PropertyValue $status @('Peer')
    if ($null -ne $peers) {
        foreach ($prop in $peers.PSObject.Properties) {
            [void]$machines.Add((Convert-PeerToMachine -Peer $prop.Value -UserMap $userMap -SelfVersion $version))
        }
    }
    $corpDns = Convert-ToNullableBool (Get-PropertyValue $prefs @('CorpDNS'))
    $routeAll = Convert-ToNullableBool (Get-PropertyValue $prefs @('RouteAll'))
    $shieldsUp = Convert-ToNullableBool (Get-PropertyValue $prefs @('ShieldsUp'))
    $autoUpdate = Get-AutoUpdatePrefValue -Prefs $prefs
    if ($null -eq $autoUpdate -and $null -ne $script:AutoUpdateOverride) { $autoUpdate = Convert-ToNullableBool $script:AutoUpdateOverride }
    $dnsInfo = $slowData.DnsInfo
    if ($corpDns -eq $false -or [string]$backend -ne 'Running') { $dnsInfo.Nameservers = Get-SystemDnsNameservers }
    $incomingAllowed = $null
    if ($null -ne $shieldsUp) { $incomingAllowed = -not ([bool]$shieldsUp) }
    $mtuInfo = $slowData.MtuInfo
    $mtuApp = $slowData.MtuApp
    $tailnetId = [string](Get-PropertyValue (Get-PropertyValue $status @('CurrentTailnet')) @('ID','TailnetID'))
    $usersJoined = New-Object System.Collections.Generic.List[string]
    if ($null -ne $userMap) {
        foreach ($prop in $userMap.PSObject.Properties) {
            $login = [string](Get-PropertyValue $prop.Value @('LoginName'))
            $display = [string](Get-PropertyValue $prop.Value @('DisplayName'))
            $label = if (-not [string]::IsNullOrWhiteSpace($display) -and -not [string]::IsNullOrWhiteSpace($login)) { $display + ' <' + $login + '>' } elseif (-not [string]::IsNullOrWhiteSpace($login)) { $login } else { $display }
            $label = ConvertTo-DiagnosticText -Text $label
            if (-not [string]::IsNullOrWhiteSpace($label)) { [void]$usersJoined.Add($label) }
        }
    }

    $exitNodeArray = Convert-ToObjectArray $exitNodes
    $machineArray = Convert-ToObjectArray $machines

    try {
        $snapshotData = [ordered]@{}
        $snapshotData['Found'] = $true
        $snapshotData['Exe'] = [string]$exe
        $snapshotData['BackendState'] = [string]$backend
        $snapshotData['Version'] = [string]$version
        $snapshotData['FullVersion'] = [string](Get-PropertyValue $status @('Version'))
        $snapshotData['User'] = ConvertTo-DiagnosticText -Text ([string]$userDisplay)
        $snapshotData['UserEmail'] = ConvertTo-DiagnosticText -Text ([string]$userEmail)
        $snapshotData['Tailnet'] = [string]$tailnet
        $snapshotData['TailnetId'] = [string]$tailnetId
        $snapshotData['ShortName'] = ConvertTo-DiagnosticText -Text ([string]$shortName)
        $snapshotData['DNSName'] = [string]$dnsName
        $snapshotData['IPv4'] = [string]$ipv4
        $snapshotData['IPv6'] = [string]$ipv6
        $snapshotData['MtuIPv4'] = [string]$mtuInfo.IPv4
        $snapshotData['MtuIPv6'] = [string]$mtuInfo.IPv6
        $snapshotData['MtuInterface'] = [string]$mtuInfo.Interface
        $snapshotData['DnsSummary'] = [string]$dnsInfo.Summary
        $snapshotData['DnsNameservers'] = [string]$dnsInfo.Nameservers
        $snapshotData['DnsRaw'] = [string]$dnsInfo.Raw
        $snapshotData['UsersJoined'] = ($usersJoined -join [Environment]::NewLine)
        $snapshotData['UsersCount'] = [int]$usersJoined.Count
        $snapshotData['AutoUpdate'] = $autoUpdate
        $snapshotData['MtuInstalled'] = [bool]$mtuApp.Installed
        $snapshotData['MtuStatus'] = [string]$mtuApp.StatusText
        $snapshotData['MtuOpenPath'] = [string]$mtuApp.OpenPath
        $snapshotData['MtuScriptPath'] = [string]$mtuApp.ScriptPath
        $snapshotData['MtuServiceName'] = [string]$mtuApp.ServiceName
        $snapshotData['MtuServiceState'] = [string]$mtuApp.ServiceState
        $snapshotData['MtuConfigPath'] = [string]$mtuApp.ConfigPath
        $snapshotData['MtuStatePath'] = [string]$mtuApp.StatePath
        $snapshotData['MtuDesired'] = [string]$mtuApp.DesiredMtu
        $snapshotData['MtuDesiredIPv4'] = [string]$mtuApp.DesiredMtuIPv4
        $snapshotData['MtuDesiredIPv6'] = [string]$mtuApp.DesiredMtuIPv6
        $snapshotData['MtuCheckInterval'] = [string]$mtuApp.CheckInterval
        $snapshotData['MtuInterfaceMatch'] = [string]$mtuApp.InterfaceMatch
        $snapshotData['MtuDetectedInterface'] = [string]$mtuApp.DetectedInterface
        $snapshotData['MtuLastError'] = [string]$mtuApp.LastError
        $snapshotData['MtuLastResult'] = [string]$mtuApp.LastResult
        $snapshotData['MtuVersion'] = [string]$mtuApp.Version
        $snapshotData['CorpDNS'] = $corpDns
        $snapshotData['RouteAll'] = $routeAll
        $snapshotData['IncomingAllowed'] = $incomingAllowed
        $snapshotData['CurrentExitNode'] = [string]$currentExit
        $snapshotData['ExitNodes'] = $exitNodeArray
        $snapshotData['Machines'] = $machineArray
        $snapshotData['LogTail'] = [string](Get-ActivityLogTail)
        return [pscustomobject]$snapshotData
    }
    catch {
        Write-Log ('Snapshot object build failed: ' + $_.Exception.Message)
        Write-Log ('Snapshot object build position: ' + $_.InvocationInfo.PositionMessage)
        return [pscustomobject]@{
            Found = $true
            Exe = [string]$exe
            BackendState = [string]$backend
            Version = [string]$version
            User = [string]$userDisplay
            UserEmail = [string]$userEmail
            Tailnet = [string]$tailnet
            TailnetId = [string]$tailnetId
            ShortName = [string]$shortName
            DNSName = [string]$dnsName
            IPv4 = [string]$ipv4
            IPv6 = [string]$ipv6
            MtuIPv4 = [string]$mtuInfo.IPv4
            MtuIPv6 = [string]$mtuInfo.IPv6
            MtuInterface = [string]$mtuInfo.Interface
            DnsSummary = [string]$dnsInfo.Summary
            DnsNameservers = [string]$dnsInfo.Nameservers
            DnsRaw = [string]$dnsInfo.Raw
            UsersJoined = ($usersJoined -join [Environment]::NewLine)
            AutoUpdate = $autoUpdate
            MtuInstalled = [bool]$mtuApp.Installed
            MtuStatus = [string]$mtuApp.StatusText
            MtuOpenPath = [string]$mtuApp.OpenPath
            MtuScriptPath = [string]$mtuApp.ScriptPath
            MtuServiceName = [string]$mtuApp.ServiceName
            MtuServiceState = [string]$mtuApp.ServiceState
            MtuConfigPath = [string]$mtuApp.ConfigPath
            MtuStatePath = [string]$mtuApp.StatePath
            MtuDesired = [string]$mtuApp.DesiredMtu
            MtuDesiredIPv4 = [string]$mtuApp.DesiredMtuIPv4
            MtuDesiredIPv6 = [string]$mtuApp.DesiredMtuIPv6
            MtuCheckInterval = [string]$mtuApp.CheckInterval
            MtuInterfaceMatch = [string]$mtuApp.InterfaceMatch
            MtuDetectedInterface = [string]$mtuApp.DetectedInterface
            MtuLastError = [string]$mtuApp.LastError
            MtuLastResult = [string]$mtuApp.LastResult
            MtuVersion = [string]$mtuApp.Version
            CorpDNS = $null
            RouteAll = $null
            IncomingAllowed = $null
            CurrentExitNode = [string]$currentExit
            ExitNodes = @()
            Machines = @()
            LogTail = [string](Get-ActivityLogTail)
        }
    }
}

function Get-PreferredExitNodeLabel {
    $label = ConvertTo-DnsName ([string]$config.preferred_exit_label)
    if ([string]::IsNullOrWhiteSpace($label)) { $label = ConvertTo-DnsName ([string]$config.preferred_exit_node) }
    return $label
}

function Get-ExitNodeToggleAvailability {
    param($Snapshot)
    $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
    if ($null -eq $snap) {
        return [pscustomobject]@{ Enabled = $false; HasExitNodes = $false; CurrentExitNodeActive = $false; ExitNodeCount = 0; State = 'Unknown'; Reason = 'Tailscale status has not been loaded yet.' }
    }
    $found = [bool](Get-ObjectPropertyOrDefault $snap 'Found' $false)
    $currentExitNode = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'CurrentExitNode' ''))
    $currentActive = -not [string]::IsNullOrWhiteSpace($currentExitNode)
    $exitNodes = @(Convert-ToObjectArray (Get-ObjectPropertyOrDefault $snap 'ExitNodes' @()))
    $hasExitNodes = ($exitNodes.Count -gt 0)
    $enabled = [bool]($found -and ($currentActive -or $hasExitNodes))
    $state = if (-not $found) { 'Not detected' } elseif ($currentActive) { 'On' } elseif ($hasExitNodes) { 'Off' } else { 'No nodes' }
    $reason = if (-not $found) { 'tailscale.exe was not detected.' } elseif ($currentActive) { 'An exit node is active and can be cleared.' } elseif ($hasExitNodes) { 'At least one exit node is available.' } else { 'This tailnet does not currently expose any exit nodes.' }
    return [pscustomobject]@{ Enabled = $enabled; HasExitNodes = $hasExitNodes; CurrentExitNodeActive = $currentActive; ExitNodeCount = [int]$exitNodes.Count; State = $state; Reason = $reason }
}

function Update-ExitNodeActionAvailability {
    param($Snapshot)
    try {
        $availability = Get-ExitNodeToggleAvailability -Snapshot $Snapshot
        if ($null -ne $script:btnToggleExit) {
            $script:btnToggleExit.Enabled = [bool]$availability.Enabled
            Set-AppToolTip -Control $script:btnToggleExit -Text $(if ([bool]$availability.Enabled) { 'Enables the preferred exit node if none is active, or clears the current exit node if one is active.' } else { [string]$availability.Reason })
        }
        if ($null -ne $script:cmbExitNode) {
            $script:cmbExitNode.Enabled = [bool]$availability.HasExitNodes
            Set-AppToolTip -Control $script:cmbExitNode -Text $(if ([bool]$availability.HasExitNodes) { "Selects which exit node the Toggle Exit Node action should prefer. Empty means the app will not force a preferred node.`r`nDefault: Empty" } else { [string]$availability.Reason })
        }
    }
    catch { Write-LogException -Context 'Update exit node action availability' -ErrorRecord $_ }
}

function Set-StartupSetting {
    param($Config)
    if ([bool]$Config.start_with_windows) {
        $launch = if (Test-Path -LiteralPath $script:InstalledScriptPath) { $script:InstalledScriptPath } else { $script:ScriptPath }
        Write-AppLauncherVbs -ScriptPath $launch -BackgroundMode:([bool]$Config.start_minimized)
        Copy-Item -LiteralPath $script:LauncherVbsPath -Destination $script:StartupVbsPath -Force
    }
    else {
        if (Test-Path -LiteralPath $script:StartupVbsPath) { Remove-Item -LiteralPath $script:StartupVbsPath -Force -ErrorAction SilentlyContinue }
    }
}

function Update-NotifyVisibility {
    if ($null -eq $script:NotifyIcon) { return }
    $script:NotifyIcon.Visible = [bool]$config.show_tray_icon
}

function New-OverlayDotBitmap {
    param([System.Drawing.Color]$Color,[int]$Size = 10)
    $safeSize = [Math]::Max(6, [int]$Size)
    $bmp = New-Object System.Drawing.Bitmap $safeSize,$safeSize
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $brush = New-Object System.Drawing.SolidBrush $Color
        try { $g.FillEllipse($brush, 1, 1, ($safeSize - 2), ($safeSize - 2)) } finally { $brush.Dispose() }
    }
    finally { $g.Dispose() }
    return $bmp
}

function Show-Overlay {
    param(
        [string]$Title,
        [string]$Message,
        [switch]$ErrorStyle,
        [string]$Indicator = 'Info',
        [switch]$Force
    )
    if ($null -ne $script:OverlayForm) {
        try { $script:OverlayForm.Close() } catch { Write-LogException -Context 'Close existing overlay' -ErrorRecord $_ }
        $script:OverlayForm = $null
    }
    $p = if ($null -ne $script:Palette) { $script:Palette } else { Get-ThemePalette }
    $indicatorColor = switch ($Indicator) {
        'On' { [System.Drawing.Color]::FromArgb(34,197,94) }
        'Off' { [System.Drawing.Color]::FromArgb(239,68,68) }
        'Warn' { [System.Drawing.Color]::FromArgb(245,158,11) }
        default { [System.Drawing.Color]::FromArgb(59,130,246) }
    }
    if ($ErrorStyle) { $indicatorColor = [System.Drawing.Color]::FromArgb(239,68,68) }

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.ShowInTaskbar = $false
    $form.StartPosition = 'Manual'
    $form.TopMost = $true
    $form.Width = 370
    $form.Height = 106
    $form.BackColor = if ($ErrorStyle) { $p.WarnBack } else { $p.OverlayBack }
    $form.Opacity = 0
    $form.Padding = New-Object System.Windows.Forms.Padding(1)
    try { $form.GetType().GetProperty('DoubleBuffered',[System.Reflection.BindingFlags]'Instance,NonPublic').SetValue($form,$true,$null) } catch { }

    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = 'Fill'
    $contentPanel.BackColor = $form.BackColor
    $contentPanel.Padding = New-Object System.Windows.Forms.Padding(14,12,14,10)
    $form.Controls.Add($contentPanel)

    $accentDot = New-Object System.Windows.Forms.PictureBox
    $accentDot.Location = New-Object System.Drawing.Point(14,17)
    $accentDot.Size = New-Object System.Drawing.Size(12,12)
    $accentDot.BackColor = [System.Drawing.Color]::Transparent
    $accentDot.SizeMode = 'CenterImage'
    $accentDot.Image = New-OverlayDotBitmap -Color $indicatorColor -Size 12
    $contentPanel.Controls.Add($accentDot)

    $paddingLeft = 32
    $paddingTop = 10
    $paddingRight = 14
    $contentWidth = $form.Width - $paddingLeft - $paddingRight - 6
    if ($contentWidth -lt 240) { $contentWidth = 240 }

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.AutoSize = $false
    $lblTitle.Location = New-Object System.Drawing.Point($paddingLeft, $paddingTop)
    $lblTitle.Size = New-Object System.Drawing.Size($contentWidth, 22)
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = if ($ErrorStyle) { $p.WarnText } else { $p.OverlayText }
    $lblTitle.UseCompatibleTextRendering = $false
    $lblTitle.TextAlign = 'MiddleLeft'
    $lblTitle.AutoEllipsis = $true
    $lblTitle.Text = $Title
    $contentPanel.Controls.Add($lblTitle)

    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.AutoSize = $false
    $lblMessage.Location = New-Object System.Drawing.Point($paddingLeft, ($paddingTop + 26))
    $lblMessage.Size = New-Object System.Drawing.Size($contentWidth, 44)
    $lblMessage.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $lblMessage.ForeColor = if ($ErrorStyle) { $p.WarnText } else { $p.MutedText }
    $lblMessage.UseCompatibleTextRendering = $false
    $lblMessage.TextAlign = 'TopLeft'
    $lblMessage.AutoEllipsis = $false
    $safeOverlayMessage = [string]$Message
    $safeOverlayMessage = $safeOverlayMessage -replace '\\r\\n', [Environment]::NewLine
    $safeOverlayMessage = $safeOverlayMessage -replace '\\n', [Environment]::NewLine
    $safeOverlayMessage = $safeOverlayMessage -replace '\\r', [Environment]::NewLine
    $lblMessage.Text = $safeOverlayMessage
    $contentPanel.Controls.Add($lblMessage)

    $anchorControl = if ($null -ne $script:MainForm) { $script:MainForm } else { $null }
    $work = if ($null -ne $anchorControl) { [System.Windows.Forms.Screen]::FromControl($anchorControl).WorkingArea } else { [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea }
    $marginRight = 24
    $marginBottom = 24
    $x = [Math]::Max($work.Left + 20, ($work.Right - $form.Width - $marginRight))
    $y = [Math]::Max($work.Top + 20, ($work.Bottom - $form.Height - $marginBottom))
    $form.Location = New-Object System.Drawing.Point($x, $y)
    $script:OverlayForm = $form

    $applyRoundedRegion = {
        param($TargetForm, [int]$Radius)
        if ($null -eq $TargetForm) { return }
        if ($TargetForm.Width -le 0 -or $TargetForm.Height -le 0) { return }
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $rect = New-Object System.Drawing.Rectangle(0,0,$TargetForm.Width,$TargetForm.Height)
        $diameter = [Math]::Max(2, ($Radius * 2))
        $path.AddArc($rect.X,$rect.Y,$diameter,$diameter,180,90)
        $path.AddArc($rect.Right-$diameter,$rect.Y,$diameter,$diameter,270,90)
        $path.AddArc($rect.Right-$diameter,$rect.Bottom-$diameter,$diameter,$diameter,0,90)
        $path.AddArc($rect.X,$rect.Bottom-$diameter,$diameter,$diameter,90,90)
        $path.CloseFigure()
        $TargetForm.Region = New-Object System.Drawing.Region($path)
        $path.Dispose()
    }.GetNewClosure()

    $form.Add_Paint({
        param($paintSender,$paintEvent)
        try {
            $paintEvent.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $rect = New-Object System.Drawing.Rectangle(0,0,($paintSender.Width - 1),($paintSender.Height - 1))
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $diameter = 22
            $path.AddArc($rect.X,$rect.Y,$diameter,$diameter,180,90)
            $path.AddArc($rect.Right-$diameter,$rect.Y,$diameter,$diameter,270,90)
            $path.AddArc($rect.Right-$diameter,$rect.Bottom-$diameter,$diameter,$diameter,0,90)
            $path.AddArc($rect.X,$rect.Bottom-$diameter,$diameter,$diameter,90,90)
            $path.CloseFigure()
            $pen = New-Object System.Drawing.Pen($p.Border,1)
            try { $paintEvent.Graphics.DrawPath($pen,$path) } finally { $pen.Dispose(); $path.Dispose() }
        }
        catch { }
    }.GetNewClosure())

    $targetOpacity = [math]::Min([math]::Max(([double]$config.overlay_opacity / 100.0), 0.0), 1.0)
    $form.Opacity = 0.0
    $fadeOut = New-Object System.Windows.Forms.Timer
    $fadeOut.Interval = 10
    $hold = New-Object System.Windows.Forms.Timer
    $hold.Interval = [math]::Max(1, [int]([double]$config.overlay_seconds * 1000.0))

    $hold.Add_Tick({
        $hold.Stop()
        $fadeOut.Start()
    }.GetNewClosure())
    $fadeOut.Add_Tick({
        $form.Opacity -= 0.35
        if ($form.Opacity -le 0) {
            $fadeOut.Stop()
            try { $form.Close() } catch { Write-LogException -Context 'Close overlay form' -ErrorRecord $_ }
        }
    }.GetNewClosure())
    $form.Add_Shown({
        & $applyRoundedRegion $form 11
        try { $form.Opacity = $targetOpacity } catch { }
        $hold.Start()
    }.GetNewClosure())
    $form.Add_SizeChanged({ & $applyRoundedRegion $form 11 }.GetNewClosure())
    $form.Add_FormClosed({
        try { $fadeOut.Stop(); $hold.Stop() } catch { Write-LogException -Context 'Stop overlay timers' -ErrorRecord $_ }
        try { if ($null -ne $accentDot -and $null -ne $accentDot.Image) { $accentDot.Image.Dispose() } } catch { }
        if ($script:OverlayForm -eq $form) { $script:OverlayForm = $null }
    }.GetNewClosure())
    $form.Show()
    $form.Refresh()
}

function Get-ModifierValue {
    param([string]$Text)
    switch ($Text) {
        'Alt' { 1 }
        'Ctrl' { 2 }
        'Shift' { 4 }
        'Ctrl+Alt' { 3 }
        'Ctrl+Shift' { 6 }
        'Alt+Shift' { 5 }
        'Ctrl+Alt+Shift' { 7 }
        'Win' { 8 }
        'Ctrl+Win' { 10 }
        'Alt+Win' { 9 }
        'None' { 0 }
        default { 3 }
    }
}

function Get-KeyValue {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $normalized = [string]$Text
    if ($normalized -match '^[0-9]$') { $normalized = 'D' + $normalized }
    try { return [int][System.Windows.Forms.Keys]::$normalized } catch { return $null }
}

function Get-ModifierTextFromKeyEvent {
    param([System.Windows.Forms.KeyEventArgs]$keyEventArgs)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($keyEventArgs.Control) { [void]$parts.Add('Ctrl') }
    if ($keyEventArgs.Alt) { [void]$parts.Add('Alt') }
    if ($keyEventArgs.Shift) { [void]$parts.Add('Shift') }
    if ($parts.Count -eq 0) { return 'None' }
    return ($parts -join '+')
}

function Get-KeyTextFromKeyCode {
    param([System.Windows.Forms.Keys]$KeyCode)
    $baseKey = ($KeyCode -band [System.Windows.Forms.Keys]::KeyCode)
    $name = [string]$baseKey
    if ($name -match '^D([0-9])$') { return $Matches[1] }
    if ($name -match '^[A-Z]$') { return $name }
    if ($name -match '^F([1-9]|1[0-2])$') { return $name }
    return $null
}

function Test-KeyDown {
    param([System.Windows.Forms.Keys]$Key)
    try {
        return (([int][TailscaleControlKeyboard]::GetAsyncKeyState([int]$Key) -band 0x8000) -ne 0)
    }
    catch {
        return $false
    }
}

function Test-AnyKeyDown {
    param([System.Windows.Forms.Keys[]]$Keys)
    foreach ($k in @($Keys)) {
        if (Test-KeyDown -Key $k) { return $true }
    }
    return $false
}

function Test-HotkeyStillPressed {
    param([string]$Name)
    try {
        $cfg = Get-Config
        if ($null -eq $cfg -or $null -eq $cfg.hotkeys) { return $false }
        $entry = $cfg.hotkeys.PSObject.Properties[$Name].Value
        if ($null -eq $entry) { return $false }
        $modText = [string]$entry.modifiers
        $keyText = [string]$entry.key
        $hasAny = $false
        if ($modText -match 'Ctrl') {
            $hasAny = $true
            if (-not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::ControlKey,[System.Windows.Forms.Keys]::LControlKey,[System.Windows.Forms.Keys]::RControlKey))) { return $false }
        }
        if ($modText -match 'Alt') {
            $hasAny = $true
            if (-not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::Menu,[System.Windows.Forms.Keys]::LMenu,[System.Windows.Forms.Keys]::RMenu))) { return $false }
        }
        if ($modText -match 'Shift') {
            $hasAny = $true
            if (-not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::ShiftKey,[System.Windows.Forms.Keys]::LShiftKey,[System.Windows.Forms.Keys]::RShiftKey))) { return $false }
        }
        $vk = Get-KeyValue -Text $keyText
        if ($null -ne $vk) {
            $hasAny = $true
            if (-not (Test-KeyDown -Key ([System.Windows.Forms.Keys]$vk))) { return $false }
        }
        return $hasAny
    }
    catch {
        return $false
    }
}

function Clear-HotkeyExecutionLock {
    $script:HotkeyReleaseTicks = 0
    $script:HotkeyExecutionLock = $false
    $script:HotkeyExecutionName = $null
}

function Initialize-HotkeyReleaseMonitor {
    if ($null -ne $script:HotkeyExecutionTimer) { return }
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 20
    $timer.add_Tick({
        param($ts,$te)
        if (-not $script:HotkeyExecutionLock) {
            $script:HotkeyReleaseTicks = 0
            return
        }
        $name = [string]$script:HotkeyExecutionName
        if ([string]::IsNullOrWhiteSpace($name)) {
            Clear-HotkeyExecutionLock
            return
        }
        $stillPressed = $false
        try { $stillPressed = Test-HotkeyStillPressed -Name $name } catch { $stillPressed = $false }
        if ($stillPressed) {
            $script:HotkeyReleaseTicks = 0
            return
        }
        $script:HotkeyReleaseTicks++
        if ($script:HotkeyReleaseTicks -ge 2) {
            Clear-HotkeyExecutionLock
        }
    })
    $script:HotkeyExecutionTimer = $timer
    $timer.Start()
}

function Start-HotkeyReleaseMonitor {
    param([string]$Name)
    Initialize-HotkeyReleaseMonitor
    $script:HotkeyReleaseTicks = 0
    $script:HotkeyExecutionLock = $true
    $script:HotkeyExecutionName = $Name
}

function Unregister-Hotkeys {
    if ($null -eq $script:HotkeyWindow) { return }
    foreach ($item in @($script:RegisteredHotkeys.GetEnumerator())) {
        try { [void][TailscaleControlHotkeys]::UnregisterHotKey($script:HotkeyWindow.WindowHandle, [int]$item.Value) } catch { Write-LogException -Context 'Unregister hotkey' -ErrorRecord $_ }
    }
    $script:RegisteredHotkeys.Clear()
}

function Register-Hotkeys {
    param($Config)
    if ($null -eq $script:HotkeyWindow) { return @() }
    if ($null -ne $Config -and $null -ne $Config.hotkeys) {
        $quickCount = Get-QuickAccountSwitchCountFromHotkeys -Hotkeys $Config.hotkeys
        try { $quickCount = [Math]::Max([int]$quickCount, [int]@($script:QuickAccountSwitchAccounts).Count) } catch { }
        Ensure-ConfigHotkeyEntries -Config $Config -Count $quickCount
    }
    Unregister-Hotkeys
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:HotkeyNames) {
        if ($null -eq $Config -or $null -eq $Config.hotkeys -or $null -eq $Config.hotkeys.PSObject.Properties[$name]) { continue }
        $entry = $Config.hotkeys.PSObject.Properties[$name].Value
        $quickIndex = Get-QuickAccountSwitchIndex -Name $name
        if ($quickIndex -gt 0 -and -not [bool]$script:QuickAccountSwitchAvailable) { continue }
        if (-not [bool]$entry.enabled) { continue }
        $modifier = Get-ModifierValue ([string]$entry.modifiers)
        $key = Get-KeyValue ([string]$entry.key)
        if ($null -eq $key) {
            [void]$errors.Add($script:ActionLabels[$name] + ' has an invalid key.')
            continue
        }
        $id = [int]$script:HotkeyIds[$name]
        $ok = [TailscaleControlHotkeys]::RegisterHotKey($script:HotkeyWindow.WindowHandle, $id, [uint32]$modifier, [uint32]$key)
        if ($ok) {
            $script:RegisteredHotkeys[$name] = $id
        }
        else {
            $message = $script:ActionLabels[$name] + ' could not be registered.'
            [void]$errors.Add($message)
            try { Write-Log ('Hotkey registration failed: ' + $message) } catch { }
        }
    }
    return @($errors)
}

function Start-ShowSettingsHotkeyFallback {
    if ($null -ne $script:HotkeyPollTimer) { return }
    $script:HotkeyPollTimer = New-Object System.Windows.Forms.Timer
    $script:HotkeyPollTimer.Interval = 15
    $script:HotkeyPollTimer.add_Tick({
        try {
            if ($script:IsCapturingHotkey -or $script:HotkeyExecutionLock) { return }
            $cfg = Get-Config
            if ($null -eq $cfg -or $null -eq $cfg.hotkeys) { return }
            $entry = Get-ObjectPropertyOrDefault $cfg.hotkeys 'ShowSettings' $null
            if ($null -eq $entry -or -not [bool](Get-ObjectPropertyOrDefault $entry 'enabled' $false)) { return }
            $modText = [string](Get-ObjectPropertyOrDefault $entry 'modifiers' 'Ctrl+Alt')
            $keyText = [string](Get-ObjectPropertyOrDefault $entry 'key' 'O')
            if ([string]::IsNullOrWhiteSpace($keyText)) { return }
            if ($modText -match 'Ctrl' -and -not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::ControlKey,[System.Windows.Forms.Keys]::LControlKey,[System.Windows.Forms.Keys]::RControlKey))) { return }
            if ($modText -match 'Alt' -and -not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::Menu,[System.Windows.Forms.Keys]::LMenu,[System.Windows.Forms.Keys]::RMenu))) { return }
            if ($modText -match 'Shift' -and -not (Test-AnyKeyDown -Keys @([System.Windows.Forms.Keys]::ShiftKey,[System.Windows.Forms.Keys]::LShiftKey,[System.Windows.Forms.Keys]::RShiftKey))) { return }
            $vk = Get-KeyValue -Text $keyText
            if ($null -eq $vk) { return }
            if (-not (Test-KeyDown -Key ([System.Windows.Forms.Keys]$vk))) { return }
            Invoke-HotkeyAction -Name 'ShowSettings'
        }
        catch { Write-LogException -Context 'Show settings hotkey fallback' -ErrorRecord $_ }
    })
    $script:HotkeyPollTimer.Start()
}

function Set-TrayMenuItemState {
    param(
        $Item,
        [string]$Label,
        [string]$State,
        [bool]$Checked,
        [bool]$Enabled = $true,
        [string]$Detail = ''
    )
    if ($null -eq $Item) { return }
    Set-TrayMenuItemFixedSize -Item $Item -BaseWidth 300 -BaseHeight 28
    $suffix = if ([string]::IsNullOrWhiteSpace([string]$State)) { '' } else { ' [' + $State + ']' }
    $detailSuffix = if ([string]::IsNullOrWhiteSpace([string]$Detail)) { '' } else { ' (' + [string]$Detail + ')' }
    $Item.Text = $Label + $suffix + $detailSuffix
    $Item.Checked = [bool]$Checked
    $Item.Enabled = [bool]$Enabled
    $known = [bool]($Enabled -and [string]$State -notmatch '^(Unknown|Not detected)$')
    $on = [bool]$Checked
    if ($State -match '^(On|Allowed|Connected)$') { $on = $true }
    if ($State -match '^(Off|Blocked|Disconnected)$') { $on = $false }
    Set-TrayMenuItemDot -Item $Item -Known:$known -On:$on
}

function Get-TrayShortText {
    param([string]$Text,[int]$Max = 64)
    $value = [string]$Text
    if ([string]::IsNullOrWhiteSpace($value)) { return '-' }
    $value = $value.Trim()
    if ($value.Length -le $Max) { return $value }
    if ($Max -le 3) { return $value.Substring(0,$Max) }
    return ($value.Substring(0,($Max - 3)) + '...')
}

function Get-TrayMutedColor {
    return [System.Drawing.Color]::FromArgb(156,163,175)
}

function New-TrayCopyGlyphBitmap {
    $bmp = New-Object System.Drawing.Bitmap 14,14
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $pen = New-Object System.Drawing.Pen (Get-TrayMutedColor), 1.4
        try {
            $g.DrawRectangle($pen, 5, 3, 6, 7)
            $g.DrawRectangle($pen, 3, 5, 6, 7)
        }
        finally { $pen.Dispose() }
    }
    finally { $g.Dispose() }
    return $bmp
}

function Set-TrayCopyInfoMenuItem {
    param($Item,[string]$Label,[string]$Value)
    if ($null -eq $Item) { return }
    try {
        Set-TrayMenuItemFixedSize -Item $Item -BaseWidth 300 -BaseHeight 28
        $cleanValue = if ([string]::IsNullOrWhiteSpace([string]$Value)) { '-' } else { [string]$Value }
        $canCopy = -not ([string]::IsNullOrWhiteSpace($cleanValue) -or $cleanValue -eq '-')
        $Item.Text = $Label + ': ' + (Get-TrayShortText -Text $cleanValue -Max 72)
        $Item.Tag = $cleanValue
        $Item.Enabled = $true
        $Item.ForeColor = [System.Drawing.Color]::FromArgb(107,114,128)
        $Item.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $Item.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $Item.ToolTipText = $(if ($canCopy) { 'Click to copy ' + $Label } else { '' })
        try { $Item.Invalidate() } catch { }
        if ($null -eq $Item.Image) { $Item.Image = New-TrayCopyGlyphBitmap }
    }
    catch { }
}

function Copy-TrayMenuValue {
    param($Item,[string]$Label)
    if ($null -eq $Item) { return }
    try {
        $value = [string]$Item.Tag
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return }
        [System.Windows.Forms.Clipboard]::SetText($value)
        $oldText = [string]$Item.Text
        $Item.Text = $Label + ': copied'
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 900
        $restoreItem = $Item
        $restoreText = $oldText
        $timer.add_Tick({
            param($timerSender,$timerEvent)
            try { $timerSender.Stop() } catch { }
            try { $timerSender.Dispose() } catch { }
            try { if ($null -ne $restoreItem) { $restoreItem.Text = $restoreText } } catch { }
        }.GetNewClosure())
        $timer.Start()
    }
    catch { Write-LogException -Context ('Copy tray ' + $Label) -ErrorRecord $_ }
}

function Register-TrayCopyMenuItem {
    param($Item,[string]$Label)
    if ($null -eq $Item) { return }
    try {
        $labelForCopy = [string]$Label
        $Item.add_Click({
            param($menuSender,$menuEvent)
            try { Copy-TrayMenuValue -Item $menuSender -Label $labelForCopy }
            catch { Write-LogException -Context ('Tray copy click ' + $labelForCopy) -ErrorRecord $_ }
        }.GetNewClosure())
        $Item.add_MouseUp({
            param($menuSender,$mouseEvent)
            try {
                if ($mouseEvent.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                    Copy-TrayMenuValue -Item $menuSender -Label $labelForCopy
                }
            }
            catch { Write-LogException -Context ('Tray copy mouse up ' + $labelForCopy) -ErrorRecord $_ }
        }.GetNewClosure())
    }
    catch { Write-LogException -Context ('Register tray copy ' + $Label) -ErrorRecord $_ }
}

function Set-TrayMenuItemFixedSize {
    param($Item,[int]$BaseWidth = 300,[int]$BaseHeight = 28)
    if ($null -eq $Item) { return }
    try {
        $width = [Math]::Max(260,[int]$BaseWidth)
        $height = [Math]::Max(24,[int]$BaseHeight)
        $Item.AutoSize = $false
        $Item.Width = $width
        $Item.Height = $height
        $Item.Size = New-Object System.Drawing.Size -ArgumentList $width,$height
        $Item.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $Item.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $Item.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
        try { $Item.ImageScaling = [System.Windows.Forms.ToolStripItemImageScaling]::None } catch { }
    }
    catch { }
}

function Register-TraySubmenuArrowGlyph {
    param($Item)
    if ($null -eq $Item) { return }
    try {
        $marker = [string]$Item.AccessibleDescription
        if ($marker -like '*ManualTraySubmenuArrowGlyph*') { return }
        if ([string]::IsNullOrWhiteSpace($marker)) { $Item.AccessibleDescription = 'ManualTraySubmenuArrowGlyph' } else { $Item.AccessibleDescription = ($marker + ';ManualTraySubmenuArrowGlyph') }
        $Item.add_Paint({
            param($paintSender,$paintEvent)
            try {
                if ($null -eq $paintSender -or $null -eq $paintEvent) { return }
                $g = $paintEvent.Graphics
                $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $color = if ($paintSender.Enabled) { [System.Drawing.SystemColors]::MenuText } else { [System.Drawing.SystemColors]::GrayText }
                $brush = New-Object System.Drawing.SolidBrush $color
                try {
                    $right = [int]$paintSender.Bounds.Width
                    if ($right -le 0) { $right = 300 }
                    $x = [Math]::Max(160, $right - 18)
                    $y = [Math]::Max(0, [int](($paintSender.Bounds.Height - 8) / 2))
                    $points = [System.Drawing.Point[]]@(
                        (New-Object System.Drawing.Point($x, $y)),
                        (New-Object System.Drawing.Point($x, ($y + 8))),
                        (New-Object System.Drawing.Point(($x + 5), ($y + 4)))
                    )
                    $g.FillPolygon($brush, $points)
                }
                finally { $brush.Dispose() }
            }
            catch { }
        }.GetNewClosure())
        try { $Item.Invalidate() } catch { }
    }
    catch { }
}

function Set-TrayDropDownRenderer {
    param($Item)
    try {
        if ($null -eq $Item) { return }
        $renderer = $null
        try { if ($null -ne $script:TrayContextMenu -and $null -ne $script:TrayContextMenu.Renderer) { $renderer = $script:TrayContextMenu.Renderer } } catch { }
        try { if ($null -eq $renderer -and $null -ne $script:TrayMenuRenderer) { $renderer = $script:TrayMenuRenderer } } catch { }
        if ($null -ne $renderer) { $Item.DropDown.Renderer = $renderer }
    }
    catch { }
}



function Update-TrayMenuFixedLayout {
    param($Menu)
    try {
        if ($null -eq $Menu) { return }
        $Menu.MinimumSize = New-Object System.Drawing.Size(300,0)
        try {
            if ($null -eq $Menu.Font -or [Math]::Abs([double]$Menu.Font.SizeInPoints - 9.0) -gt 0.1) {
                $Menu.Font = New-Object System.Drawing.Font('Segoe UI',9.0,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point)
            }
        } catch { }
        foreach ($item in @($Menu.Items)) {
            try {
                if ($item -is [System.Windows.Forms.ToolStripSeparator]) { continue }
                if ($item -eq $script:TrayMenuSelectAccount -or $item -eq $script:TrayMenuChooseExitNode) {
                    Set-TrayMenuItemFixedSize -Item $item -BaseWidth 300 -BaseHeight 28
                }
                else {
                    Set-TrayMenuItemFixedSize -Item $item -BaseWidth 300 -BaseHeight 28
                }
            } catch { }
        }
    }
    catch { Write-LogException -Context 'Update tray menu fixed layout' -ErrorRecord $_ }
}

function Register-TrayVersionSuffix {
    param($Item)
    if ($null -eq $Item) { return }
    try {
        Set-TrayMenuItemFixedSize -Item $Item -BaseWidth 300 -BaseHeight 28
        $Item.add_Paint({
            param($paintSender,$paintEvent)
            try {
                $versionText = [string]$paintSender.Tag
                if ([string]::IsNullOrWhiteSpace($versionText)) { return }
                $brush = New-Object System.Drawing.SolidBrush (Get-TrayMutedColor)
                try {
                    $paintEvent.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
                    $right = [int]$paintSender.Bounds.Width
                    if ($right -le 0) { $right = 300 }
                    $rect = New-Object System.Drawing.Rectangle(150,0,($right - 162),$paintSender.Bounds.Height)
                    [System.Windows.Forms.TextRenderer]::DrawText($paintEvent.Graphics, $versionText, $paintSender.Font, $rect, $brush.Color, ([System.Windows.Forms.TextFormatFlags]::Right -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor [System.Windows.Forms.TextFormatFlags]::NoPrefix))
                }
                finally { $brush.Dispose() }
            }
            catch { }
        })
    }
    catch { }
}

function Register-TrayExternalLinkGlyph {
    param($Item)
    if ($null -eq $Item) { return }
    try {
        Set-TrayMenuItemFixedSize -Item $Item -BaseWidth 300 -BaseHeight 28
        $Item.ToolTipText = 'Open Tailscale Admin Panel in your browser'
        $Item.add_Paint({
            param($paintSender,$paintEvent)
            try {
                if ($null -eq $paintSender -or $null -eq $paintEvent) { return }
                $g = $paintEvent.Graphics
                $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $color = Get-TrayMutedColor
                $pen = New-Object System.Drawing.Pen $color, 1.6
                try {
                    $right = [int]$paintSender.Bounds.Width
                    if ($right -le 0) { $right = 300 }
                    $box = 16
                    $x = [Math]::Max(160, $right - 26)
                    $y = [Math]::Max(0, [int](($paintSender.Bounds.Height - $box) / 2))

                    $g.DrawLine($pen, $x + 4, $y + 12, $x + 12, $y + 4)
                    $g.DrawLine($pen, $x + 8, $y + 4, $x + 12, $y + 4)
                    $g.DrawLine($pen, $x + 12, $y + 4, $x + 12, $y + 8)
                }
                finally { $pen.Dispose() }
            }
            catch { }
        })
    }
    catch { }
}

function Set-TrayCurrentDeviceInfoVisibility {
    param([bool]$Visible)
    try {
        foreach ($item in @($script:TrayMenuInfoUser,$script:TrayMenuInfoAccountEmail,$script:TrayMenuInfoTailnet,$script:TrayMenuInfoStatus,$script:TrayMenuInfoDevice,$script:TrayMenuInfoMagicDns,$script:TrayMenuInfoIPv4,$script:TrayMenuInfoIPv6,$script:TrayMenuInfoDns,$script:TrayMenuInfoTailscaleVersion)) {
            if ($null -ne $item) { $item.Visible = [bool]$Visible }
        }
        if ($null -ne $script:TrayMenuInfoTopSeparator) { $script:TrayMenuInfoTopSeparator.Visible = [bool]$Visible }
        if ($null -ne $script:TrayMenuInfoBottomSeparator) { $script:TrayMenuInfoBottomSeparator.Visible = $true }
    }
    catch { }
}

function Update-TrayInfoMenuState {
    param($Snapshot)
    try {
        $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
        $tailnet = '-'
        $status = '-'
        $user = '-'
        $accountEmail = '-'
        $device = '-'
        $tailscaleVersion = '-'
        $magicDns = '-'
        $ipv4 = '-'
        $ipv6 = '-'
        $dns = '-'
        $mtuDetected = $false
        $mtuVersion = ''

        try {
            if ($null -ne $script:TrayMenuShow) {
                $script:TrayMenuShow.Text = 'Tailscale Control'
                $script:TrayMenuShow.Tag = [string]$script:AppVersion
                $script:TrayMenuShow.ToolTipText = 'Open Tailscale Control'
                Set-TrayTitleIcon
                try { $script:TrayMenuShow.Invalidate() } catch { }
                $script:TrayMenuShow.Enabled = $true
            }
        }
        catch { }

        if ($null -ne $snap) {
            $found = [bool](Get-ObjectPropertyOrDefault $snap 'Found' $false)
            $backendState = [string](Get-ObjectPropertyOrDefault $snap 'BackendState' '')
            $tailnet = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'Tailnet' ''))
            $user = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $snap 'User' ''))
            $accountEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $snap 'UserEmail' ''))
            if ([string]::IsNullOrWhiteSpace($accountEmail)) { $accountEmail = $user }
            if ([string]::IsNullOrWhiteSpace($user)) { $user = $accountEmail }
            $device = [string](Get-ObjectPropertyOrDefault $snap 'ShortName' '')
            if ([string]::IsNullOrWhiteSpace($device)) { $device = [string](Get-ObjectPropertyOrDefault $snap 'DNSName' '') }
            $status = if (-not $found) { 'Not detected' } elseif ($backendState -eq 'Running') { 'Connected' } else { 'Disconnected' }
            $tailscaleVersion = ConvertTo-PlainVersion ([string](Get-ObjectPropertyOrDefault $snap 'Version' ''))
            $magicDns = [string](Get-ObjectPropertyOrDefault $snap 'DNSName' '')
            $ipv4 = [string](Get-ObjectPropertyOrDefault $snap 'IPv4' '')
            $ipv6 = [string](Get-ObjectPropertyOrDefault $snap 'IPv6' '')

            $dnsNameservers = [string](Get-ObjectPropertyOrDefault $snap 'DnsNameservers' '')
            $corpDnsValue = Get-ObjectPropertyOrDefault $snap 'CorpDNS' $null
            $dns = Get-ActiveDnsDisplayText -Nameservers $dnsNameservers -CorpDns $corpDnsValue -PreferSystem:($backendState -ne 'Running')
            if ([string]::IsNullOrWhiteSpace($dns)) { $dns = [string](Get-ObjectPropertyOrDefault $snap 'DnsSummary' '') }
            $mtuDetected = [bool](Get-ObjectPropertyOrDefault $snap 'MtuInstalled' $false)
            $mtuVersion = [string](Get-ObjectPropertyOrDefault $snap 'MtuVersion' '')
        }
        else {
            try {
                $mtuInfo = Get-TailscaleMtuAppInfo
                $mtuDetected = [bool]($null -ne $mtuInfo -and [bool]$mtuInfo.Installed)
                if ($mtuDetected) { $mtuVersion = [string]$mtuInfo.Version }
            }
            catch { $mtuDetected = $false }
        }

        try {
            if ($null -ne $script:TrayMenuMtu) {
                $script:TrayMenuMtu.Visible = $true
                $script:TrayMenuMtu.Text = 'Tailscale MTU'
                $script:TrayMenuMtu.ToolTipText = 'Open Tailscale MTU'
                if ($mtuDetected) {
                    $cleanMtuVersion = ConvertTo-PlainVersion ([string]$mtuVersion)
                    $script:TrayMenuMtu.Tag = $(if ([string]::IsNullOrWhiteSpace($cleanMtuVersion)) { '' } else { [string]$cleanMtuVersion })
                    Set-TrayMtuIcon
                    try { $script:TrayMenuMtu.Invalidate() } catch { }
                    $script:TrayMenuMtu.Enabled = $true
                }
                else {
                    $script:TrayMenuMtu.Tag = ''
                    Set-TrayMtuIcon
                    try { $script:TrayMenuMtu.Invalidate() } catch { }
                    $script:TrayMenuMtu.Enabled = $false
                }
            }
        }
        catch { }

        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoUser -Label 'User' -Value $user
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoAccountEmail -Label 'Account Email' -Value $accountEmail
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoTailnet -Label 'Tailnet' -Value $tailnet
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoStatus -Label 'Status' -Value $status
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoMagicDns -Label 'MagicDNS' -Value $magicDns
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoDevice -Label 'Device' -Value $device
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoIPv4 -Label 'IPv4' -Value $ipv4
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoIPv6 -Label 'IPv6' -Value $ipv6
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoDns -Label 'DNS' -Value $dns
        Set-TrayCopyInfoMenuItem -Item $script:TrayMenuInfoTailscaleVersion -Label 'Tailscale Version' -Value $tailscaleVersion
        $showDeviceInfo = [bool](Get-ObjectPropertyOrDefault (Get-Config) 'show_current_device_info_in_tray' $false)
        Set-TrayCurrentDeviceInfoVisibility -Visible:$showDeviceInfo
    }
    catch { Write-LogException -Context 'Update tray info menu state' -ErrorRecord $_ }
}

function Set-TrayExitNodeSelection {
    param([string]$Label,[switch]$Clear)
    try {
        $cfg = Get-Config
        if ($Clear) {
            $cfg.preferred_exit_label = ''
            $cfg.preferred_exit_node = ''
            $previousLoading = $script:IsLoadingConfig
            try {
                $script:IsLoadingConfig = $true
                if ($null -ne $script:cmbExitNode) { $script:cmbExitNode.SelectedIndex = -1 }
            }
            finally { $script:IsLoadingConfig = $previousLoading }
            Save-Config -Config $cfg
            Write-Log 'Preferred exit node cleared from tray.'
            Update-TrayExitNodeMenuState -Snapshot $script:Snapshot
            return
        }

        $selectedLabel = ConvertTo-DnsName ([string]$Label)
        if ([string]::IsNullOrWhiteSpace($selectedLabel)) { throw 'No preferred exit node was selected.' }
        $cfg.preferred_exit_label = $selectedLabel
        $cfg.preferred_exit_node = $selectedLabel
        $previousLoading = $script:IsLoadingConfig
        try {
            $script:IsLoadingConfig = $true
            if ($null -ne $script:cmbExitNode -and $script:cmbExitNode.Items.Contains($selectedLabel)) {
                $script:cmbExitNode.SelectedItem = $selectedLabel
            }
        }
        finally { $script:IsLoadingConfig = $previousLoading }
        Save-Config -Config $cfg
        Write-Log ('Preferred exit node set from tray: ' + $selectedLabel)
        Update-TrayExitNodeMenuState -Snapshot $script:Snapshot
    }
    catch { Write-LogException -Context 'Set preferred exit node from tray' -ErrorRecord $_ }
}

function Get-TailscaleAccountOverlayTailnet {
    param($Account)
    $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Identifier' ''))
    $details = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Details' ''))
    $raw = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Raw' ''))
    foreach ($candidate in @($identifier,$details,$raw)) {
        if ([string]::IsNullOrWhiteSpace([string]$candidate)) { continue }
        $m = [regex]::Match([string]$candidate, '(?i)([a-z0-9][a-z0-9\-]{1,62}\.ts\.net)')
        if ($m.Success) { return [string]$m.Groups[1].Value }
    }
    try {
        $snapTailnet = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $script:Snapshot 'Tailnet' ''))
        if (-not [string]::IsNullOrWhiteSpace($snapTailnet)) { return $snapTailnet }
    } catch { }
    if (-not [string]::IsNullOrWhiteSpace($identifier) -and $identifier -notmatch '@') { return $identifier }
    return '-'
}

function Get-TailscaleAccountPreferredEmail {
    param($Account)
    if ($null -eq $Account) { return '' }
    $emailPattern = '[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}'
    foreach ($prop in @('QuickUserEmail','User','Email','AccountEmail','LoginName','Login','QuickTailnetEmail','Identifier','SwitchIdentifier')) {
        $value = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account $prop ''))
        if ($value -match $emailPattern) { return [string]$Matches[0] }
    }
    foreach ($prop in @('QuickDisplay','Details','Raw')) {
        $value = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account $prop ''))
        $matches = @([regex]::Matches($value, $emailPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))
        if ($matches.Count -gt 1) { return [string]$matches[$matches.Count - 1].Value }
        if ($matches.Count -eq 1) { return [string]$matches[0].Value }
    }
    return ''
}

function Get-TailscaleAccountOverlayIdentifier {
    param($Account)
    $preferredEmail = Get-TailscaleAccountPreferredEmail -Account $Account
    if (-not [string]::IsNullOrWhiteSpace($preferredEmail)) { return $preferredEmail }
    $quickDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickDisplay' ''))
    if (-not [string]::IsNullOrWhiteSpace($quickDisplay)) { return $quickDisplay }
    $user = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'User' ''))
    $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Identifier' ''))
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'SwitchIdentifier' '')) }
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = $user }
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = '-' }
    return $identifier
}

function Get-TailscaleAccountSwitchOutputIdentifier {
    param($Account)
    if ($null -eq $Account) { return '' }
    $quickDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickDisplay' ''))
    if (-not [string]::IsNullOrWhiteSpace($quickDisplay)) { return $quickDisplay }
    $tailnetEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickTailnetEmail' ''))
    $userEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickUserEmail' ''))
    if (-not [string]::IsNullOrWhiteSpace($tailnetEmail) -and -not [string]::IsNullOrWhiteSpace($userEmail)) { return ($tailnetEmail + ' | ' + $userEmail) }
    if (-not [string]::IsNullOrWhiteSpace($tailnetEmail)) { return $tailnetEmail }
    $preferredEmail = Get-TailscaleAccountPreferredEmail -Account $Account
    if (-not [string]::IsNullOrWhiteSpace($preferredEmail)) { return $preferredEmail }
    $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Identifier' ''))
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'SwitchIdentifier' '')) }
    return $identifier
}

function Get-CurrentTailnetForAccountOverlay {
    param($FallbackAccount = $null)
    try {
        $tailnet = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $script:Snapshot 'Tailnet' ''))
        if ($tailnet -match '(?i)^[a-z0-9][a-z0-9\-]{1,62}\.ts\.net$') { return $tailnet }
    } catch { }
    try {
        $dnsName = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $script:Snapshot 'DNSName' ''))
        $m = [regex]::Match($dnsName, '(?i)([a-z0-9][a-z0-9\-]{1,62}\.ts\.net)')
        if ($m.Success) { return [string]$m.Groups[1].Value }
    } catch { }
    try {
        if ($null -ne $FallbackAccount) {
            $candidate = Get-TailscaleAccountOverlayTailnet -Account $FallbackAccount
            if ($candidate -match '(?i)^[a-z0-9][a-z0-9\-]{1,62}\.ts\.net$') { return $candidate }
        }
    } catch { }
    return '-'
}

function Set-AccountNoConnectedState {
    param([string]$Message = 'No account is currently connected. Double-click an account to switch.')
    try {
        if ($null -ne $script:lblAccountIdentifier) { Set-UiValue $script:lblAccountIdentifier 'No account connected' }
        if ($null -ne $script:lblAccountEmail) { Set-UiValue $script:lblAccountEmail '-' }
        if ($null -ne $script:lblAccountActiveUser) { Set-UiValue $script:lblAccountActiveUser 'No account connected' }
        if ($null -ne $script:lblAccountTailnet) { Set-UiValue $script:lblAccountTailnet '-' }
        if ($null -ne $script:lblAccountDevice) { Set-UiValue $script:lblAccountDevice '-' }
        if ($null -ne $script:lblAccountDnsName) { Set-UiValue $script:lblAccountDnsName '-' }
        if ($null -ne $script:lblAccountVisibleDevices) { Set-UiValue $script:lblAccountVisibleDevices '0' }
        if ($null -ne $script:lblAccountVisibleUsers) { Set-UiValue $script:lblAccountVisibleUsers '0' }
        if ($null -ne $script:lblAccountListStatus) { Set-UiValue $script:lblAccountListStatus $Message }
        if ($null -ne $script:gridAccounts) {
            foreach ($row in @($script:gridAccounts.Rows)) {
                try {
                    if ($null -ne $row.Tag) {
                        $row.Tag.Active = $false
                        $row.Tag.Status = 'Disconnected'
                    }
                    $row.Cells['AccountStatus'].Value = 'Disconnected'
                    Set-AccountRowVisualState -Row $row -State 'Disconnected'
                } catch { }
            }
        }
    } catch { }
}

function Start-DelayedAccountRefresh {
    param([int]$Milliseconds = 3000)
    try {
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = [Math]::Max(500, [int]$Milliseconds)
        $timer.add_Tick({
            param($sender,$eventArgs)
            try { $sender.Stop(); $sender.Dispose() } catch { }
            try { Reset-SlowSnapshotCache } catch { }
            try { Update-Status -RefreshAccounts } catch { Write-LogException -Context 'Delayed account refresh' -ErrorRecord $_ }
        })
        $timer.Start()
    } catch { }
}

function Start-SwitchAccountCompletionOverlay {
    param(
        [string]$Identifier,
        [string]$PreviousTailnet,
        [string]$Title = 'Tailscale account switched',
        [string]$Indicator = 'Info'
    )
    try {
        $state = [pscustomobject]@{
            Identifier = [string]$Identifier
            PreviousTailnet = [string]$PreviousTailnet
            Title = [string]$Title
            Indicator = [string]$Indicator
            Attempts = 0
        }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 750
        $timer.add_Tick({
            param($sender,$eventArgs)
            try {
                $state.Attempts = [int]$state.Attempts + 1
                try { Reset-SlowSnapshotCache } catch { }
                try { Update-Status -RefreshAccounts } catch { Write-LogException -Context 'Refresh after switch account' -ErrorRecord $_ }
                $tailnet = Get-CurrentTailnetForAccountOverlay
                $previous = [string]$state.PreviousTailnet
                $ready = ($tailnet -match '(?i)^[a-z0-9][a-z0-9\-]{1,62}\.ts\.net$')
                if ($ready -and -not [string]::IsNullOrWhiteSpace($previous) -and $previous -ne '-' -and $tailnet -eq $previous -and [int]$state.Attempts -lt 4) {
                    $ready = $false
                }
                if ($ready -or [int]$state.Attempts -ge 8) {
                    try { $sender.Stop(); $sender.Dispose() } catch { }
                    $id = ConvertTo-DiagnosticText -Text ([string]$state.Identifier)
                    if ([string]::IsNullOrWhiteSpace($id)) { $id = 'selected account' }
                    if ([string]::IsNullOrWhiteSpace($tailnet)) { $tailnet = '-' }
                    $message = 'Switched to "' + $id + '".' + [Environment]::NewLine + [Environment]::NewLine + 'Tailnet: ' + $tailnet
                    Show-ToggleOverlay -Title ([string]$state.Title) -Message $message -Indicator ([string]$state.Indicator)
                    $script:PendingSwitchAccountIdentifier = ''
                    $script:PendingSwitchAccountDisplayIdentifier = ''
                    $script:PendingSwitchPreviousTailnet = ''
                    try { Update-LoggedAccountsView -Force $true } catch { }
                    Start-DelayedAccountRefresh -Milliseconds 900
                    Start-DelayedAccountRefresh -Milliseconds 2500
                }
            }
            catch {
                try { $sender.Stop(); $sender.Dispose() } catch { }
                $script:PendingSwitchAccountIdentifier = ''
                $script:PendingSwitchAccountDisplayIdentifier = ''
                $script:PendingSwitchPreviousTailnet = ''
                Show-Overlay -Title 'Switch account refresh failed' -Message $_.Exception.Message -ErrorStyle
            }
        }.GetNewClosure())
        $timer.Start()
    }
    catch {
        $id = ConvertTo-DiagnosticText -Text ([string]$Identifier)
        if ([string]::IsNullOrWhiteSpace($id)) { $id = 'selected account' }
        $tailnet = Get-CurrentTailnetForAccountOverlay
        Show-ToggleOverlay -Title $Title -Message ('Switched to "' + $id + '".' + [Environment]::NewLine + [Environment]::NewLine + 'Tailnet: ' + $tailnet) -Indicator $Indicator
        $script:PendingSwitchAccountIdentifier = ''
        $script:PendingSwitchAccountDisplayIdentifier = ''
        $script:PendingSwitchPreviousTailnet = ''
    }
}

function Invoke-TailscaleAccountSwitch {
    param($Account,$Button = $null)
    if ($null -eq $Account) { return }
    if ([bool](Get-ObjectPropertyOrDefault $Account 'Active' $false)) {
        $activeAccountLabel = [string](Get-ObjectPropertyOrDefault $Account 'Identifier' '-')
        $activeTailnet = Get-CurrentTailnetForAccountOverlay
        if ([string]::IsNullOrWhiteSpace([string]$activeTailnet)) { $activeTailnet = '-' }
        Show-Overlay -Title 'Account already active' -Message ('Account: ' + $activeAccountLabel + [Environment]::NewLine + [Environment]::NewLine + 'Tailnet: ' + $activeTailnet) -Indicator 'Info'
        return
    }
    $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'SwitchIdentifier' ''))
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Identifier' '')) }
    if ([string]::IsNullOrWhiteSpace($identifier)) { throw 'The selected account does not expose a switch identifier.' }
    $displayIdentifier = Get-TailscaleAccountSwitchOutputIdentifier -Account $Account
    $script:PendingSwitchAccountIdentifier = [string]$identifier
    $script:PendingSwitchAccountDisplayIdentifier = [string]$displayIdentifier
    $script:PendingSwitchPreviousTailnet = Get-CurrentTailnetForAccountOverlay
    try {
        if ($null -ne $script:gridAccounts -and $null -ne $script:gridAccounts.CurrentRow -and $script:gridAccounts.CurrentRow.Tag -eq $Account) {
            $script:gridAccounts.CurrentRow.Cells['AccountStatus'].Value = 'Connecting...'
            Set-AccountRowVisualState -Row $script:gridAccounts.CurrentRow -State 'Connecting...'
        }
    } catch { }
    try {
        Start-TailscaleActionProcessAsync -Title 'Switch Account' -Arguments @('switch',$identifier) -SuccessTitle 'Tailscale account switched' -SuccessMessage '' -Indicator 'Info' -Button $Button -BusyText 'Switching...' -Kind 'SwitchAccount'
        Start-DelayedAccountRefresh -Milliseconds 1800
        Start-DelayedAccountRefresh -Milliseconds 4200
    }
    catch {
        $script:PendingSwitchAccountIdentifier = ''
        $script:PendingSwitchAccountDisplayIdentifier = ''
        $script:PendingSwitchPreviousTailnet = ''
        throw
    }
}

function Add-TailscaleAccount {
    param($Button = $null)
    $exe = ''
    try { if ($null -ne $script:Snapshot -and -not [string]::IsNullOrWhiteSpace([string]$script:Snapshot.Exe)) { $exe = [string]$script:Snapshot.Exe } } catch { }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { $exe = Find-TailscaleExe }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { throw 'tailscale.exe was not detected.' }
    try {
        if ($null -ne $Button) {
            try { $Button.Enabled = $false; $Button.Text = 'Authenticating...' } catch { }
        }
        $powerShellExe = [string]$script:PowerShellExe
        if ([string]::IsNullOrWhiteSpace([string]$powerShellExe)) {
            try { $powerShellExe = [string](Get-Command powershell.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1) } catch { }
        }
        if ([string]::IsNullOrWhiteSpace([string]$powerShellExe)) { $powerShellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe' }
        Remove-StaleTailscaleLoginScripts -OlderThanMinutes 1
        $loginScriptPath = Join-Path $script:AppRoot ('tailscale-login-' + [guid]::NewGuid().ToString('N') + '.ps1')
        $normalizedExe = ([string]$exe).Trim()
        if ($normalizedExe.StartsWith('& ')) { $normalizedExe = $normalizedExe.Substring(2).Trim() }
        if ($normalizedExe.Length -ge 2 -and $normalizedExe.StartsWith('"') -and $normalizedExe.EndsWith('"')) { $normalizedExe = $normalizedExe.Substring(1, $normalizedExe.Length - 2) }
        $exeLiteral = "'" + ($normalizedExe -replace "'", "''") + "'"
        $loginScript = @(
            '$ErrorActionPreference = ''Continue''',
            '$tailscaleExe = ' + $exeLiteral,
            'Write-Host ''Starting Tailscale authentication...''',
            'Write-Host ''''',
            'try {',
            '    & $tailscaleExe login',
            '    $rc = [int]$LASTEXITCODE',
            '}',
            'catch {',
            '    Write-Host ''''',
            '    Write-Host (''Tailscale authentication failed: '' + $_.Exception.Message)',
            '    $rc = 1',
            '}',
            'if ($rc -ne 0) {',
            '    Write-Host ''''',
            '    Write-Host (''Tailscale authentication failed with exit code '' + $rc + ''.'')',
            '    Write-Host ''''',
            '    Write-Host ''Press any key to continue . . .'' -NoNewline',
            '    try { [void]$Host.UI.RawUI.ReadKey(''NoEcho,IncludeKeyDown'') } catch { try { [void][Console]::ReadKey($true) } catch { Start-Sleep -Seconds 5 } }',
            '}',
            'try { Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue } catch { }',
            'exit $rc'
        ) -join "`r`n"
        Set-Content -LiteralPath $loginScriptPath -Value $loginScript -Encoding ASCII -Force
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $powerShellExe
        $psi.Arguments = ConvertTo-ProcessArgumentString -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',$loginScriptPath)
        $psi.UseShellExecute = $true
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
        $process = [System.Diagnostics.Process]::Start($psi)
        Write-ActivityCommandBlock -Title 'Add Another Account' -CommandText 'tailscale login' -ExitCode 0 -Output 'Opened tailscale login in a visible terminal. The terminal will close automatically after authentication succeeds.'
        Set-AccountNoConnectedState -Message 'Authentication is running. The account list will refresh when the terminal closes.'
        try { Reset-SlowSnapshotCache } catch { }
        $state = [pscustomobject]@{ Process = $process; Button = $Button; Attempts = 0; ScriptPath = $loginScriptPath }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $timer.add_Tick({
            param($sender,$eventArgs)
            try {
                $state.Attempts = [int]$state.Attempts + 1
                $done = $false
                try { $done = ($null -eq $state.Process -or $state.Process.HasExited) } catch { $done = $true }
                if (-not $done -and [int]$state.Attempts -lt 720) { return }
                try { $sender.Stop(); $sender.Dispose() } catch { }
                try { if ($null -ne $state.Process) { $state.Process.Dispose() } } catch { }
                try {
                    if (-not [string]::IsNullOrWhiteSpace([string]$state.ScriptPath) -and (Test-Path -LiteralPath ([string]$state.ScriptPath))) {
                        Remove-Item -LiteralPath ([string]$state.ScriptPath) -Force -ErrorAction SilentlyContinue
                    }
                } catch { }
                try { Reset-SlowSnapshotCache } catch { }
                try { Update-Status -RefreshAccounts } catch { Write-LogException -Context 'Refresh after add account' -ErrorRecord $_ }
                try { Update-LoggedAccountsView -Force $true } catch { }
                if ($null -ne $state.Button) {
                    try { $state.Button.Text = 'Add Another Account'; $state.Button.Enabled = $true } catch { }
                }
                Start-DelayedAccountRefresh -Milliseconds 1200
                Start-DelayedAccountRefresh -Milliseconds 3000
            }
            catch {
                try { $sender.Stop(); $sender.Dispose() } catch { }
                if ($null -ne $state.Button) {
                    try { $state.Button.Text = 'Add Another Account'; $state.Button.Enabled = $true } catch { }
                }
                Write-LogException -Context 'Monitor add account authentication' -ErrorRecord $_
            }
        }.GetNewClosure())
        $timer.Start()
        Start-DelayedAccountRefresh -Milliseconds 5000
    }
    catch {
        try {
            if (-not [string]::IsNullOrWhiteSpace([string]$loginScriptPath) -and (Test-Path -LiteralPath ([string]$loginScriptPath))) {
                Remove-Item -LiteralPath ([string]$loginScriptPath) -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        if ($null -ne $Button) {
            try { $Button.Text = 'Add Another Account'; $Button.Enabled = $true } catch { }
        }
        throw
    }
}

function Invoke-TailscaleLogoutCurrentAccount {
    param($Button = $null)
    Start-TailscaleActionProcessAsync -Title 'Logout Account' -Arguments @('logout') -SuccessTitle 'Tailscale account logged out' -SuccessMessage 'The active Tailscale account was logged out. Run Add Another Account or Connect to authenticate again.' -Indicator 'Off' -Button $Button -BusyText 'Logging out...' -Kind 'LogoutAccount'
}

function Set-AllowLanAccessPreferenceFromTray {
    param([bool]$Enabled)
    try {
        $cfg = Get-Config
        $cfg.allow_lan_on_exit = [bool]$Enabled
        Save-Config -Config $cfg
        $previousLoading = $script:IsLoadingConfig
        try {
            $script:IsLoadingConfig = $true
            if ($null -ne $script:chkAllowLan) { $script:chkAllowLan.Checked = [bool]$Enabled }
        }
        finally { $script:IsLoadingConfig = $previousLoading }
        try { if ($null -ne $script:TrayMenuAllowLanAccess) { $script:TrayMenuAllowLanAccess.Checked = [bool]$Enabled } } catch { }
        Update-TrayExitNodeMenuState -Snapshot $script:Snapshot
        $activeExit = $false
        try { $activeExit = -not [string]::IsNullOrWhiteSpace([string]$script:Snapshot.CurrentExitNode) } catch { }
        if ($activeExit) {
            $arg = '--exit-node-allow-lan-access=' + ($(if ([bool]$Enabled) { 'true' } else { 'false' }))
            $script:SuppressNextAsyncToggleOverlay = $true
            Start-TailscaleActionProcessAsync -Title 'Update local network access' -Arguments @('set',$arg) -SuccessTitle 'Local network access updated' -SuccessMessage ($(if ([bool]$Enabled) { 'Local network access is allowed while using an exit node.' } else { 'Local network access is blocked while using an exit node.' })) -Indicator ($(if ([bool]$Enabled) { 'On' } else { 'Off' })) -BusyText 'Updating...' -Kind 'ToggleExitLan'
        }
        else {
            Write-ActivityCommandBlock -Title 'Exit node local network access saved' -CommandText 'settings allow_lan_on_exit' -ExitCode 0 -Output ($(if ([bool]$Enabled) { 'Local network access will be allowed the next time an exit node is enabled.' } else { 'Local network access will be blocked the next time an exit node is enabled.' }))
        }
    }
    catch { Show-Overlay -Title 'Update local network access failed' -Message $_.Exception.Message -ErrorStyle }
}

function Switch-TailscaleAccountFromTray {
    param($Account)
    try {
        Invoke-TailscaleAccountSwitch -Account $Account
    }
    catch { Show-Overlay -Title 'Switch account failed' -Message $_.Exception.Message -ErrorStyle }
}

function Clear-TrayDropDownItemsSafe {
    param($MenuItem)
    try {
        if ($null -eq $MenuItem) { return }
        $items = $MenuItem.DropDownItems
        if ($null -eq $items) { return }
        $oldItems = @()
        foreach ($item in @($items)) { if ($null -ne $item) { $oldItems += $item } }
        $items.Clear()
        foreach ($item in $oldItems) {
            try { $item.Dispose() } catch { }
        }
    }
    catch {
        try { $MenuItem.DropDownItems.Clear() } catch { }
    }
}

function Get-TrayMachineSignaturePart {
    param($Machine)
    try {
        $fields = @(
            'Machine','Owner','Status','DNSName','IPv4','IPv6','OS','Connection','Relay','LastSeen','IsLocal'
        )
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($field in $fields) {
            [void]$parts.Add(([string]$field + '=' + [string](Get-ObjectPropertyOrDefault $Machine $field '')))
        }
        return ($parts -join '|')
    }
    catch { return [string]$Machine }
}

function Get-TrayNetworkDevicesSignature {
    param($Snapshot)
    try {
        if ($null -eq $Snapshot) { return 'snapshot:null' }
        $found = [bool](Get-ObjectPropertyOrDefault $Snapshot 'Found' $false)
        $parts = New-Object System.Collections.Generic.List[string]
        [void]$parts.Add('found=' + [string]$found)
        [void]$parts.Add('user=' + [string](Get-ObjectPropertyOrDefault $Snapshot 'User' ''))
        [void]$parts.Add('email=' + [string](Get-ObjectPropertyOrDefault $Snapshot 'UserEmail' ''))
        $machines = @(Convert-ToObjectArray (Get-ObjectPropertyOrDefault $Snapshot 'Machines' @()))
        foreach ($machine in @($machines | Sort-Object @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'Owner' '') } }, @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'Machine' '') } }, @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'DNSName' '') } })) {
            [void]$parts.Add((Get-TrayMachineSignaturePart -Machine $machine))
        }
        return ($parts -join '||')
    }
    catch { return ([guid]::NewGuid().ToString()) }
}

function Get-TrayAccountsSignature {
    param([object[]]$Accounts)
    try {
        if ($null -eq $Accounts -or $Accounts.Count -le 0) { return 'accounts:none' }
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($account in @($Accounts)) {
            $part = @(
                [string](Get-ObjectPropertyOrDefault $account 'Identifier' ''),
                [string](Get-ObjectPropertyOrDefault $account 'SwitchIdentifier' ''),
                [string](Get-ObjectPropertyOrDefault $account 'User' ''),
                [string](Get-ObjectPropertyOrDefault $account 'Active' $false)
            ) -join '|'
            [void]$parts.Add($part)
        }
        return ($parts -join '||')
    }
    catch { return ([guid]::NewGuid().ToString()) }
}

function Get-TrayExitNodeSignature {
    param($Snapshot)
    try {
        if ($null -eq $Snapshot) { return 'exitnodes:null' }
        $preferred = Get-PreferredExitNodeLabel
        $nodes = @(Convert-ToObjectArray (Get-ObjectPropertyOrDefault $Snapshot 'ExitNodes' @()))
        $parts = New-Object System.Collections.Generic.List[string]
        [void]$parts.Add('found=' + [string](Get-ObjectPropertyOrDefault $Snapshot 'Found' $false))
        [void]$parts.Add('preferred=' + [string]$preferred)
        try { [void]$parts.Add('allowLan=' + [string]([bool](Get-Config).allow_lan_on_exit)) } catch { [void]$parts.Add('allowLan=unknown') }
        foreach ($node in @($nodes | Sort-Object @{ Expression = { [string](Get-PropertyValue $_ @('DNSName','Name','Node')) } })) {
            [void]$parts.Add((@(
                [string](Get-PropertyValue $node @('DNSName')),
                [string](Get-PropertyValue $node @('Name')),
                [string](Get-PropertyValue $node @('IPv4'))
            ) -join '|'))
        }
        return ($parts -join '||')
    }
    catch { return ([guid]::NewGuid().ToString()) }
}

function Get-TrayNetworkValue {
    param($Machine,[string]$Property,[string]$Fallback = '-')
    try {
        $value = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Machine $Property ''))
        if ([string]::IsNullOrWhiteSpace($value)) { return [string]$Fallback }
        return [string]$value
    }
    catch { return [string]$Fallback }
}

function Get-TrayNetworkDeviceLabel {
    param($Machine)
    $name = Get-TrayNetworkValue -Machine $Machine -Property 'Machine'
    if ($name -eq '-') { $name = Get-TrayNetworkValue -Machine $Machine -Property 'DNSName' }
    if ($name -eq '-') { $name = Get-TrayNetworkValue -Machine $Machine -Property 'IPv4' }
    return $name
}

function Get-TrayNetworkConnectionText {
    param($Machine)
    $conn = Get-TrayNetworkValue -Machine $Machine -Property 'Connection'
    $relay = Get-TrayNetworkValue -Machine $Machine -Property 'Relay' -Fallback ''
    if ($conn -eq 'Relay' -and -not [string]::IsNullOrWhiteSpace($relay)) { return ('Relay (' + $relay + ')') }
    return $conn
}

function Add-TrayNetworkInfoItem {
    param($Menu,[string]$Label,[string]$Value)
    try {
        if ($null -eq $Menu) { return }
        $clean = if ([string]::IsNullOrWhiteSpace([string]$Value)) { '-' } else { [string]$Value }
        $item = $Menu.DropDownItems.Add($Label + ': ' + (Get-TrayShortText -Text $clean -Max 76))
        Set-TrayMenuItemFixedSize -Item $item -BaseWidth 520 -BaseHeight 28
        Set-TrayCopyInfoMenuItem -Item $item -Label $Label -Value $clean
        Set-TrayMenuItemFixedSize -Item $item -BaseWidth 520 -BaseHeight 28
        Register-TrayCopyMenuItem -Item $item -Label $Label
        $item.ToolTipText = 'Click to copy ' + $Label
    }
    catch { }
}

function Add-TrayNetworkDeviceMenuItem {
    param($ParentMenu,$Machine)
    try {
        if ($null -eq $ParentMenu -or $null -eq $Machine) { return }
        $deviceLabel = Get-TrayNetworkDeviceLabel -Machine $Machine
        $deviceMenu = New-Object System.Windows.Forms.ToolStripMenuItem
        $deviceMenu.Text = $deviceLabel
        $deviceMenu.ToolTipText = ''
        Set-TrayMenuItemFixedSize -Item $deviceMenu -BaseWidth 320 -BaseHeight 28
        Set-TrayDropDownRenderer -Item $deviceMenu
        try { $deviceMenu.DropDownDirection = [System.Windows.Forms.ToolStripDropDownDirection]::Right } catch { }
        $statusValue = Get-TrayNetworkValue -Machine $Machine -Property 'Status'
        $knownStatus = -not ([string]::IsNullOrWhiteSpace([string]$statusValue) -or [string]$statusValue -eq '-')
        $onlineStatus = ([string]$statusValue -eq 'Online' -or [string]$statusValue -eq 'This device' -or [string]$statusValue -eq 'Connected')
        Set-TrayMenuItemDot -Item $deviceMenu -Known:$knownStatus -On:$onlineStatus
        [void]$ParentMenu.DropDownItems.Add($deviceMenu)
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'Device' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'Machine')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'Owner' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'Owner')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'MagicDNS' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'DNSName')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'IPv4' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'IPv4')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'IPv6' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'IPv6')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'OS' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'OS')
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'Conn' -Value (Get-TrayNetworkConnectionText -Machine $Machine)
        Add-TrayNetworkInfoItem -Menu $deviceMenu -Label 'Last Seen' -Value (Get-TrayNetworkValue -Machine $Machine -Property 'LastSeen')
    }
    catch { }
}

function Add-TrayNetworkOwnerMenu {
    param($ParentMenu,[string]$Owner,[object[]]$Machines)
    try {
        if ($null -eq $ParentMenu) { return }
        $ownerLabel = ConvertTo-DiagnosticText -Text ([string]$Owner)
        if ([string]::IsNullOrWhiteSpace($ownerLabel)) { $ownerLabel = 'Unknown Owner' }
        $ownerMenu = New-Object System.Windows.Forms.ToolStripMenuItem
        $ownerMenu.Text = $ownerLabel
        $ownerMenu.ToolTipText = ''
        Set-TrayMenuItemFixedSize -Item $ownerMenu -BaseWidth 320 -BaseHeight 28
        Set-TrayDropDownRenderer -Item $ownerMenu
        try { $ownerMenu.DropDownDirection = [System.Windows.Forms.ToolStripDropDownDirection]::Right } catch { }
        [void]$ParentMenu.DropDownItems.Add($ownerMenu)
        $list = @($Machines | Sort-Object @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'Machine' '') } })
        if ($list.Count -le 0) {
            $empty = $ownerMenu.DropDownItems.Add('No devices')
            Set-TrayMenuItemFixedSize -Item $empty -BaseWidth 420 -BaseHeight 28
            $empty.Enabled = $false
            return
        }
        foreach ($machine in $list) { Add-TrayNetworkDeviceMenuItem -ParentMenu $ownerMenu -Machine $machine }
    }
    catch { }
}

function Update-TrayNetworkDevicesMenuState {
    param($Snapshot)
    if ($null -eq $script:TrayMenuNetworkDevices) { return }
    try {
        $menu = $script:TrayMenuNetworkDevices
        try { if ($menu.DropDown.Visible) { return } } catch { }
        $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
        $signature = Get-TrayNetworkDevicesSignature -Snapshot $snap
        if ($script:TrayNetworkDevicesSignature -eq $signature -and $menu.DropDownItems.Count -gt 0 -and [string]$menu.DropDownItems[0].Text -ne 'Loading...') { return }
        Clear-TrayDropDownItemsSafe -MenuItem $menu
        $script:TrayNetworkDevicesSignature = $signature
        $menu.Text = 'Network Devices'
        Set-TrayMenuItemFixedSize -Item $menu -BaseWidth 300 -BaseHeight 28
        Set-TrayDropDownRenderer -Item $menu
        Register-TraySubmenuArrowGlyph -Item $menu
        if ($null -eq $snap -or -not [bool](Get-ObjectPropertyOrDefault $snap 'Found' $false)) {
            $menu.Enabled = $true
            $empty = $menu.DropDownItems.Add('Tailscale not detected')
            Set-TrayMenuItemFixedSize -Item $empty -BaseWidth 420 -BaseHeight 28
            $empty.Enabled = $false
            return
        }
        $machines = @(Convert-ToObjectArray (Get-ObjectPropertyOrDefault $snap 'Machines' @()))
        if ($machines.Count -le 0) {
            $menu.Enabled = $true
            $empty = $menu.DropDownItems.Add('No devices detected')
            Set-TrayMenuItemFixedSize -Item $empty -BaseWidth 420 -BaseHeight 28
            $empty.Enabled = $false
            return
        }
        $menu.Enabled = $true
        $selfOwner = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $snap 'User' ''))
        $selfEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $snap 'UserEmail' ''))
        $myDevices = @($machines | Where-Object {
            $owner = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $_ 'Owner' ''))
            [bool](Get-ObjectPropertyOrDefault $_ 'IsLocal' $false) -or
            (-not [string]::IsNullOrWhiteSpace($selfOwner) -and [string]::Equals($owner, $selfOwner, [System.StringComparison]::OrdinalIgnoreCase)) -or
            (-not [string]::IsNullOrWhiteSpace($selfEmail) -and [string]::Equals($owner, $selfEmail, [System.StringComparison]::OrdinalIgnoreCase))
        })
        Add-TrayNetworkOwnerMenu -ParentMenu $menu -Owner 'My Devices' -Machines $myDevices
        $otherMachines = @($machines | Where-Object {
            $candidate = $_
            -not (@($myDevices | Where-Object { $_ -eq $candidate }).Count -gt 0)
        })
        if ($otherMachines.Count -gt 0) { [void]$menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) }
        $owners = @($otherMachines | ForEach-Object { ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $_ 'Owner' 'Unknown Owner')) } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
        foreach ($owner in $owners) {
            $ownerDevices = @($otherMachines | Where-Object { [string]::Equals((ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $_ 'Owner' 'Unknown Owner'))), [string]$owner, [System.StringComparison]::OrdinalIgnoreCase) })
            Add-TrayNetworkOwnerMenu -ParentMenu $menu -Owner $owner -Machines $ownerDevices
        }
        if ($owners.Count -le 0 -and $otherMachines.Count -gt 0) { Add-TrayNetworkOwnerMenu -ParentMenu $menu -Owner 'Other Devices' -Machines $otherMachines }
    }
    catch { Write-LogException -Context 'Update tray network devices menu state' -ErrorRecord $_ }
}

function Update-TraySelectAccountMenuState {
    if ($null -eq $script:TrayMenuSelectAccount) { return }
    try {
        $menu = $script:TrayMenuSelectAccount
        try { if ($menu.DropDown.Visible) { return } } catch { }
        $accounts = @()
        try { $accounts = @(Get-TailscaleSwitchAccounts | Sort-Object -Property @{ Expression = { [string]$_.Identifier } }, @{ Expression = { [string]$_.User } }) } catch { }
        $signature = Get-TrayAccountsSignature -Accounts $accounts
        if ($script:TraySelectAccountSignature -eq $signature -and $menu.DropDownItems.Count -gt 0) { return }
        Clear-TrayDropDownItemsSafe -MenuItem $menu
        $script:TraySelectAccountSignature = $signature
        $menu.Text = 'Switch Tailnet'
        Set-TrayMenuItemFixedSize -Item $menu -BaseWidth 300 -BaseHeight 28
        Register-TraySubmenuArrowGlyph -Item $menu
        if ($accounts.Count -gt 0) {
            $menu.Enabled = $true
            foreach ($account in $accounts) {
                $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'Identifier' ''))
                $user = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'User' ''))
                if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'SwitchIdentifier' '')) }
                if ([string]::IsNullOrWhiteSpace($user)) { $user = '-' }
                if ([string]::IsNullOrWhiteSpace($identifier)) { continue }
                $text = $identifier
                if (-not [string]::IsNullOrWhiteSpace($user) -and $user -ne '-') { $text = $identifier + '  |  ' + $user }
                $item = $menu.DropDownItems.Add($text)
                Set-TrayMenuItemFixedSize -Item $item -BaseWidth 540 -BaseHeight 28
                $item.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                $item.ToolTipText = $text
                $item.Tag = $account
                $item.Checked = [bool](Get-ObjectPropertyOrDefault $account 'Active' $false)
                $item.Enabled = -not [bool](Get-ObjectPropertyOrDefault $account 'Active' $false)
                $acctForClick = $account
                $item.add_Click({ Switch-TailscaleAccountFromTray -Account $acctForClick }.GetNewClosure())
            }
        }
        else {
            $menu.Enabled = $true
            $emptyItem = $menu.DropDownItems.Add('No logged accounts')
            Set-TrayMenuItemFixedSize -Item $emptyItem -BaseWidth 540 -BaseHeight 28
            $emptyItem.Enabled = $false
        }
        [void]$menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $addItem = $menu.DropDownItems.Add('Add Another Account...')
        Set-TrayMenuItemFixedSize -Item $addItem -BaseWidth 540 -BaseHeight 28
        $addItem.add_Click({ try { Add-TailscaleAccount } catch { Show-Overlay -Title 'Add another account failed' -Message $_.Exception.Message -ErrorStyle } })
        [void]$menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $logoutItem = $menu.DropDownItems.Add('Logout')
        Set-TrayMenuItemFixedSize -Item $logoutItem -BaseWidth 540 -BaseHeight 28
        $logoutItem.Enabled = ($accounts.Count -gt 0)
        $logoutItem.add_Click({ try { Invoke-TailscaleLogoutCurrentAccount } catch { Show-Overlay -Title 'Logout failed' -Message $_.Exception.Message -ErrorStyle } })
    }
    catch { Write-LogException -Context 'Update tray account menu state' -ErrorRecord $_ }
}

function Update-TrayExitNodeMenuState {
    param($Snapshot)
    if ($null -eq $script:TrayMenuChooseExitNode) { return }
    try {
        $menu = $script:TrayMenuChooseExitNode
        try { if ($menu.DropDown.Visible) { return } } catch { }
        $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
        $signature = Get-TrayExitNodeSignature -Snapshot $snap
        if ($script:TrayExitNodeSignature -eq $signature -and $menu.DropDownItems.Count -gt 0) { return }
        Clear-TrayDropDownItemsSafe -MenuItem $menu
        $script:TrayExitNodeSignature = $signature
        $menu.Text = 'Preferred Exit Node'
        Set-TrayMenuItemFixedSize -Item $menu -BaseWidth 300 -BaseHeight 28
        Register-TraySubmenuArrowGlyph -Item $menu
        if ($null -eq $snap -or -not [bool](Get-ObjectPropertyOrDefault $snap 'Found' $false)) {
            $menu.Enabled = $false
            [void]$menu.DropDownItems.Add('Tailscale not detected')
            return
        }
        $nodes = @(Convert-ToObjectArray $snap.ExitNodes)
        if ($nodes.Count -eq 0) {
            $menu.Enabled = $false
            [void]$menu.DropDownItems.Add('No exit nodes detected')
            return
        }
        $menu.Enabled = $true
        $preferred = Get-PreferredExitNodeLabel
        foreach ($node in $nodes) {
            if ($null -eq $node) { continue }
            $label = ConvertTo-DnsName ([string](Get-PropertyValue $node @('DNSName','Name','Node')))
            if ([string]::IsNullOrWhiteSpace($label)) { continue }
            $labelForClick = $label
            $item = $menu.DropDownItems.Add($label)
            Set-TrayMenuItemFixedSize -Item $item -BaseWidth 300 -BaseHeight 28
            $item.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $nodeName = ConvertTo-DnsName ([string](Get-PropertyValue $node @('Name')))
            $nodeDns = ConvertTo-DnsName ([string](Get-PropertyValue $node @('DNSName')))
            $nodeIp = ConvertTo-DnsName ([string](Get-PropertyValue $node @('IPv4')))
            $item.Checked = (-not [string]::IsNullOrWhiteSpace($preferred) -and ($preferred -eq $label -or $preferred -eq $nodeName -or $preferred -eq $nodeDns -or $preferred -eq $nodeIp))
            $item.add_Click({ Set-TrayExitNodeSelection -Label $labelForClick }.GetNewClosure())
        }
        [void]$menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $allowLan = $menu.DropDownItems.Add('Allow local network access')
        $script:TrayMenuAllowLanAccess = $allowLan
        Set-TrayMenuItemFixedSize -Item $allowLan -BaseWidth 300 -BaseHeight 28
        $allowLan.CheckOnClick = $false
        $allowLan.Checked = [bool](Get-Config).allow_lan_on_exit
        $allowLan.add_Click({
            try { Set-AllowLanAccessPreferenceFromTray -Enabled:(!$script:TrayMenuAllowLanAccess.Checked) }
            catch { Show-Overlay -Title 'Update local network access failed' -Message $_.Exception.Message -ErrorStyle }
        })
    }
    catch { Write-LogException -Context 'Update tray preferred exit node menu state' -ErrorRecord $_ }
}

function Update-TrayMenuState {
    param($Snapshot)
    try { Update-TrayInfoMenuState -Snapshot $Snapshot } catch { Write-LogException -Context 'Update tray info menu state' -ErrorRecord $_ }
    try { Update-TrayNetworkDevicesMenuState -Snapshot $Snapshot } catch { Write-LogException -Context 'Update tray network devices menu state' -ErrorRecord $_ }
    $items = @($script:TrayMenuToggleConnect,$script:TrayMenuToggleExitNode,$script:TrayMenuToggleSubnets,$script:TrayMenuToggleDns,$script:TrayMenuToggleIncoming)
    if ((@($items | Where-Object { $null -ne $_ })).Count -eq 0) { return }

    $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
    if ($null -eq $snap) {
        Set-TrayMenuItemState -Item $script:TrayMenuToggleConnect -Label 'Connect' -State 'Unknown' -Checked:$false -Enabled:$false
        Set-TrayMenuItemState -Item $script:TrayMenuToggleExitNode -Label 'Exit Node' -State 'Unknown' -Checked:$false -Enabled:$false
        Set-TrayMenuItemState -Item $script:TrayMenuToggleSubnets -Label 'Subnets' -State 'Unknown' -Checked:$false -Enabled:$false
        Set-TrayMenuItemState -Item $script:TrayMenuToggleDns -Label 'Accept DNS' -State 'Unknown' -Checked:$false -Enabled:$false
        Set-TrayMenuItemState -Item $script:TrayMenuToggleIncoming -Label 'Incoming' -State 'Unknown' -Checked:$false -Enabled:$false
        Update-TrayExitNodeMenuState -Snapshot $snap
        return
    }

    $found = [bool]$snap.Found
    $connected = $found -and ([string]$snap.BackendState -eq 'Running')
    $dnsState = Convert-ToNullableBool $snap.CorpDNS
    $subnetState = Convert-ToNullableBool $snap.RouteAll
    $incomingState = Convert-ToNullableBool $snap.IncomingAllowed
    $exitNodeEnabled = -not [string]::IsNullOrWhiteSpace([string]$snap.CurrentExitNode)
    $exitNodeAvailability = Get-ExitNodeToggleAvailability -Snapshot $snap

    Set-TrayMenuItemState -Item $script:TrayMenuToggleConnect -Label 'Connect' -State $(if (-not $found) { 'Not detected' } elseif ($connected) { 'On' } else { 'Off' }) -Checked:$connected -Enabled:$found
    $exitNodeDetail = if ($exitNodeEnabled) { Resolve-ExitNodeDisplay -Snapshot $snap } else { 'None' }
    Set-TrayMenuItemState -Item $script:TrayMenuToggleExitNode -Label 'Exit Node' -State ([string]$exitNodeAvailability.State) -Checked:$exitNodeEnabled -Enabled:([bool]$exitNodeAvailability.Enabled) -Detail $exitNodeDetail
    Set-TrayMenuItemState -Item $script:TrayMenuToggleSubnets -Label 'Subnets' -State $(if (-not $found) { 'Not detected' } elseif ($null -eq $subnetState) { 'Unknown' } elseif ($subnetState) { 'On' } else { 'Off' }) -Checked:([bool]($subnetState -eq $true)) -Enabled:$found
    Set-TrayMenuItemState -Item $script:TrayMenuToggleDns -Label 'Accept DNS' -State $(if (-not $found) { 'Not detected' } elseif ($null -eq $dnsState) { 'Unknown' } elseif ($dnsState) { 'On' } else { 'Off' }) -Checked:([bool]($dnsState -eq $true)) -Enabled:$found
    Set-TrayMenuItemState -Item $script:TrayMenuToggleIncoming -Label 'Incoming' -State $(if (-not $found) { 'Not detected' } elseif ($null -eq $incomingState) { 'Unknown' } elseif ($incomingState) { 'On' } else { 'Off' }) -Checked:([bool]($incomingState -eq $true)) -Enabled:$found
    Update-TrayExitNodeMenuState -Snapshot $snap
}

function Update-TrayText {
    param($Snapshot)
    if ($null -eq $script:NotifyIcon) { return }
    $state = if (-not $Snapshot.Found) { 'Not detected' } elseif ($Snapshot.BackendState -eq 'Running') { 'Connected' } else { 'Disconnected' }
    $device = if ([string]::IsNullOrWhiteSpace([string]$Snapshot.ShortName)) { 'Unknown' } else { [string]$Snapshot.ShortName }
    $text = ('Tailscale Control - {0} - {1}' -f $state, $device)
    if ($text.Length -gt 63) { $text = $text.Substring(0,63) }
    $script:NotifyIcon.Text = $text
}

function Update-StatusBanner {
    param($Snapshot)
    if ($null -eq $script:lblBanner) { return }
    if (-not $Snapshot.Found) {
        $script:lblBanner.Text = 'Tailscale not detected.'
        $script:lblBanner.BackColor = [System.Drawing.Color]::FromArgb(254,226,226)
        $script:lblBanner.ForeColor = [System.Drawing.Color]::FromArgb(153,27,27)
    }
    elseif ($Snapshot.BackendState -eq 'Running') {
        $script:lblBanner.Text = 'Connected and ready for hotkeys.'
        $script:lblBanner.BackColor = $script:Palette.SuccessBack
        $script:lblBanner.ForeColor = $script:Palette.SuccessText
    }
    else {
        $script:lblBanner.Text = 'Detected, but currently disconnected.'
        $script:lblBanner.BackColor = $script:Palette.WarnBack
        $script:lblBanner.ForeColor = $script:Palette.WarnText
    }
}

function Set-UiValue {
    param($Control,[string]$Value)
    if ($null -ne $Control) { $Control.Text = if ([string]::IsNullOrWhiteSpace($Value)) { '-' } else { $Value } }
}

function Get-StateDotColor {
    param([bool]$Known,[bool]$On)
    if (-not $Known) { return [System.Drawing.Color]::FromArgb(107,114,128) }
    if ($On) { return [System.Drawing.Color]::FromArgb(22,163,74) }
    return [System.Drawing.Color]::FromArgb(220,38,38)
}

function New-MainStateDotBitmap {
    param([bool]$Known,[bool]$On)
    $bmp = New-Object System.Drawing.Bitmap 14,14
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $brush = New-Object System.Drawing.SolidBrush (Get-StateDotColor -Known:$Known -On:$On)
        try { $g.FillEllipse($brush, 3, 3, 8, 8) } finally { $brush.Dispose() }
    }
    finally { $g.Dispose() }
    return $bmp
}

function Set-StateValue {
    param($ValueControl,$IndicatorControl,[string]$Value,[bool]$Known = $true,[bool]$On = $false)
    Set-UiValue $ValueControl $Value
    try {
        if ($null -ne $IndicatorControl) {
            if ($null -ne $IndicatorControl.Image) { try { $IndicatorControl.Image.Dispose() } catch { } }
            $IndicatorControl.Text = ''
            $IndicatorControl.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $IndicatorControl.Image = New-MainStateDotBitmap -Known:$Known -On:$On
            $IndicatorControl.ForeColor = Get-StateDotColor -Known:$Known -On:$On
        }
    }
    catch { }
}

function New-StateDotBitmap {
    param([bool]$Known,[bool]$On)
    $bmp = New-Object System.Drawing.Bitmap 16,16
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $brush = New-Object System.Drawing.SolidBrush (Get-StateDotColor -Known:$Known -On:$On)
        try { $g.FillEllipse($brush, 4, 4, 8, 8) } finally { $brush.Dispose() }
    }
    finally { $g.Dispose() }
    return $bmp
}
function Set-TrayMenuItemDot {
    param($Item,[bool]$Known,[bool]$On)
    if ($null -eq $Item) { return }
    try {
        if ($null -ne $Item.Image) { try { $Item.Image.Dispose() } catch { } }
        $Item.Image = New-StateDotBitmap -Known:$Known -On:$On
        $Item.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        try { $Item.ImageScaling = [System.Windows.Forms.ToolStripItemImageScaling]::None } catch { }
        $Item.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
    }
    catch { }
}

function Update-MachineDetailsView {
    param($Machine)
    $detailText = ''
    if ($null -ne $Machine) {
        $pingState = Get-PingStateForMachine -MachineKey (Get-MachineKey -Machine $Machine)
        $role = Get-MachineRoleSummary -Machine $Machine
        $detailLines = New-Object System.Collections.Generic.List[string]
        [void]$detailLines.Add('SELECTED DEVICE')
        [void]$detailLines.Add(('  Device           : {0}' -f [string]$Machine.Machine))
        [void]$detailLines.Add(('  Owner            : {0}' -f [string]$Machine.Owner))
        [void]$detailLines.Add(('  Role             : {0}' -f $role))
        [void]$detailLines.Add(('  Status           : {0}' -f [string]$Machine.Status))
        [void]$detailLines.Add(('  Connection       : {0}' -f [string]$Machine.Connection))
        [void]$detailLines.Add(('  Last Seen        : {0}' -f [string]$Machine.LastSeen))
        [void]$detailLines.Add(('  DERP             : {0}' -f $(if ([string]$Machine.Connection -eq 'Relay' -and -not [string]::IsNullOrWhiteSpace([string]$Machine.Relay)) { ConvertTo-DiagnosticText -Text ([string]$Machine.Relay) } else { '-' })))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('TARGETS')
        [void]$detailLines.Add(('  DNS              : {0}' -f (Get-SelectedMachineTargetValue -Machine $Machine -Kind 'DNS')))
        [void]$detailLines.Add(('  IPv4             : {0}' -f [string]$Machine.IPv4))
        [void]$detailLines.Add(('  IPv6             : {0}' -f [string]$Machine.IPv6))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('PING SUMMARY')
        if ($null -ne $pingState) {
            [void]$detailLines.Add(('  DNS              : {0}' -f [string]$pingState.DNSMs))
            [void]$detailLines.Add(('  IPv4             : {0}' -f [string]$pingState.IPv4Ms))
            [void]$detailLines.Add(('  IPv6             : {0}' -f [string]$pingState.IPv6Ms))
            [void]$detailLines.Add(('  Path             : {0}' -f [string]$pingState.Path))
            [void]$detailLines.Add(('  DERP             : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$pingState.Derp)) { '-' } else { ConvertTo-DiagnosticText -Text ([string]$pingState.Derp) })))
        } else {
            [void]$detailLines.Add('  DNS              : Not run yet')
            [void]$detailLines.Add('  IPv4             : Not run yet')
            [void]$detailLines.Add('  IPv6             : Not run yet')
        }
        [void]$detailLines.Add('')
        [void]$detailLines.Add('NETWORK AND FEATURES')
        [void]$detailLines.Add(('  Exit Node        : {0}' -f [string]$Machine.ExitNodeText))
        [void]$detailLines.Add(('  Exit Option      : {0}' -f [string]$Machine.ExitNodeOptionText))
        [void]$detailLines.Add(('  Tags             : {0}' -f [string]$Machine.Tags))
        [void]$detailLines.Add(('  Allowed IPs      : {0}' -f [string]$Machine.AllowedIPs))
        [void]$detailLines.Add(('  Routes           : {0}' -f [string]$Machine.PrimaryRoutes))
        [void]$detailLines.Add(('  Version          : {0}' -f [string]$Machine.Version))
        [void]$detailLines.Add(('  OS               : {0}' -f [string]$Machine.OS))
        [void]$detailLines.Add(('  Last Handshake   : {0}' -f [string]$Machine.LastHandshake))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('JSON LINK STATE')
        [void]$detailLines.Add(('  Device ID        : {0}' -f [string]$Machine.DeviceId))
        [void]$detailLines.Add(('  Active           : {0}' -f [string]$Machine.ActiveText))
        [void]$detailLines.Add(('  In Network Map   : {0}' -f [string]$Machine.InNetworkMapText))
        [void]$detailLines.Add(('  In MagicSock     : {0}' -f [string]$Machine.InMagicSockText))
        [void]$detailLines.Add(('  In Engine        : {0}' -f [string]$Machine.InEngineText))
        [void]$detailLines.Add(('  Current Addr     : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$Machine.CurAddr)) { '-' } else { [string]$Machine.CurAddr })))
        [void]$detailLines.Add(('  Candidate Addrs  : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$Machine.Addrs)) { '-' } else { [string]$Machine.Addrs })))
        [void]$detailLines.Add(('  Peer API         : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$Machine.PeerAPIURL)) { '-' } else { [string]$Machine.PeerAPIURL })))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('TRAFFIC')
        [void]$detailLines.Add(('  Received         : {0}' -f [string]$Machine.RxText))
        [void]$detailLines.Add(('  Sent             : {0}' -f [string]$Machine.TxText))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('CAPABILITIES')
        [void]$detailLines.Add(('  Summary          : {0}' -f [string]$Machine.CapabilitiesText))
        [void]$detailLines.Add(('  Count            : {0}' -f [string]$Machine.CapabilitiesCount))
        [void]$detailLines.Add(('  Taildrop         : {0}' -f [string]$Machine.TaildropTarget))
        [void]$detailLines.Add(('  File sharing     : {0}' -f $(if ([string]::IsNullOrWhiteSpace([string]$Machine.NoFileSharingReason)) { 'Available or not restricted' } else { [string]$Machine.NoFileSharingReason })))
        [void]$detailLines.Add('')
        [void]$detailLines.Add('TIMESTAMPS')
        [void]$detailLines.Add(('  Created          : {0}' -f [string]$Machine.Created))
        [void]$detailLines.Add(('  Last Write       : {0}' -f [string]$Machine.LastWrite))
        [void]$detailLines.Add(('  Last Handshake   : {0}' -f [string]$Machine.LastHandshake))
        $detailText = ($detailLines -join [Environment]::NewLine)
    }
    if ($null -ne $script:txtMachineDetails) { $script:txtMachineDetails.Text = Limit-UiText -Text $detailText }
    $varMachineJson = Get-Variable -Name txtMachineJson -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $varMachineJson -and $null -ne $varMachineJson.Value) { $varMachineJson.Value.Text = '' }
}

function Update-MachinesView {
    param($Snapshot)
    if ($null -eq $script:gridMachines) { return }
    $filter = ''
    if ($null -ne $script:txtMachineFilter) { $filter = ([string]$script:txtMachineFilter.Text).Trim().ToLowerInvariant() }

    $selectedKey = ''
    if ($null -ne $script:gridMachines.CurrentRow -and $null -ne $script:gridMachines.CurrentRow.Tag) {
        $selectedMachine = $script:gridMachines.CurrentRow.Tag
        $selectedKey = ([string]$selectedMachine.Machine + '|' + [string]$selectedMachine.Owner + '|' + [string]$selectedMachine.IPv4 + '|' + [string]$selectedMachine.IPv6)
    }

    $script:gridMachines.Rows.Clear()
    $selectedRow = $null

    foreach ($m in @($Snapshot.Machines)) {
        $terms = @(
            [string]$m.Machine,
            [string]$m.ShortName,
            [string]$m.Owner,
            [string]$m.IPv4,
            [string]$m.IPv6,
            [string]$m.DNSName,
            [string]$m.LastSeen,
            [string]$m.OS,
            [string]$m.Status,
            [string]$m.Connection,
            [string]$m.Relay
        ) -join ' '
        if (-not [string]::IsNullOrWhiteSpace($filter) -and $terms.ToLowerInvariant().IndexOf($filter) -lt 0) { continue }

        $idx = $script:gridMachines.Rows.Add()
        $row = $script:gridMachines.Rows[$idx]
        $row.Cells['Machine'].Value = [string]$m.Machine
        $row.Cells['Owner'].Value = [string]$m.Owner
        $row.Cells['IPv4'].Value = [string]$m.IPv4
        $row.Cells['IPv6'].Value = [string]$m.IPv6
        $row.Cells['OS'].Value = [string]$m.OS
        $connectionDisplay = [string]$m.Connection
        if ($connectionDisplay -eq 'Relay' -and -not [string]::IsNullOrWhiteSpace([string]$m.Relay)) { $connectionDisplay = 'Relay (' + (ConvertTo-DiagnosticText -Text ([string]$m.Relay)) + ')' }
        $row.Cells['Connection'].Value = $connectionDisplay
        $row.Cells['DNSName'].Value = ConvertTo-DnsName ([string]$m.DNSName)
        $row.Cells['LastSeen'].Value = [string]$m.LastSeen
        $row.Tag = $m
        $connText = [string]$m.Connection
        if ($connText -eq 'Direct') { $row.Cells['Connection'].Style.BackColor = [System.Drawing.Color]::FromArgb(232,246,237) }
        elseif ($connText -eq 'Relay') { $row.Cells['Connection'].Style.BackColor = [System.Drawing.Color]::FromArgb(252,245,220) }
        elseif ($connText -eq 'Offline' -or $connText -eq 'Stopped') { $row.Cells['Connection'].Style.BackColor = [System.Drawing.Color]::FromArgb(250,236,236) }
        $lastSeenText = [string]$m.LastSeen
        if ($lastSeenText -eq 'Connected') { $row.Cells['LastSeen'].Style.BackColor = [System.Drawing.Color]::FromArgb(232,246,237) }
        elseif (-not [string]::IsNullOrWhiteSpace($lastSeenText)) { $row.Cells['LastSeen'].Style.BackColor = [System.Drawing.Color]::FromArgb(250,236,236) }
        if ([string]$m.Status -eq 'This device') { $row.Cells['Machine'].Style.BackColor = [System.Drawing.Color]::FromArgb(236,241,248) }

        $rowKey = ([string]$m.Machine + '|' + [string]$m.Owner + '|' + [string]$m.IPv4 + '|' + [string]$m.IPv6)
        if (-not [string]::IsNullOrWhiteSpace($selectedKey) -and $rowKey -eq $selectedKey) {
            $selectedRow = $row
        }
    }

    if ($script:gridMachines.Rows.Count -gt 0) {
        if ($null -eq $selectedRow) { $selectedRow = $script:gridMachines.Rows[0] }
        $selectedRow.Selected = $true
        $script:gridMachines.CurrentCell = $selectedRow.Cells[0]
        Update-MachineDetailsView -Machine $selectedRow.Tag
        $newSelectedKey = ([string]$selectedRow.Tag.Machine + '|' + [string]$selectedRow.Tag.Owner + '|' + [string]$selectedRow.Tag.IPv4 + '|' + [string]$selectedRow.Tag.IPv6)
        if ([string]::IsNullOrWhiteSpace($selectedKey) -or $newSelectedKey -ne $selectedKey) {
            Update-PingSelection -Machine $selectedRow.Tag
        }
    }
    else {
        Update-MachineDetailsView -Machine $null
        Update-PingSelection -Machine $null
    }

    try {
        if ($null -ne $script:gridMachines -and $script:MachineColumnsInitialized) {
            Set-MachineColumnLayout -PreserveCurrent
        }
    }
    catch {
        Write-Log ('Machine column auto-fit failed: ' + $_.Exception.Message)
    }
}

function Get-VisibleOwnerCount {
    param($Snapshot)
    try {
        $owners = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in @(Get-ObjectPropertyOrDefault $Snapshot 'Machines' @())) {
            $owner = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $m 'Owner' ''))
            if (-not [string]::IsNullOrWhiteSpace($owner)) { [void]$owners.Add($owner.Trim()) }
        }
        return [int]$owners.Count
    }
    catch { return 0 }
}

function ConvertTo-TailscaleAccountDisplayIdentifier {
    param(
        [string]$Details,
        [string]$Fallback
    )
    $value = ConvertTo-DiagnosticText -Text ([string]$Details)
    if ([string]::IsNullOrWhiteSpace($value)) { $value = ConvertTo-DiagnosticText -Text ([string]$Fallback) }
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    $value = [regex]::Replace($value, '(?i)\btailnet(?:name)?\s*:\s*', '')
    $value = [regex]::Replace($value, '(?i)\bdetails\s*:\s*', '')
    $value = [regex]::Replace($value, '\s+\|\s+', ' | ')
    return $value.Trim()
}

function ConvertTo-TailscaleSwitchAccountListFromJson {
    param([string]$Text)
    $accounts = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return @() }
    try {
        $parsed = ConvertFrom-Json -InputObject ([string]$Text) -ErrorAction Stop
        $items = @()
        if ($null -eq $parsed) { return @() }
        elseif ($parsed -is [System.Array]) { $items = @($parsed) }
        elseif ($null -ne (Get-ObjectPropertyOrDefault $parsed 'Accounts' $null)) { $items = @(Get-ObjectPropertyOrDefault $parsed 'Accounts' @()) }
        elseif ($null -ne (Get-ObjectPropertyOrDefault $parsed 'Profiles' $null)) { $items = @(Get-ObjectPropertyOrDefault $parsed 'Profiles' @()) }
        elseif ($null -ne (Get-ObjectPropertyOrDefault $parsed 'Items' $null)) { $items = @(Get-ObjectPropertyOrDefault $parsed 'Items' @()) }
        else { $items = @($parsed) }

        $index = 0
        foreach ($item in $items) {
            if ($null -eq $item) { continue }
            $raw = ConvertTo-DiagnosticText -Text ($item | ConvertTo-Json -Compress -Depth 6)
            $active = $false
            foreach ($prop in @('Active','Current','IsActive','Selected','Connected')) {
                try {
                    $v = Get-ObjectPropertyOrDefault $item $prop $null
                    if ($null -ne $v -and [bool]$v) { $active = $true; break }
                } catch { }
            }
            $identifier = ''
            foreach ($prop in @('Identifier','ID','Id','Name','ProfileName','LoginName','Login','Email','Account','User')) {
                $v = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $item $prop ''))
                if (-not [string]::IsNullOrWhiteSpace($v)) { $identifier = $v; break }
            }
            $user = ''
            foreach ($prop in @('User','LoginName','Login','Email','Account','AccountName','DisplayName')) {
                $v = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $item $prop ''))
                if ($v -match '[A-Z0-9._%+-]+@[A-Z0-9.-]+.[A-Z]{2,}') { $user = [string]$Matches[0]; break }
                if (-not [string]::IsNullOrWhiteSpace($v) -and [string]::IsNullOrWhiteSpace($user)) { $user = $v }
            }
            if ([string]::IsNullOrWhiteSpace($user) -and $identifier -match '[A-Z0-9._%+-]+@[A-Z0-9.-]+.[A-Z]{2,}') { $user = [string]$Matches[0] }
            if ([string]::IsNullOrWhiteSpace($identifier) -and -not [string]::IsNullOrWhiteSpace($user)) { $identifier = $user }
            if ([string]::IsNullOrWhiteSpace($identifier)) { continue }
            if ([string]::IsNullOrWhiteSpace($user)) { $user = $identifier }

            $details = ''
            foreach ($prop in @('Tailnet','TailnetName','ControlURL','Status','ProfileID','ProfileId')) {
                $v = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $item $prop ''))
                if (-not [string]::IsNullOrWhiteSpace($v)) {
                    if ([string]::IsNullOrWhiteSpace($details)) { $details = ($prop + ': ' + $v) } else { $details += ' | ' + $prop + ': ' + $v }
                }
            }
            if ([string]::IsNullOrWhiteSpace($details)) { $details = $raw }
            $switchIdentifier = [string]$identifier
            $displayIdentifier = ConvertTo-TailscaleAccountDisplayIdentifier -Details $details -Fallback $identifier
            if ([string]::IsNullOrWhiteSpace($displayIdentifier)) { $displayIdentifier = $switchIdentifier }
            $index++
            [void]$accounts.Add([pscustomobject]@{
                Index = [int]$index
                Active = [bool]$active
                Status = $(if ([bool]$active) { 'Connected' } else { 'Disconnected' })
                Identifier = [string]$displayIdentifier
                SwitchIdentifier = [string]$switchIdentifier
                User = [string]$user
                Details = [string]$details
                Raw = [string]$raw
            })
        }
    }
    catch { return @() }
    return @($accounts.ToArray())
}

function ConvertTo-TailscaleSwitchAccountList {
    param([string]$Text)
    $accounts = New-Object System.Collections.Generic.List[object]
    $lines = @([string]$Text -split "`r?`n")
    $index = 0
    foreach ($line in $lines) {
        $raw = ConvertTo-DiagnosticText -Text ([string]$line)
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $trimmed = $raw.Trim()
        if ($trimmed -match '(?i)^(logged\s+in|available\s+accounts|accounts?:|profile\s+name|current\s+)') { continue }
        if ($trimmed -match '(?i)^(id\s+|name\s+|account\s+|profile\s+)') { continue }

        $active = $false
        if ($trimmed -match '^([*▶>]+)\s*') {
            $active = $true
            $trimmed = ($trimmed -replace '^([*▶>]+)\s*','').Trim()
        }
        if ($trimmed -match '\s+[*▶>]\s*$') {
            $active = $true
            $trimmed = ($trimmed -replace '\s+[*▶>]\s*$','').Trim()
        }
        if ($trimmed -match '(?i)\b(active|current|connected)\b') { $active = $true }
        $clean = ($trimmed -replace '(?i)\s*\(?\b(active|current|connected)\b\)?\s*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($clean)) { continue }

        $columns = @([regex]::Split($clean, '\t+|\s{2,}') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($columns.Count -eq 0) { $columns = @($clean) }
        $tokens = @($clean -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $emailMatches = @([regex]::Matches($clean, '[A-Z0-9._%+-]+@[A-Z0-9.-]+.[A-Z]{2,}', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))

        $identifier = ''
        if ($emailMatches.Count -gt 0) { $identifier = [string]$emailMatches[0].Value }
        elseif ($columns.Count -gt 0) { $identifier = [string]$columns[0] }
        elseif ($tokens.Count -gt 0) { $identifier = [string]$tokens[0] }
        $identifier = ConvertTo-DiagnosticText -Text $identifier
        if ([string]::IsNullOrWhiteSpace($identifier)) { continue }

        $user = ''
        if ($emailMatches.Count -gt 1) { $user = [string]$emailMatches[$emailMatches.Count - 1].Value }
        elseif ($emailMatches.Count -eq 1) { $user = [string]$emailMatches[0].Value }
        else { $user = $identifier }
        $user = ConvertTo-DiagnosticText -Text $user

        $detailParts = New-Object System.Collections.Generic.List[string]
        foreach ($part in $columns) {
            $p = ConvertTo-DiagnosticText -Text ([string]$part)
            if (-not [string]::IsNullOrWhiteSpace($p) -and $p -ne $identifier -and $p -ne $user) { [void]$detailParts.Add($p) }
        }
        $details = if ($detailParts.Count -gt 0) { $detailParts -join ' | ' } else { $clean }
        $switchIdentifier = [string]$identifier
        $displayIdentifier = ConvertTo-TailscaleAccountDisplayIdentifier -Details $details -Fallback $identifier
        if ([string]::IsNullOrWhiteSpace($displayIdentifier)) { $displayIdentifier = $switchIdentifier }

        $index++
        [void]$accounts.Add([pscustomobject]@{
            Index = [int]$index
            Active = [bool]$active
            Status = $(if ([bool]$active) { 'Connected' } else { 'Disconnected' })
            Identifier = [string]$displayIdentifier
            SwitchIdentifier = [string]$switchIdentifier
            User = [string]$user
            Details = ConvertTo-DiagnosticText -Text ([string]$details)
            Raw = ConvertTo-DiagnosticText -Text ([string]$raw)
        })
    }
    return @($accounts.ToArray())
}

function Get-TailscaleSwitchAccounts {
    param([string]$Exe = '')
    $resolvedExe = Get-TailscaleCommandExe -Exe $Exe
    if ([string]::IsNullOrWhiteSpace([string]$resolvedExe)) { throw 'tailscale.exe was not detected.' }
    try {
        $jsonResult = Invoke-TailscaleCommand -Exe $resolvedExe -Arguments @('switch','--list','--json')
        if ($jsonResult.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$jsonResult.Output)) {
            $jsonAccounts = @(ConvertTo-TailscaleSwitchAccountListFromJson -Text ([string]$jsonResult.Output))
            if ($jsonAccounts.Count -gt 0) { return @($jsonAccounts) }
        }
    }
    catch { Write-Log ('Account JSON list failed, falling back to text: ' + $_.Exception.Message) }
    $result = Invoke-TailscaleCommand -Exe $resolvedExe -Arguments @('switch','--list')
    if ($result.ExitCode -ne 0) {
        $msg = [string]$result.Output
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = 'tailscale switch --list failed.' }
        throw $msg
    }
    return @(ConvertTo-TailscaleSwitchAccountList -Text ([string]$result.Output))
}

function ConvertTo-QuickAccountSwitchTextAccountList {
    param([string]$Text)
    $items = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return @() }
    $emailPattern = '[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}'
    $lines = @([string]$Text -split "`r?`n")
    $index = 0
    foreach ($line in $lines) {
        $raw = ConvertTo-DiagnosticText -Text ([string]$line)
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $trimmed = $raw.Trim()
        if ($trimmed -match '(?i)^(logged\s+in|available\s+accounts|accounts?:|profile\s+name|current\s+)') { continue }
        if ($trimmed -match '(?i)^(id\s+|name\s+|account\s+|profile\s+)') { continue }
        $clean = ($trimmed -replace '^([*▶>]+)\s*','').Trim()
        $clean = ($clean -replace '\s+[*▶>]\s*$','').Trim()
        $clean = ($clean -replace '(?i)\s*\(?\b(active|current|connected)\b\)?\s*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($clean)) { continue }
        $emailMatches = @([regex]::Matches($clean, $emailPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))
        if ($emailMatches.Count -le 0) { continue }
        $tailnetEmail = ConvertTo-DiagnosticText -Text ([string]$emailMatches[0].Value)
        $userEmail = ConvertTo-DiagnosticText -Text ([string]$emailMatches[$emailMatches.Count - 1].Value)
        $switchIdentifier = ''
        $columns = @([regex]::Split($clean, '\t+|\s{2,}|\s+\|\s+') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        foreach ($column in $columns) {
            $candidate = (ConvertTo-DiagnosticText -Text ([string]$column)).Trim().Trim([char[]]@('|',',',';'))
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            if ($candidate -match $emailPattern) { continue }
            if ($candidate -match '(?i)^(active|current|connected)$') { continue }
            $switchIdentifier = $candidate
            break
        }
        if ([string]::IsNullOrWhiteSpace($switchIdentifier)) { $switchIdentifier = $tailnetEmail }
        $index++
        [void]$items.Add([pscustomobject]@{
            Index = [int]$index
            SwitchIdentifier = [string]$switchIdentifier
            TailnetEmail = [string]$tailnetEmail
            UserEmail = [string]$userEmail
            Display = $(if (-not [string]::IsNullOrWhiteSpace($userEmail)) { $tailnetEmail + ' | ' + $userEmail } else { $tailnetEmail })
            Raw = [string]$raw
        })
    }
    return @($items.ToArray())
}

function Get-QuickAccountSwitchTextAccountList {
    try {
        $resolvedExe = Get-TailscaleCommandExe -Exe ''
        if ([string]::IsNullOrWhiteSpace([string]$resolvedExe)) { return @() }
        $result = Invoke-TailscaleCommand -Exe $resolvedExe -Arguments @('switch','--list')
        if ($result.ExitCode -ne 0) { return @() }
        return @(ConvertTo-QuickAccountSwitchTextAccountList -Text ([string]$result.Output))
    }
    catch {
        try { Write-Log ('Quick account text list failed: ' + $_.Exception.Message) } catch { }
        return @()
    }
}

function Add-QuickAccountSwitchDisplayFields {
    param($Accounts,$DisplayRows)
    $accountList = @($Accounts)
    $rows = @($DisplayRows)
    $result = New-Object System.Collections.Generic.List[object]
    $used = @{}
    for ($i = 0; $i -lt $accountList.Count; $i++) {
        $account = $accountList[$i]
        if ($null -eq $account) { continue }
        $switchIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'SwitchIdentifier' ''))
        $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'Identifier' ''))
        $user = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'User' ''))
        $match = $null
        for ($r = 0; $r -lt $rows.Count; $r++) {
            if ($used.ContainsKey([string]$r)) { continue }
            $row = $rows[$r]
            $rowSwitch = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $row 'SwitchIdentifier' ''))
            $rowTailnet = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $row 'TailnetEmail' ''))
            $sameSwitch = (-not [string]::IsNullOrWhiteSpace($switchIdentifier) -and -not [string]::IsNullOrWhiteSpace($rowSwitch) -and [string]::Equals($switchIdentifier, $rowSwitch, [System.StringComparison]::OrdinalIgnoreCase))
            $sameIdentifier = (-not [string]::IsNullOrWhiteSpace($identifier) -and ((-not [string]::IsNullOrWhiteSpace($rowSwitch) -and [string]::Equals($identifier, $rowSwitch, [System.StringComparison]::OrdinalIgnoreCase)) -or (-not [string]::IsNullOrWhiteSpace($rowTailnet) -and [string]::Equals($identifier, $rowTailnet, [System.StringComparison]::OrdinalIgnoreCase))))
            if ($sameSwitch -or $sameIdentifier) {
                $match = $row
                $used[[string]$r] = $true
                break
            }
        }
        if ($null -eq $match -and $i -lt $rows.Count -and -not $used.ContainsKey([string]$i)) {
            $match = $rows[$i]
            $used[[string]$i] = $true
        }
        if ($null -ne $match) {
            try { Add-Member -InputObject $account -MemberType NoteProperty -Name 'QuickTailnetEmail' -Value ([string](Get-ObjectPropertyOrDefault $match 'TailnetEmail' '')) -Force } catch { }
            try { Add-Member -InputObject $account -MemberType NoteProperty -Name 'QuickUserEmail' -Value ([string](Get-ObjectPropertyOrDefault $match 'UserEmail' '')) -Force } catch { }
            try { Add-Member -InputObject $account -MemberType NoteProperty -Name 'QuickDisplay' -Value ([string](Get-ObjectPropertyOrDefault $match 'Display' '')) -Force } catch { }
        }
        [void]$result.Add($account)
    }
    return @($result.ToArray())
}

function Get-QuickAccountSwitchLoggedAccounts {
    param([switch]$Force)
    try {
        if (-not $Force -and $null -ne $script:QuickAccountSwitchAccounts -and @($script:QuickAccountSwitchAccounts).Count -gt 0) {
            return @($script:QuickAccountSwitchAccounts)
        }
        $displayRows = @(Get-QuickAccountSwitchTextAccountList)
        $accounts = @(Get-TailscaleSwitchAccounts | Sort-Object -Property @{ Expression = { [string]$_.Identifier } }, @{ Expression = { [string]$_.User } })
        if ($displayRows.Count -gt 0) { $accounts = @(Add-QuickAccountSwitchDisplayFields -Accounts $accounts -DisplayRows $displayRows) }
        $accounts = @($accounts | Sort-Object -Property @{ Expression = { $d = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $_ 'QuickDisplay' '')); if ([string]::IsNullOrWhiteSpace($d)) { $d = Get-QuickAccountSwitchDisplay -Account $_ -Index 0 }; $d.ToLowerInvariant() } }, @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'SwitchIdentifier' '') } }, @{ Expression = { [string](Get-ObjectPropertyOrDefault $_ 'Identifier' '') } })
        $script:QuickAccountSwitchAccounts = @($accounts)
        return @($accounts)
    }
    catch {
        try { Write-Log ('Quick account switch account load failed: ' + $_.Exception.Message) } catch { }
        if ($null -ne $script:QuickAccountSwitchAccounts) { return @($script:QuickAccountSwitchAccounts) }
        return @()
    }
}

function Sync-QuickAccountSwitchFromLoggedAccounts {
    param($Config,[switch]$Force)
    try {
        $accounts = @(Get-QuickAccountSwitchLoggedAccounts -Force:$Force)
        $script:QuickAccountSwitchAvailable = ([int]$accounts.Count -gt 1)
        $count = [Math]::Max([int]$script:QuickAccountSwitchMinimumRows, [int]$accounts.Count)
        if ($count -gt [int]$script:QuickAccountSwitchMaximumRows) { $count = [int]$script:QuickAccountSwitchMaximumRows }
        Ensure-QuickAccountSwitchDefinitions -Count $count
        if ($null -eq $Config) { $Config = Get-Config }
        Ensure-ConfigHotkeyEntries -Config $Config -Count $count
        if ($null -ne $script:QuickAccountSwitchSectionPanel) {
            Update-QuickAccountSwitchRows -Accounts $accounts
        }
        else {
            try { Save-Config -Config $Config } catch { }
        }
        return @($accounts)
    }
    catch {
        try { Write-LogException -Context 'Sync quick account switch from logged accounts' -ErrorRecord $_ } catch { }
        return @()
    }
}

function Test-AccountMatchesPendingSwitch {
    param($Account)
    try {
        $pending = [string]$script:PendingSwitchAccountIdentifier
        if ([string]::IsNullOrWhiteSpace($pending) -or $null -eq $Account) { return $false }
        foreach ($prop in @('SwitchIdentifier','Identifier','User')) {
            $v = [string](Get-ObjectPropertyOrDefault $Account $prop '')
            if (-not [string]::IsNullOrWhiteSpace($v) -and $v -eq $pending) { return $true }
        }
    } catch { }
    return $false
}

function Set-AccountRowVisualState {
    param($Row,[string]$State)
    if ($null -eq $Row) { return }
    try {
        $normalBack = [System.Drawing.Color]::White
        $normalFore = [System.Drawing.Color]::FromArgb(32,43,54)
        $selectedBack = [System.Drawing.Color]::FromArgb(232,241,255)
        $connectedBack = [System.Drawing.Color]::FromArgb(232,246,237)
        $connectedFore = [System.Drawing.Color]::FromArgb(45,97,65)
        $connectingBack = [System.Drawing.Color]::FromArgb(252,245,220)
        $connectingFore = [System.Drawing.Color]::FromArgb(126,96,35)

        $back = $normalBack
        $fore = $normalFore
        $selBack = $selectedBack
        $selFore = $normalFore
        if ($State -eq 'Connected') {
            $back = $connectedBack
            $fore = $connectedFore
            $selBack = $connectedBack
            $selFore = $connectedFore
        }
        elseif ($State -eq 'Connecting...') {
            $back = $connectingBack
            $fore = $connectingFore
            $selBack = $connectingBack
            $selFore = $connectingFore
        }

        $Row.DefaultCellStyle.BackColor = $back
        $Row.DefaultCellStyle.ForeColor = $fore
        $Row.DefaultCellStyle.SelectionBackColor = $selBack
        $Row.DefaultCellStyle.SelectionForeColor = $selFore
        foreach ($cell in @($Row.Cells)) {
            try {
                $cell.Style.BackColor = $back
                $cell.Style.ForeColor = $fore
                $cell.Style.SelectionBackColor = $selBack
                $cell.Style.SelectionForeColor = $selFore
            } catch { }
        }
        if ($Row.Cells.Contains('AccountStatus')) {
            $Row.Cells['AccountStatus'].Style.Font = $script:SummaryValueFont
        }
    } catch { }
}

function Set-AccountLoggedAccountsValue {
    param([string]$Value)
    try {
        if ($null -eq $script:lblAccountLoggedAccounts) { return }
        $v = ConvertTo-DiagnosticText -Text ([string]$Value)
        if ([string]::IsNullOrWhiteSpace($v)) { $v = '-' }
        if ($null -ne $script:lblAccountLoggedAccounts.Tag -and [string]$script:lblAccountLoggedAccounts.Tag -eq 'InlineLoggedAccounts') {
            Set-UiValue $script:lblAccountLoggedAccounts ('Logged Accounts: ' + $v)
        }
        else {
            Set-UiValue $script:lblAccountLoggedAccounts $v
        }
    }
    catch { }
}

function Update-LoggedAccountsView {
    param([bool]$Force = $false)
    if ($null -eq $script:gridAccounts) { return }
    try {
        $selectedTab = $null
        try { if ($null -ne $script:leftTabs) { $selectedTab = $script:leftTabs.SelectedTab } } catch { }
        if (-not $Force -and $selectedTab -ne $script:tabAccount -and $script:gridAccounts.Rows.Count -gt 0) { return }

        $accounts = @(Get-QuickAccountSwitchLoggedAccounts -Force)
        try {
            $activePending = @($accounts | Where-Object { [bool]$_.Active -and (Test-AccountMatchesPendingSwitch -Account $_) } | Select-Object -First 1)
            if ($activePending.Count -gt 0) {
                $script:PendingSwitchAccountIdentifier = ''
                $script:PendingSwitchAccountDisplayIdentifier = ''
                $script:PendingSwitchPreviousTailnet = ''
            }
        } catch { }
        $script:gridAccounts.Rows.Clear()
        foreach ($account in $accounts) {
            $idx = $script:gridAccounts.Rows.Add()
            $row = $script:gridAccounts.Rows[$idx]
            $status = [string]$account.Status
            if (Test-AccountMatchesPendingSwitch -Account $account) { $status = 'Connecting...' }
            $row.Cells['AccountStatus'].Value = $status
            $row.Cells['AccountIdentifier'].Value = [string]$account.Identifier
            $row.Cells['AccountUser'].Value = [string]$account.User
            $row.Tag = $account
            Set-AccountRowVisualState -Row $row -State $status
        }
        Set-AccountLoggedAccountsValue ([string]$accounts.Count)
        try { Update-QuickAccountSwitchRows -Accounts $accounts } catch { Write-LogException -Context 'Update quick account switch rows' -ErrorRecord $_ }
        try {
            $activeAccount = @($accounts | Where-Object { [bool]$_.Active } | Select-Object -First 1)
            if ($activeAccount.Count -gt 0) {
                if ($null -ne $script:lblAccountIdentifier) { Set-UiValue $script:lblAccountIdentifier ([string]$activeAccount[0].Identifier) }
            }
            else {
                if ($null -ne $script:lblAccountIdentifier) { Set-UiValue $script:lblAccountIdentifier 'No account connected' }
                if ($null -ne $script:lblAccountEmail) { Set-UiValue $script:lblAccountEmail '-' }
                if ($null -ne $script:lblAccountActiveUser) { Set-UiValue $script:lblAccountActiveUser 'No account connected' }
                if ($null -ne $script:lblAccountTailnet) { Set-UiValue $script:lblAccountTailnet '-' }
                if ($null -ne $script:lblAccountDevice) { Set-UiValue $script:lblAccountDevice '-' }
                if ($null -ne $script:lblAccountDnsName) { Set-UiValue $script:lblAccountDnsName '-' }
                if ($null -ne $script:lblAccountVisibleDevices) { Set-UiValue $script:lblAccountVisibleDevices '0' }
                if ($null -ne $script:lblAccountVisibleUsers) { Set-UiValue $script:lblAccountVisibleUsers '0' }
            }
        } catch { }
        if ($null -ne $script:lblAccountListStatus) {
            $activeAccountCount = @($accounts | Where-Object { [bool]$_.Active }).Count
            Set-UiValue $script:lblAccountListStatus ($(if ($accounts.Count -eq 0) { 'No logged accounts were returned by tailscale switch --list.' } elseif ($activeAccountCount -eq 0) { 'No account is currently connected. Double-click an account to switch.' } else { 'Double-click a disconnected account to switch.' }))
        }
    }
    catch {
        if ($null -ne $script:lblAccountListStatus) { Set-UiValue $script:lblAccountListStatus ('Account list failed: ' + $_.Exception.Message) }
        Write-Log ('Account list failed: ' + $_.Exception.Message)
    }
}

function Update-AccountView {
    param($Snapshot,[bool]$RefreshAccounts = $false)
    try {
        if ($null -eq $script:tabAccount) { return }
        $snap = if ($null -ne $Snapshot) { $Snapshot } else { $script:Snapshot }
        if ($null -eq $snap) { return }
        $hasActiveAccountRow = $false
        $hasAccountRows = $false
        try {
            if ($null -ne $script:gridAccounts) {
                $hasAccountRows = ($script:gridAccounts.Rows.Count -gt 0)
                foreach ($row in @($script:gridAccounts.Rows)) {
                    if ($null -ne $row.Tag -and [bool]$row.Tag.Active) { $hasActiveAccountRow = $true; break }
                }
            }
        } catch { }
        $backendState = [string](Get-ObjectPropertyOrDefault $snap 'BackendState' '')
        $tailnetNow = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'Tailnet' ''))
        $emailNow = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $snap 'UserEmail' ''))
        $isConnectedAccount = ([string]$backendState -eq 'Running' -and (-not [string]::IsNullOrWhiteSpace($tailnetNow) -or -not [string]::IsNullOrWhiteSpace($emailNow)))
        if ($hasAccountRows -and -not $hasActiveAccountRow) { $isConnectedAccount = $false }
        if (-not $isConnectedAccount) {
            Set-UiValue $script:lblAccountIdentifier 'No account connected'
            Set-UiValue $script:lblAccountEmail '-'
            Set-UiValue $script:lblAccountActiveUser 'No account connected'
            Set-UiValue $script:lblAccountTailnet '-'
            Set-UiValue $script:lblAccountDevice '-'
            Set-UiValue $script:lblAccountDnsName '-'
            Set-UiValue $script:lblAccountVisibleDevices '0'
            Set-UiValue $script:lblAccountVisibleUsers '0'
            $loggedNoAccount = '-'
            try { if ($null -ne $script:gridAccounts -and $script:gridAccounts.Rows.Count -gt 0) { $loggedNoAccount = [string]$script:gridAccounts.Rows.Count } } catch { }
            Set-AccountLoggedAccountsValue $loggedNoAccount
            if ($RefreshAccounts) { Update-LoggedAccountsView -Force $true }
            return
        }
        Set-UiValue $script:lblAccountActiveUser ([string](Get-ObjectPropertyOrDefault $snap 'User' ''))
        Set-UiValue $script:lblAccountEmail ([string](Get-ObjectPropertyOrDefault $snap 'UserEmail' ''))
        Set-UiValue $script:lblAccountTailnet (ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'Tailnet' '')))
        Set-UiValue $script:lblAccountDevice ([string](Get-ObjectPropertyOrDefault $snap 'ShortName' ''))
        Set-UiValue $script:lblAccountDnsName (ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'DNSName' '')))
        $machines = @(Get-ObjectPropertyOrDefault $snap 'Machines' @())
        Set-UiValue $script:lblAccountVisibleDevices ([string]$machines.Count)
        Set-UiValue $script:lblAccountVisibleUsers ([string](Get-VisibleOwnerCount -Snapshot $snap))
        $accountIdentifier = ConvertTo-DnsName ([string](Get-ObjectPropertyOrDefault $snap 'Tailnet' ''))
        if ([string]::IsNullOrWhiteSpace($accountIdentifier)) { $accountIdentifier = [string](Get-ObjectPropertyOrDefault $snap 'UserEmail' '') }
        try {
            if ($null -ne $script:gridAccounts) {
                foreach ($row in @($script:gridAccounts.Rows)) {
                    if ($null -ne $row.Tag -and [bool]$row.Tag.Active -and -not [string]::IsNullOrWhiteSpace([string]$row.Tag.Identifier)) {
                        $accountIdentifier = [string]$row.Tag.Identifier
                        break
                    }
                }
            }
        } catch { }
        if ($null -ne $script:lblAccountIdentifier) { Set-UiValue $script:lblAccountIdentifier $accountIdentifier }
        $logged = '-'
        try { if ($null -ne $script:gridAccounts -and $script:gridAccounts.Rows.Count -gt 0) { $logged = [string]$script:gridAccounts.Rows.Count } } catch { }
        Set-AccountLoggedAccountsValue $logged
        if ($RefreshAccounts) { Update-LoggedAccountsView -Force $true }
    }
    catch { Write-Log ('Update account view failed: ' + $_.Exception.Message) }
}

function Switch-SelectedTailscaleAccount {
    param($Button = $null)
    if ($null -eq $script:gridAccounts -or $null -eq $script:gridAccounts.CurrentRow -or $null -eq $script:gridAccounts.CurrentRow.Tag) { return }
    Invoke-TailscaleAccountSwitch -Account $script:gridAccounts.CurrentRow.Tag -Button $Button
}

function Get-MachineKey {
    param($Machine)
    if ($null -eq $Machine) { return '' }
    return (
        [string](Get-PropertyValue $Machine @('Machine')) + '|' +
        [string](Get-PropertyValue $Machine @('Owner')) + '|' +
        [string](Get-PropertyValue $Machine @('IPv4')) + '|' +
        [string](Get-PropertyValue $Machine @('IPv6'))
    )
}

function Get-SelectedMachine {
    if ($null -ne $script:gridMachines -and $null -ne $script:gridMachines.CurrentRow -and $null -ne $script:gridMachines.CurrentRow.Tag) {
        return $script:gridMachines.CurrentRow.Tag
    }
    return $null
}

function Update-DiagnosticsSelectionSummary {
    param($Machine = $null)
    if ($null -eq $script:lblDiagSelection) { return }
    if ($null -eq $Machine) { $Machine = Get-SelectedMachine }
    if ($null -eq $Machine) {
        $script:lblDiagSelection.Text = 'No device selected'
        return
    }

    $status = if ([bool](Get-PropertyValue $Machine @('IsLocal'))) { 'Local device' } elseif ([string](Get-PropertyValue $Machine @('Status')) -eq 'Offline') { 'Offline' } else { 'Online' }
    $conn = [string](Get-PropertyValue $Machine @('Connection'))
    if ([string]::IsNullOrWhiteSpace($conn)) { $conn = '-' }
    $dnsTarget = Get-SelectedMachineTargetValue -Machine $Machine -Kind 'DNS'
    $ipv4Target = Get-SelectedMachineTargetValue -Machine $Machine -Kind 'IPv4'
    if ([string]::IsNullOrWhiteSpace($dnsTarget)) { $dnsTarget = '-' }
    if ([string]::IsNullOrWhiteSpace($ipv4Target)) { $ipv4Target = '-' }
    $script:lblDiagSelection.Text = ('{0}  |  {1}  |  Conn: {2}  |  DNS: {3}  |  IPv4: {4}' -f [string](Get-PropertyValue $Machine @('Machine')), $status, $conn, $dnsTarget, $ipv4Target)
}

function Update-SelectedDeviceActionButtons {
    $machine = Get-SelectedMachine
    $allowPing = $false
    $allowWhois = $false
    if ($null -ne $machine -and -not $script:IsPingDiagnosticsTaskRunning -and -not $script:IsDiagnosticsCommandTaskRunning) {
        $allowWhois = $true
        if (-not [bool](Get-PropertyValue $machine @('IsLocal')) -and [string](Get-PropertyValue $machine @('Status')) -ne 'Offline' -and -not $script:IsPingDiagnosticsTaskRunning -and -not $script:IsDiagnosticsCommandTaskRunning) { $allowPing = $true }
    }

    foreach ($btn in @($script:btnCmdPingAll,$script:btnCmdPingDns,$script:btnCmdPingIPv4,$script:btnCmdPingIPv6)) {
        if ($null -ne $btn) {
            $btn.Tag = [pscustomobject]@{ AllowPing = $allowPing }
            $btn.FlatStyle = 'Standard'
            $btn.UseVisualStyleBackColor = $true
            $btn.ForeColor = [System.Drawing.SystemColors]::ControlText
            $btn.Enabled = $allowPing
        }
    }
    foreach ($btn in @($script:btnCmdWhois)) {
        if ($null -ne $btn) {
            $btn.FlatStyle = 'Standard'
            $btn.UseVisualStyleBackColor = $true
            $btn.ForeColor = [System.Drawing.SystemColors]::ControlText
            $btn.Enabled = $allowWhois
        }
    }
}

function Set-DiagnosticsOutput {
    param([string]$Text,[string]$Mode = 'Command')
    $script:DiagnosticsContentMode = [string]$Mode
    if ($null -eq $script:txtPingDetails) { return }
    $script:txtPingDetails.Text = Limit-UiText -Text ([string]$Text)
    try {
        $script:txtPingDetails.SelectionStart = 0
        $script:txtPingDetails.SelectionLength = 0
        $script:txtPingDetails.ScrollToCaret()
    }
    catch { Write-LogException -Context 'Scroll ping details output' -ErrorRecord $_ }
}

function Set-DnsResolveOutput {
    param([string]$Text)
    if ($null -eq $script:txtDnsResolveOutput) { return }
    $script:txtDnsResolveOutput.Text = Limit-UiText -Text ([string]$Text)
    try {
        $script:txtDnsResolveOutput.SelectionStart = 0
        $script:txtDnsResolveOutput.SelectionLength = 0
        $script:txtDnsResolveOutput.ScrollToCaret()
    }
    catch { Write-LogException -Context 'Set DNS resolve output scroll' -ErrorRecord $_ }
}

function Set-PublicIpOutput {
    param([string]$Text)
    if ($null -eq $script:txtPublicIpOutput) { return }
    $script:txtPublicIpOutput.Text = Limit-UiText -Text ([string]$Text)
    try {
        $script:txtPublicIpOutput.SelectionStart = 0
        $script:txtPublicIpOutput.SelectionLength = 0
        $script:txtPublicIpOutput.ScrollToCaret()
    }
    catch { Write-LogException -Context 'Set Public IP output scroll' -ErrorRecord $_ }
}

function Get-DefaultDnsResolveDomain {
    return 'example.com'
}

function Update-DnsResolveDefaultDomain {
    param([switch]$Force)
    if ($null -eq $script:txtDnsResolveDomain) { return }
    $next = [string](Get-DefaultDnsResolveDomain)
    $current = [string]$script:txtDnsResolveDomain.Text
    if ($Force -or [string]::IsNullOrWhiteSpace($current) -or $current -eq [string]$script:DnsResolveAutoDomain) {
        $script:txtDnsResolveDomain.Text = $next
        $script:DnsResolveAutoDomain = $next
    }
}

function Get-DnsResolveModeText {
    if ($null -eq $script:cmbDnsResolveResolver -or $null -eq $script:cmbDnsResolveResolver.SelectedItem) { return 'Current' }
    $text = ([string]$script:cmbDnsResolveResolver.SelectedItem).Trim()
    if ($text -match '(?i)^system') { return 'System DNS' }
    if ($text -match '(?i)^tailscale') { return 'Tailscale DNS' }
    if ($text -match '(?i)^other') { return 'Other' }
    return 'Current'
}

function Get-DnsResolveCacheMode {
    try {
        if ($null -ne $script:radDnsResolveUseCache -and [bool]$script:radDnsResolveUseCache.Checked) { return 'Allow cache' }
    }
    catch { }
    return 'Bypass cache'
}

function Get-FirstDnsServerFromText {
    param([string]$Text)
    $raw = ([string]$Text).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) { return '' }
    if ($raw -match '(?i)^system\s+default') { return '' }
    $clean = $raw -replace '(?i)^(using|dns used|dns|resolver|server)\s*:\s*',''
    foreach ($match in [regex]::Matches($clean, '\b(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}\b')) {
        $value = ([string]$match.Value).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
    }
    foreach ($match in [regex]::Matches($clean, '(?i)(?:^|[\s,;|\[\(])([0-9a-f]{0,4}:[0-9a-f:%.]{2,})(?:$|[\s,;|\]\)])')) {
        $value = ([string]$match.Groups[1].Value).Trim().Trim('[',']','(',')',',',';','|')
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $ipObj = $null
        if ([System.Net.IPAddress]::TryParse(($value -replace '%.*$',''), [ref]$ipObj)) { return $value }
    }
    foreach ($part in @($clean -split '[,;|]')) {
        $value = ([string]$part).Trim().Trim(',', ';', '|', '[', ']', '(', ')')
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($value -match '(?i)^system\s+default') { continue }
        return $value
    }
    return ''
}

function Test-IsTailscaleVirtualDnsProxyAddress {
    param([string]$Address)
    $value = ([string]$Address).Trim().Trim(',', ';', '|')
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    $value = ($value -replace '%.*$','')
    if ($value -eq '100.100.100.100') { return $true }
    if ($value -match '(?i)^fd7a:115c:a1e0::53$') { return $true }
    return $false
}

function Get-TailscaleDeliveredDnsResolveServers {
    param($Snapshot = $null)
    $items = New-Object System.Collections.Generic.List[string]
    try {
        $snapshot = $Snapshot
        if ($null -eq $snapshot) { $snapshot = Get-CurrentSnapshot }
        $raw = [string](Get-ObjectPropertyOrDefault $snapshot 'DnsRaw' '')
        $formatted = ''
        try { $formatted = [string](Format-DnsOutputText -Raw $raw) } catch { $formatted = '' }
        $combined = @($formatted, $raw) -join [Environment]::NewLine
        $lines = @($combined -split "`r?`n")
        $section = ''
        $subsection = ''

        foreach ($line in $lines) {
            $trim = ([string]$line).Trim()
            if ([string]::IsNullOrWhiteSpace($trim)) { continue }

            if ($trim -match '^={3,}.*MagicDNS configuration.*={3,}$') { $section = 'MagicRaw'; $subsection = ''; continue }
            if ($trim -match '^={3,}.*System DNS configuration.*={3,}$') { $section = 'SystemRaw'; $subsection = ''; continue }

            if ($trim -match '^TAILSCALE RESOLVERS$') { $section = 'TailscaleFormatted'; $subsection = 'Resolvers'; continue }
            if ($trim -match '^Resolvers \(in preference order\):$' -and $section -eq 'MagicRaw') { $subsection = 'Resolvers'; continue }
            if ($trim -match '^Nameservers IP Addresses:$' -and $section -eq 'MagicRaw') { $subsection = 'Resolvers'; continue }

            if ($trim -match '^(SPLIT DNS ROUTES|Split DNS Routes:|SEARCH DOMAINS|Search Domains:|SYSTEM DNS SERVERS|FALLBACK RESOLVERS|Fallback Resolvers:|OVERVIEW|DNS|Certificate Domains:|Additional DNS Records:|Filtered suffixes).*$') {
                if ($section -eq 'TailscaleFormatted') { $section = ''; $subsection = '' }
                elseif ($section -eq 'MagicRaw') { $subsection = '' }
                continue
            }

            if ($section -eq 'MagicRaw') {
                if ($subsection -eq 'Resolvers' -and $trim -match '^\-\s+(.+)$') {
                    $value = Get-FirstDnsServerFromText -Text ([string]$Matches[1])
                    if ([string]::IsNullOrWhiteSpace($value)) { continue }
                    if ($value -match '^\(no\s+') { continue }
                    if ($value -match '->') { continue }
                    if (Test-IsTailscaleVirtualDnsProxyAddress -Address $value) { continue }
                    if (-not $items.Contains($value)) { [void]$items.Add($value) }
                }
                continue
            }

            if ($section -eq 'TailscaleFormatted' -and $subsection -eq 'Resolvers') {
                $value = Get-FirstDnsServerFromText -Text $trim
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                if ($value -match '^\(no\s+') { continue }
                if ($value -match '->') { continue }
                if (Test-IsTailscaleVirtualDnsProxyAddress -Address $value) { continue }
                if (-not $items.Contains($value)) { [void]$items.Add($value) }
                continue
            }
        }
    }
    catch { }
    return @($items)
}

function Get-SystemDnsResolveServers {
    $items = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($family in @('IPv4','IPv6')) {
            try {
                $rows = Get-DnsClientServerAddress -AddressFamily $family -ErrorAction Stop
                foreach ($row in @($rows)) {
                    foreach ($addr in @($row.ServerAddresses)) {
                        $value = ([string]$addr).Trim()
                        if ([string]::IsNullOrWhiteSpace($value)) { continue }
                        if ($value -match '^(::1|127\.0\.0\.1)$') { continue }
                        if ($value -match '(?i)^fec0:0:0:ffff::[1-3]$') { continue }
                        if (Test-IsTailscaleVirtualDnsProxyAddress -Address $value) { continue }
                        if (-not $items.Contains($value)) { [void]$items.Add($value) }
                    }
                }
            }
            catch { }
        }
    }
    catch { }
    return @($items)
}

function Get-DnsResolverChainDisplay {
    param(
        $Snapshot = $null,
        [switch]$PreferSystem
    )
    if ($null -eq $Snapshot) {
        try { $Snapshot = Get-CurrentSnapshot } catch { $Snapshot = $null }
    }
    $systemDns = [string](Get-SystemDnsNameservers)
    if ([string]::IsNullOrWhiteSpace($systemDns)) { $systemDns = 'System DNS' }
    $backendState = ''
    $corpDnsValue = $null
    try {
        if ($null -ne $Snapshot) {
            $backendState = [string](Get-ObjectPropertyOrDefault $Snapshot 'BackendState' '')
            $corpDnsValue = Get-ObjectPropertyOrDefault $Snapshot 'CorpDNS' $null
        }
    }
    catch { }
    if ($PreferSystem -or ($backendState -and $backendState -ne 'Running')) { return $systemDns }
    $dnsKnown = ($null -ne $corpDnsValue)
    $dnsOn = [bool]($dnsKnown -and (Convert-ToNullableBool $corpDnsValue))
    if (-not $dnsOn) { return $systemDns }
    $overrides = @(Get-TailscaleDeliveredDnsResolveServers -Snapshot $Snapshot)
    if (@($overrides).Count -gt 0) { return ('100.100.100.100 -> ' + (@($overrides) -join ', ')) }
    return ('100.100.100.100 -> System DNS (' + $systemDns + ')')
}

function Get-DnsResolveServerInfoForMode {
    param([string]$Mode,$Snapshot = $null)
    $modeText = [string]$Mode
    if ($modeText -match '(?i)^tailscale') {
        $display = Get-DnsResolverChainDisplay -Snapshot $Snapshot
        if ([string]::IsNullOrWhiteSpace($display) -or $display -notmatch '100\.100\.100\.100') { $display = '100.100.100.100 -> System DNS' }
        $overrides = @(Get-TailscaleDeliveredDnsResolveServers -Snapshot $Snapshot)
        if (@($overrides).Count -gt 0) {
            $query = Get-FirstDnsServerFromText -Text ([string]$overrides[0])
            if ([string]::IsNullOrWhiteSpace($query)) { $query = [string]$overrides[0] }
            return [pscustomobject]@{ QueryServer = $query; Display = $display; Mode = 'Tailscale DNS' }
        }
        return [pscustomobject]@{ QueryServer = ''; Display = $display; Mode = 'Tailscale DNS' }
    }
    if ($modeText -match '(?i)^other') {
        $other = ''
        try { $other = ([string]$script:txtDnsResolveOtherServer.Text).Trim() } catch { }
        $query = Get-FirstDnsServerFromText -Text $other
        if ([string]::IsNullOrWhiteSpace($query)) { $query = $other }
        return [pscustomobject]@{ QueryServer = $query; Display = $other; Mode = 'Other' }
    }
    if ($modeText -match '(?i)^system') {
        $servers = @(Get-SystemDnsResolveServers)
        $display = if ($servers.Count -gt 0) { $servers -join ', ' } else { 'Windows default resolver' }
        return [pscustomobject]@{ QueryServer = ''; Display = $display; Mode = 'System DNS' }
    }
    $currentDisplay = ''
    try { $currentDisplay = Get-DnsResolverChainDisplay -Snapshot $Snapshot } catch { }
    if ([string]::IsNullOrWhiteSpace($currentDisplay)) { $currentDisplay = 'System default resolver' }
    return [pscustomobject]@{ QueryServer = ''; Display = $currentDisplay; Mode = 'Current' }
}

function Update-DnsResolveServerPreview {
    param($Snapshot = $null)
    try {
        if ($null -eq $script:lblDnsResolveServerPreview) { return }
        $mode = Get-DnsResolveModeText
        $info = Get-DnsResolveServerInfoForMode -Mode $mode -Snapshot $Snapshot
        $display = [string]$info.Display
        if ([string]::IsNullOrWhiteSpace($display)) { $display = '-' }
        $script:lblDnsResolveServerPreview.Text = 'Using: ' + $display
    }
    catch { }
}

function Update-DnsResolveOtherState {
    try {
        $mode = Get-DnsResolveModeText
        $isOther = [string]$mode -eq 'Other'
        if ($null -ne $script:txtDnsResolveOtherServer) { $script:txtDnsResolveOtherServer.Enabled = $isOther -and -not $script:IsDnsResolveTaskRunning }
        if ($null -ne $script:lblDnsResolveOther) { $script:lblDnsResolveOther.Enabled = $isOther -and -not $script:IsDnsResolveTaskRunning }
        Update-DnsResolveServerPreview
    }
    catch { Write-LogException -Context 'Update DNS resolve other state' -ErrorRecord $_ }
}

function Start-DnsResolveProcess {
    param([string]$Domain,[string]$Mode,[string]$Server,[string]$ServerDisplay,[string]$CacheMode)
    $payload = [pscustomobject]@{ Domain = [string]$Domain; Mode = [string]$Mode; Server = [string]$Server; ServerDisplay = [string]$ServerDisplay; CacheMode = [string]$CacheMode }
    $payloadJson = ConvertTo-Json -InputObject $payload -Compress -Depth 4
    $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payloadJson))
    $childScript = @"
`$ProgressPreference = 'SilentlyContinue'
`$InformationPreference = 'SilentlyContinue'
`$VerbosePreference = 'SilentlyContinue'
`$WarningPreference = 'SilentlyContinue'
`$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new(`$false) } catch { }
`$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
`$payload = `$payloadJson | ConvertFrom-Json
`$domain = ([string]`$payload.Domain).Trim()
`$mode = ([string]`$payload.Mode).Trim()
`$server = ([string]`$payload.Server).Trim()
`$serverDisplay = ([string]`$payload.ServerDisplay).Trim()
`$cacheMode = ([string]`$payload.CacheMode).Trim()
if ([string]::IsNullOrWhiteSpace(`$cacheMode)) { `$cacheMode = 'Bypass cache' }
if ([string]::IsNullOrWhiteSpace(`$serverDisplay)) {
    if ([string]::IsNullOrWhiteSpace(`$server)) { `$serverDisplay = 'System default resolver' } else { `$serverDisplay = `$server }
}
if (-not [string]::IsNullOrWhiteSpace(`$server)) { `$queryTargetDisplay = `$server } else { `$queryTargetDisplay = 'system default resolver' }
`$startedLocal = Get-Date
`$sw = [Diagnostics.Stopwatch]::StartNew()
`$records = New-Object System.Collections.Generic.List[object]
`$errors = New-Object System.Collections.Generic.List[string]
`$queryStats = New-Object System.Collections.Generic.List[object]
`$ipLookupDurationMs = `$null
if (`$cacheMode -match '(?i)^(cache|use|allow)') {
    `$cacheMode = 'Allow cache'
    `$cacheStatus = 'Cache allowed; DNS client cache was not flushed before lookup'
}
else {
    `$cacheMode = 'Bypass cache'
    `$cacheStatus = 'Bypass cache requested; Clear-DnsClientCache attempted before lookup'
    try { Clear-DnsClientCache -ErrorAction Stop | Out-Null } catch { `$cacheStatus = 'Bypass cache requested; Clear-DnsClientCache failed: ' + [string]`$_.Exception.Message }
}
function Add-RecordValue {
    param([string]`$Type,[string]`$Name,[string]`$Value)
    `$clean = ([string]`$Value).Trim()
    if ([string]::IsNullOrWhiteSpace(`$clean)) { return }
    foreach (`$existing in `$records.ToArray()) {
        if ([string]`$existing.Type -eq [string]`$Type -and [string]`$existing.Value -eq `$clean) { return }
    }
    `$records.Add([pscustomobject]@{ Type=`$Type; Name=`$Name; Value=`$clean }) | Out-Null
}
function Add-QueryTiming {
    param([string]`$Type,[double]`$Ms,[string]`$Result)
    `$serverText = if ([string]::IsNullOrWhiteSpace(`$server) -and `$mode -match '(?i)^tailscale') { 'Windows/Tailscale-selected server' } elseif ([string]::IsNullOrWhiteSpace(`$server)) { 'Windows-selected server' } else { `$server }
    `$resultText = ([string]`$Result).Trim()
    if ([string]::IsNullOrWhiteSpace(`$resultText)) { `$resultText = 'completed' }
    `$queryStats.Add([pscustomobject]@{ Type=[string]`$Type; Ms=[double]`$Ms; Server=[string]`$serverText; Result=[string]`$resultText }) | Out-Null
}
function Convert-DnsRowValue {
    param([object]`$Row,[string]`$Type)
    if (`$null -eq `$Row) { return '' }
    if (`$Type -eq 'A' -or `$Type -eq 'AAAA') {
        try { return [string]`$Row.IPAddress } catch { return '' }
    }
    if (`$Type -eq 'CNAME') {
        try { return [string]`$Row.NameHost } catch { return '' }
    }
    if (`$Type -eq 'MX') {
        `$pref = ''
        `$exchange = ''
        try { `$pref = [string]`$Row.Preference } catch { }
        try { `$exchange = [string]`$Row.NameExchange } catch { }
        if (-not [string]::IsNullOrWhiteSpace(`$pref) -and -not [string]::IsNullOrWhiteSpace(`$exchange)) { return ('pref=' + `$pref + ' ' + `$exchange) }
        return `$exchange
    }
    if (`$Type -eq 'NS') {
        try { return [string]`$Row.NameHost } catch { return '' }
    }
    if (`$Type -eq 'TXT') {
        try { if (`$null -ne `$Row.Strings) { return [string]::Join(' ', @(`$Row.Strings)) } } catch { }
        try { return [string]`$Row.Text } catch { return '' }
    }
    return ''
}
function Invoke-DnsResolveExternalCommand {
    param([string]`$FileName,[string[]]`$Arguments,[int]`$TimeoutMs = 9000)
    `$psi = New-Object System.Diagnostics.ProcessStartInfo
    `$psi.FileName = `$FileName
    `$quotedArgs = New-Object System.Collections.Generic.List[string]
    foreach (`$a in @(`$Arguments)) {
        `$argText = [string]`$a
        if (`$argText -match '\s|"') { `$argText = '"' + (`$argText -replace '"','\"') + '"' }
        `$quotedArgs.Add(`$argText) | Out-Null
    }
    `$psi.Arguments = [string]::Join(' ', `$quotedArgs.ToArray())
    `$psi.UseShellExecute = `$false
    `$psi.RedirectStandardOutput = `$true
    `$psi.RedirectStandardError = `$true
    `$psi.CreateNoWindow = `$true
    try { `$psi.StandardOutputEncoding = [Text.UTF8Encoding]::new(`$false) } catch { }
    try { `$psi.StandardErrorEncoding = [Text.UTF8Encoding]::new(`$false) } catch { }
    `$p = New-Object System.Diagnostics.Process
    `$p.StartInfo = `$psi
    try {
        [void]`$p.Start()
        if (-not `$p.WaitForExit([int]`$TimeoutMs)) {
            try { `$p.Kill() } catch { }
            return [pscustomobject]@{ ExitCode=124; Stdout=''; Stderr=('Timed out after ' + [string]`$TimeoutMs + ' ms') }
        }
        `$stdout = `$p.StandardOutput.ReadToEnd()
        `$stderr = `$p.StandardError.ReadToEnd()
        return [pscustomobject]@{ ExitCode=[int]`$p.ExitCode; Stdout=[string]`$stdout; Stderr=[string]`$stderr }
    }
    catch {
        return [pscustomobject]@{ ExitCode=999; Stdout=''; Stderr=[string]`$_.Exception.Message }
    }
    finally { try { `$p.Dispose() } catch { } }
}
function Test-DnsResolveNameMatch {
    param([string]`$Name)
    `$nameText = ([string]`$Name).Trim().TrimEnd('.')
    `$domainText = ([string]`$domain).Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace(`$nameText) -or [string]::IsNullOrWhiteSpace(`$domainText)) { return `$false }
    return `$nameText.Equals(`$domainText, [StringComparison]::OrdinalIgnoreCase)
}
function Convert-NslookupTxtValue {
    param([string]`$Value)
    `$v = ([string]`$Value).Trim()
    if ([string]::IsNullOrWhiteSpace(`$v)) { return '' }
    `$v = (`$v -replace '^\s*"', '')
    `$v = (`$v -replace '"\s*$', '')
    `$v = (`$v -replace '"\s+"', '')
    return `$v.Trim()
}
function Invoke-NslookupTxtFallback {
    `$beforeCount = [int]`$records.Count
    try {
        `$nsArgs = @('-querytype=TXT', `$domain)
        if (-not [string]::IsNullOrWhiteSpace(`$server)) { `$nsArgs += `$server }
        `$result = Invoke-DnsResolveExternalCommand -FileName 'nslookup.exe' -Arguments `$nsArgs -TimeoutMs 9000
        `$outLines = @(([string]`$result.Stdout) -split "`r?`n")
        `$collecting = `$false
        `$currentName = ''
        foreach (`$lineRaw in @(`$outLines)) {
            `$line = ([string]`$lineRaw).Trim()
            if ([string]::IsNullOrWhiteSpace(`$line)) { continue }
            if (`$line -match '(?i)^Server:|^Address:|^Non-authoritative answer') { continue }
            if (`$line -match '(?i)^([^\s]+)\s+text\s*=\s*(.*)$') {
                `$currentName = ([string]`$Matches[1]).Trim()
                `$collecting = Test-DnsResolveNameMatch -Name `$currentName
                `$value = Convert-NslookupTxtValue -Value ([string]`$Matches[2])
                if (`$collecting -and -not [string]::IsNullOrWhiteSpace(`$value)) { Add-RecordValue -Type 'TXT' -Name `$currentName -Value `$value }
                continue
            }
            if (`$collecting -and `$line -match '^".+"$') {
                `$value = Convert-NslookupTxtValue -Value `$line
                if (-not [string]::IsNullOrWhiteSpace(`$value)) { Add-RecordValue -Type 'TXT' -Name `$currentName -Value `$value }
                continue
            }
        }
        return ([int]`$records.Count -gt `$beforeCount)
    }
    catch { return `$false }
}
function Invoke-ResolveDnsNameType {
    param([string]`$Type)
    `$cmdSw = [Diagnostics.Stopwatch]::StartNew()
    `$usedTcpFallback = `$false
    `$resultText = 'completed'
    try {
        `$params = @{ Name = `$domain; Type = `$Type; DnsOnly = `$true; QuickTimeout = `$true; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace(`$server)) { `$params.Server = `$server }
        try {
            `$rows = @(Resolve-DnsName @params)
        }
        catch {
            `$firstMessage = [string]`$_.Exception.Message
            if (`$Type -eq 'TXT') {
                if (`$firstMessage -match '(?i)bad\s+dns\s+packet|forcibly\s+closed|connection.*closed|truncated|response.*too\s+large|packet|could\s+not\s+be\s+parsed') {
                    if (Invoke-NslookupTxtFallback) {
                        `$cmdSw.Stop()
                        Add-QueryTiming -Type `$Type -Ms ([double]`$cmdSw.Elapsed.TotalMilliseconds) -Result 'resolved via nslookup fallback'
                        return
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace(`$server) -and `$firstMessage -match '(?i)truncated|tcp|size|packet') {
                    `$usedTcpFallback = `$true
                    `$params.TcpOnly = `$true
                    try {
                        `$rows = @(Resolve-DnsName @params)
                    }
                    catch {
                        `$tcpMessage = [string]`$_.Exception.Message
                        if (`$tcpMessage -match '(?i)bad\s+dns\s+packet|forcibly\s+closed|connection.*closed|truncated|response.*too\s+large|packet|could\s+not\s+be\s+parsed') {
                            if (Invoke-NslookupTxtFallback) {
                                `$cmdSw.Stop()
                                Add-QueryTiming -Type `$Type -Ms ([double]`$cmdSw.Elapsed.TotalMilliseconds) -Result 'resolved via nslookup fallback'
                                return
                            }
                        }
                        throw
                    }
                }
                else {
                    throw
                }
            }
            else {
                throw
            }
        }
        `$cmdSw.Stop()
        `$addedCount = 0
        foreach (`$row in @(`$rows)) {
            `$nameText = ''
            try { `$nameText = [string]`$row.Name } catch { }
            if ([string]::IsNullOrWhiteSpace(`$nameText)) { `$nameText = `$domain }
            `$valueText = Convert-DnsRowValue -Row `$row -Type `$Type
            if (-not [string]::IsNullOrWhiteSpace(`$valueText)) {
                `$beforeRecordCount = [int]`$records.Count
                Add-RecordValue -Type `$Type -Name `$nameText -Value `$valueText
                if ([int]`$records.Count -gt `$beforeRecordCount) { `$addedCount++ }
            }
        }
        if (`$rows.Count -eq 0 -or `$addedCount -eq 0) {
            `$resultText = 'no records'
        }
        elseif (`$usedTcpFallback) {
            `$resultText = 'resolved via TCP fallback'
        }
        else {
            `$resultText = 'resolved'
        }
        Add-QueryTiming -Type `$Type -Ms ([double]`$cmdSw.Elapsed.TotalMilliseconds) -Result `$resultText
    }
    catch {
        `$cmdSw.Stop()
        `$message = [string]`$_.Exception.Message
        if ([string]::IsNullOrWhiteSpace(`$message)) { `$message = 'Resolve-DnsName failed' }
        if (`$Type -eq 'TXT' -and `$message -match '(?i)bad\s+dns\s+packet|forcibly\s+closed|connection.*closed|truncated|response.*too\s+large|packet|could\s+not\s+be\s+parsed') {
            if (Invoke-NslookupTxtFallback) {
                Add-QueryTiming -Type `$Type -Ms ([double]`$cmdSw.Elapsed.TotalMilliseconds) -Result 'resolved via nslookup fallback'
                return
            }
            `$message = 'TXT response could not be parsed by Windows DNS client or the selected DNS server closed the TXT query'
        }
        if (`$message -match '(?i)timed\s*out|timeout') { `$message = 'DNS request timed out' }
        `$errors.Add((`$Type + ': ' + `$message)) | Out-Null
        Add-QueryTiming -Type `$Type -Ms ([double]`$cmdSw.Elapsed.TotalMilliseconds) -Result ('warning: ' + `$message)
    }
}
`$ipSw = [Diagnostics.Stopwatch]::StartNew()
foreach (`$typeName in @('A','AAAA')) { Invoke-ResolveDnsNameType -Type `$typeName }
`$ipSw.Stop()
`$ipLookupDurationMs = [double]`$ipSw.Elapsed.TotalMilliseconds
`$baseTimeouts = @(`$errors.ToArray() | Where-Object { [string]`$_ -match '(?i)^(A|AAAA): .*timed[ -]?out|^(A|AAAA): DNS request timed out' })
if (`$records.Count -eq 0 -and `$baseTimeouts.Count -ge 2) {
    `$errors.Add('Advanced DNS record queries skipped because A and AAAA both timed out.') | Out-Null
}
else {
    foreach (`$typeName in @('CNAME','TXT','MX','NS')) { Invoke-ResolveDnsNameType -Type `$typeName }
}
`$sw.Stop()
if (`$null -eq `$ipLookupDurationMs) { `$ipLookupDurationMs = [double]`$sw.Elapsed.TotalMilliseconds }
`$duration = [Math]::Round([double]`$ipLookupDurationMs, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
`$totalDuration = [Math]::Round([double]`$sw.Elapsed.TotalMilliseconds, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
`$lines = New-Object System.Collections.Generic.List[string]
`$lines.Add('DNS RESOLVE TEST') | Out-Null
`$lines.Add(('Date          : ' + `$startedLocal.ToString('yyyy-MM-dd HH:mm:ss zzz'))) | Out-Null
`$lines.Add(('Domain        : ' + `$domain)) | Out-Null
`$lines.Add(('Resolver type : ' + `$mode)) | Out-Null
if (-not [string]::IsNullOrWhiteSpace(`$server)) {
    `$lines.Add(('DNS used      : ' + `$serverDisplay)) | Out-Null
    `$lines.Add(('Query server  : ' + `$server)) | Out-Null
    if (`$server -eq '100.100.100.100') { `$lines.Add('Effective DNS : Tailscale-selected upstream, exact server not exposed by Resolve-DnsName') | Out-Null }
    else { `$lines.Add(('Effective DNS : ' + `$server)) | Out-Null }
}
else {
    if (`$mode -match '(?i)^tailscale') {
        `$lines.Add(('DNS used      : ' + `$serverDisplay)) | Out-Null
        `$lines.Add('Query server  : system default resolver') | Out-Null
        `$lines.Add(('Effective DNS : ' + `$serverDisplay)) | Out-Null
    }
    else {
        `$lines.Add(('DNS candidates: ' + `$serverDisplay)) | Out-Null
        `$lines.Add('Query server  : system default resolver') | Out-Null
        `$lines.Add('Effective DNS : Windows-selected server, exact server not exposed by Resolve-DnsName') | Out-Null
    }
}
`$lines.Add(('Cache mode    : ' + `$cacheMode)) | Out-Null
`$lines.Add(('Cache         : ' + `$cacheStatus)) | Out-Null
if ([string]::IsNullOrWhiteSpace(`$server) -and `$mode -match '(?i)^tailscale') { `$lines.Add('Lookup tool   : Resolve-DnsName; Windows DNS client using Tailscale DNS path') | Out-Null } elseif ([string]::IsNullOrWhiteSpace(`$server)) { `$lines.Add('Lookup tool   : Resolve-DnsName; Windows DNS client') | Out-Null } elseif (`$server -eq '100.100.100.100') { `$lines.Add('Lookup tool   : Resolve-DnsName; Tailscale Quad100 resolver') | Out-Null } else { `$lines.Add('Lookup tool   : Resolve-DnsName; direct DNS query') | Out-Null }
if (`$records.Count -gt 0 -and `$errors.Count -gt 0) { `$lines.Add('Status        : resolved with warnings') | Out-Null } elseif (`$records.Count -gt 0) { `$lines.Add('Status        : resolved') | Out-Null } else { `$lines.Add('Status        : not resolved') | Out-Null }
`$lines.Add(('IP lookup     : ' + `$duration + ' ms')) | Out-Null
`$lines.Add(('Total test    : ' + `$totalDuration + ' ms')) | Out-Null
`$lines.Add('') | Out-Null
if (`$queryStats.Count -gt 0) {
    `$lines.Add('QUERY TIMINGS') | Out-Null
    foreach (`$stat in @(`$queryStats.ToArray())) {
        `$msText = [Math]::Round([double]`$stat.Ms, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
        `$typePad = ([string]`$stat.Type).PadRight(5)
        `$lines.Add(('  ' + `$typePad + `$msText.PadLeft(8) + ' ms | ' + [string]`$stat.Result + ' | ' + [string]`$stat.Server)) | Out-Null
    }
    `$lines.Add('') | Out-Null
}
foreach (`$typeName in @('A','AAAA','CNAME','TXT','MX','NS')) {
    `$items = @(`$records.ToArray() | Where-Object { [string]`$_.Type -eq `$typeName })
    if (`$items.Count -gt 0) {
        `$lines.Add((`$typeName + ' RECORDS')) | Out-Null
        foreach (`$item in `$items) { `$lines.Add(('  ' + [string]`$item.Value)) | Out-Null }
        `$lines.Add('') | Out-Null
    }
}
if (`$errors.Count -gt 0) {
    `$criticalErrors = @(`$errors.ToArray() | Where-Object { [string]`$_ -match '^(A|AAAA):' })
    `$softErrors = @(`$errors.ToArray() | Where-Object { [string]`$_ -notmatch '^(A|AAAA):' })
    if (`$softErrors.Count -gt 0) {
        `$lines.Add('QUERY WARNINGS') | Out-Null
        foreach (`$err in `$softErrors) { `$lines.Add(('  ' + [string]`$err)) | Out-Null }
        `$lines.Add('') | Out-Null
    }
    if (`$criticalErrors.Count -gt 0 -or `$records.Count -eq 0) {
        `$lines.Add('QUERY ERRORS') | Out-Null
        if (`$criticalErrors.Count -gt 0) { foreach (`$err in `$criticalErrors) { `$lines.Add(('  ' + [string]`$err)) | Out-Null } }
        elseif (`$records.Count -eq 0) { foreach (`$err in `$errors.ToArray()) { `$lines.Add(('  ' + [string]`$err)) | Out-Null } }
        `$lines.Add('') | Out-Null
    }
}
if (`$records.Count -eq 0 -and `$baseTimeouts.Count -ge 2) {
    `$lines.Add('RESOLVER NOTE') | Out-Null
    if ([string]::IsNullOrWhiteSpace(`$server)) {
        `$lines.Add('  The active system resolver did not answer before the timeout.') | Out-Null
    }
    else {
        `$lines.Add(('  The selected DNS server did not answer before the timeout: ' + `$server)) | Out-Null
    }
    `$lines.Add('') | Out-Null
}
[Console]::WriteLine(([string]::Join([Environment]::NewLine, `$lines)).TrimEnd())
if (`$records.Count -gt 0) { exit 0 } else { exit 2 }
"@
    $runnerPath = Join-Path $env:TEMP ('tailscale-control-dnsresolve-' + [guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($runnerPath, $childScript, [Text.Encoding]::Unicode)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:PowerShellExe
    $quotedRunnerPath = '"' + ($runnerPath -replace '"','\"') + '"'
    $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ' + $quotedRunnerPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    try { $psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    try { $psi.StandardErrorEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    return [pscustomobject]@{ Process = $proc; StartedUtc = [datetime]::UtcNow; Domain = [string]$Domain; Mode = [string]$Mode; Server = [string]$Server; ServerDisplay = [string]$ServerDisplay; CacheMode = [string]$CacheMode; ScriptPath = $runnerPath }
}

function Start-PublicIpTestProcess {
    param([string]$Mode = 'Fast')
    $snapshot = Get-CurrentSnapshot
    $currentExitNode = ''
    $currentExitNodeDisplay = ''
    $exitInfo = $null
    $exitNodeName = ''
    $exitNodeDns = ''
    $exitNodeIPv4 = ''
    $exitNodeIPv6 = ''
    $exitNodeId = ''
    $deviceName = ''
    $backendState = ''
    $dnsName = ''
    $ipv4 = ''
    $ipv6 = ''
    try { $currentExitNode = [string](Get-ObjectPropertyOrDefault $snapshot 'CurrentExitNode' '') } catch { }
    try { $exitInfo = Resolve-ExitNodeInfo -Snapshot $snapshot } catch { $exitInfo = $null }
    try { $currentExitNodeDisplay = [string](Resolve-ExitNodeDetailedDisplay -Snapshot $snapshot) } catch { $currentExitNodeDisplay = $currentExitNode }
    if ([string]::IsNullOrWhiteSpace($currentExitNodeDisplay)) { $currentExitNodeDisplay = $currentExitNode }
    try {
        if ($null -ne $exitInfo) {
            $exitNodeName = [string]$exitInfo.Name
            $exitNodeDns = [string]$exitInfo.DNSName
            $exitNodeIPv4 = [string]$exitInfo.IPv4
            $exitNodeIPv6 = [string]$exitInfo.IPv6
            $exitNodeId = [string]$exitInfo.ID
        }
    } catch { }
    try { $deviceName = [string](Get-ObjectPropertyOrDefault $snapshot 'ShortName' '') } catch { }
    try { $dnsName = [string](Get-ObjectPropertyOrDefault $snapshot 'DNSName' '') } catch { }
    try { if ([string]::IsNullOrWhiteSpace($deviceName)) { $deviceName = $dnsName } } catch { }
    try { if ([string]::IsNullOrWhiteSpace($deviceName)) { $deviceName = [Environment]::MachineName } } catch { }
    try { $backendState = [string](Get-ObjectPropertyOrDefault $snapshot 'BackendState' '') } catch { }
    try { $ipv4 = [string](Get-ObjectPropertyOrDefault $snapshot 'TailscaleIPv4' '') } catch { }
    try { $ipv6 = [string](Get-ObjectPropertyOrDefault $snapshot 'TailscaleIPv6' '') } catch { }
    $payload = [pscustomobject]@{ DeviceName = [string]$deviceName; DNSName = [string]$dnsName; TailscaleIPv4 = [string]$ipv4; TailscaleIPv6 = [string]$ipv6; BackendState = [string]$backendState; CurrentExitNode = [string]$currentExitNode; CurrentExitNodeDisplay = [string]$currentExitNodeDisplay; ExitNodeName = [string]$exitNodeName; ExitNodeDNS = [string]$exitNodeDns; ExitNodeIPv4 = [string]$exitNodeIPv4; ExitNodeIPv6 = [string]$exitNodeIPv6; ExitNodeID = [string]$exitNodeId; Mode = [string]$Mode }
    $payloadJson = ConvertTo-Json -InputObject $payload -Compress -Depth 4
    $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payloadJson))
    $childScript = @"
`$ProgressPreference = 'SilentlyContinue'
`$InformationPreference = 'SilentlyContinue'
`$VerbosePreference = 'SilentlyContinue'
`$WarningPreference = 'SilentlyContinue'
`$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new(`$false) } catch { }
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
`$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
`$payload = `$payloadJson | ConvertFrom-Json
`$mode = [string]`$payload.Mode
if ([string]::IsNullOrWhiteSpace(`$mode)) { `$mode = 'Fast' }
`$startedLocal = Get-Date
`$sw = [Diagnostics.Stopwatch]::StartNew()
`$results = New-Object System.Collections.Generic.List[object]
function Add-NoCacheQuery {
    param([string]`$Uri)
    `$sep = '?'
    if (`$Uri.Contains('?')) { `$sep = '&' }
    return `$Uri + `$sep + '_=' + [string][datetime]::UtcNow.Ticks
}
function Invoke-IpEndpoint {
    param([string]`$Name,[string]`$Family,[string]`$Uri)
    `$fullUri = Add-NoCacheQuery -Uri `$Uri
    `$endpointSw = [Diagnostics.Stopwatch]::StartNew()
    `$wc = `$null
    try {
        `$wc = New-Object System.Net.WebClient
        try { `$wc.Encoding = [Text.Encoding]::UTF8 } catch { }
        `$wc.Headers.Add('Cache-Control','no-cache, no-store, max-age=0')
        `$wc.Headers.Add('Pragma','no-cache')
        `$wc.Headers.Add('User-Agent','TailscaleControl-PublicIPTest')
        `$body = ([string]`$wc.DownloadString(`$fullUri)).Trim()
        `$endpointSw.Stop()
        `$ipObj = `$null
        `$valid = [System.Net.IPAddress]::TryParse(`$body, [ref]`$ipObj)
        `$results.Add([pscustomobject]@{ Name=[string]`$Name; Family=[string]`$Family; Uri=[string]`$Uri; Ok=`$true; Value=[string]`$body; ValidIp=[bool]`$valid; Ms=[double]`$endpointSw.Elapsed.TotalMilliseconds; Error='' }) | Out-Null
    }
    catch {
        `$endpointSw.Stop()
        `$results.Add([pscustomobject]@{ Name=[string]`$Name; Family=[string]`$Family; Uri=[string]`$Uri; Ok=`$false; Value=''; ValidIp=`$false; Ms=[double]`$endpointSw.Elapsed.TotalMilliseconds; Error=[string]`$_.Exception.Message }) | Out-Null
    }
    finally { try { if (`$null -ne `$wc) { `$wc.Dispose() } } catch { } }
}
function Get-PublicIpCountryInfo {
    param([string]`$Ip)
    if ([string]::IsNullOrWhiteSpace(`$Ip)) { return [pscustomobject]@{ Country='not detected'; Region=''; City=''; Org=''; Source='-' } }
    `$wc = `$null
    try {
        `$wc = New-Object System.Net.WebClient
        try { `$wc.Encoding = [Text.Encoding]::UTF8 } catch { }
        `$wc.Headers.Add('Cache-Control','no-cache, no-store, max-age=0')
        `$wc.Headers.Add('Pragma','no-cache')
        `$wc.Headers.Add('User-Agent','TailscaleControl-PublicIPGeo')
        `$geoRaw = ([string]`$wc.DownloadString(('https://ipapi.co/' + [Uri]::EscapeDataString(`$Ip) + '/json/'))).Trim()
        `$geo = `$geoRaw | ConvertFrom-Json
        `$country = [string]`$geo.country_name
        `$region = [string]`$geo.region
        `$city = [string]`$geo.city
        `$org = [string]`$geo.org
        if ([string]::IsNullOrWhiteSpace(`$country)) { `$country = [string]`$geo.country }
        return [pscustomobject]@{ Country=`$country; Region=`$region; City=`$city; Org=`$org; Source='ipapi.co' }
    }
    catch {
        return [pscustomobject]@{ Country='not detected'; Region=''; City=''; Org=''; Source=('lookup failed: ' + [string]`$_.Exception.Message) }
    }
    finally { try { if (`$null -ne `$wc) { `$wc.Dispose() } } catch { } }
}
if (`$mode -match '(?i)^detailed') {
    Invoke-IpEndpoint -Name 'api.ipify.org' -Family 'IPv4' -Uri 'https://api.ipify.org/'
    Invoke-IpEndpoint -Name 'ipv4.icanhazip.com' -Family 'IPv4' -Uri 'https://ipv4.icanhazip.com/'
    Invoke-IpEndpoint -Name 'api64.ipify.org' -Family 'Auto' -Uri 'https://api64.ipify.org/'
    Invoke-IpEndpoint -Name 'api6.ipify.org' -Family 'IPv6' -Uri 'https://api6.ipify.org/'
    Invoke-IpEndpoint -Name 'ipv6.icanhazip.com' -Family 'IPv6' -Uri 'https://ipv6.icanhazip.com/'
}
else {
    `$mode = 'Fast'
    Invoke-IpEndpoint -Name 'api.ipify.org' -Family 'IPv4' -Uri 'https://api.ipify.org/'
    Invoke-IpEndpoint -Name 'api64.ipify.org' -Family 'Auto' -Uri 'https://api64.ipify.org/'
}
`$sw.Stop()
`$totalMs = [Math]::Round([double]`$sw.Elapsed.TotalMilliseconds, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
`$okItems = @(`$results.ToArray() | Where-Object { [bool]`$_.Ok })
`$validItems = @(`$okItems | Where-Object { [bool]`$_.ValidIp -and -not [string]::IsNullOrWhiteSpace([string]`$_.Value) })
`$ipv4Values = New-Object System.Collections.Generic.List[string]
`$ipv6Values = New-Object System.Collections.Generic.List[string]
foreach (`$r in @(`$validItems)) {
    `$value = ([string]`$r.Value).Trim()
    if (`$value -match ':') {
        if (-not `$ipv6Values.Contains(`$value)) { `$ipv6Values.Add(`$value) | Out-Null }
    }
    else {
        if (-not `$ipv4Values.Contains(`$value)) { `$ipv4Values.Add(`$value) | Out-Null }
    }
}
`$primaryPublicIp = ''
if (`$ipv4Values.Count -gt 0) { `$primaryPublicIp = [string]`$ipv4Values[0] } elseif (`$ipv6Values.Count -gt 0) { `$primaryPublicIp = [string]`$ipv6Values[0] }
if (`$mode -match '(?i)^detailed') { `$geoInfo = Get-PublicIpCountryInfo -Ip `$primaryPublicIp } else { `$geoInfo = [pscustomobject]@{ Country='not run in Fast mode'; Region=''; City=''; Org=''; Source='Detailed mode only' } }
`$lines = New-Object System.Collections.Generic.List[string]
`$lines.Add('PUBLIC IP TEST') | Out-Null
`$lines.Add(('Date             : ' + `$startedLocal.ToString('yyyy-MM-dd HH:mm:ss zzz'))) | Out-Null
`$lines.Add(('Mode             : ' + `$mode)) | Out-Null
`$lines.Add(('Device           : ' + [string]`$payload.DeviceName)) | Out-Null
if (-not [string]::IsNullOrWhiteSpace([string]`$payload.DNSName)) { `$lines.Add(('MagicDNS         : ' + [string]`$payload.DNSName)) | Out-Null }
if (-not [string]::IsNullOrWhiteSpace([string]`$payload.TailscaleIPv4)) { `$lines.Add(('Tailscale IPv4   : ' + [string]`$payload.TailscaleIPv4)) | Out-Null }
if (-not [string]::IsNullOrWhiteSpace([string]`$payload.TailscaleIPv6)) { `$lines.Add(('Tailscale IPv6   : ' + [string]`$payload.TailscaleIPv6)) | Out-Null }
`$lines.Add(('Backend status   : ' + [string]`$payload.BackendState)) | Out-Null
if ([string]::IsNullOrWhiteSpace([string]`$payload.CurrentExitNode)) {
    `$lines.Add('Active exit node : none') | Out-Null
}
else {
    `$exitNodeLabel = [string]`$payload.ExitNodeName
    if ([string]::IsNullOrWhiteSpace(`$exitNodeLabel)) { `$exitNodeLabel = [string]`$payload.ExitNodeDNS }
    if ([string]::IsNullOrWhiteSpace(`$exitNodeLabel)) { `$exitNodeLabel = [string]`$payload.CurrentExitNode }
    `$lines.Add(('Active exit node : ' + `$exitNodeLabel)) | Out-Null
}
`$lines.Add('Cache            : no-cache headers and unique query parameter') | Out-Null
`$lines.Add(('Duration         : ' + `$totalMs + ' ms')) | Out-Null
if (-not [string]::IsNullOrWhiteSpace([string]`$payload.CurrentExitNode)) {
    `$lines.Add('') | Out-Null
    `$lines.Add('ACTIVE EXIT NODE') | Out-Null
    if (-not [string]::IsNullOrWhiteSpace([string]`$payload.ExitNodeName)) { `$lines.Add(('  Name   : ' + [string]`$payload.ExitNodeName)) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]`$payload.ExitNodeDNS)) { `$lines.Add(('  MagicDNS: ' + [string]`$payload.ExitNodeDNS)) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]`$payload.ExitNodeIPv4)) { `$lines.Add(('  IPv4   : ' + [string]`$payload.ExitNodeIPv4)) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]`$payload.ExitNodeIPv6)) { `$lines.Add(('  IPv6   : ' + [string]`$payload.ExitNodeIPv6)) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]`$payload.ExitNodeID)) { `$lines.Add(('  ID     : ' + [string]`$payload.ExitNodeID)) | Out-Null }
}
`$lines.Add('') | Out-Null
`$lines.Add('OBSERVED PUBLIC IP') | Out-Null
if (`$ipv4Values.Count -gt 0) { `$lines.Add(('  IPv4: ' + ([string]::Join(', ', `$ipv4Values.ToArray())))) | Out-Null } else { `$lines.Add('  IPv4: not detected') | Out-Null }
if (`$ipv6Values.Count -gt 0) { `$lines.Add(('  IPv6: ' + ([string]::Join(', ', `$ipv6Values.ToArray())))) | Out-Null } else { `$lines.Add('  IPv6: not detected') | Out-Null }
`$lines.Add('') | Out-Null
`$lines.Add('GEOLOCATION') | Out-Null
`$lines.Add(('  Country: ' + [string]`$geoInfo.Country)) | Out-Null
if (-not [string]::IsNullOrWhiteSpace([string]`$geoInfo.Region)) { `$lines.Add(('  Region : ' + [string]`$geoInfo.Region)) | Out-Null }
if (-not [string]::IsNullOrWhiteSpace([string]`$geoInfo.City)) { `$lines.Add(('  City   : ' + [string]`$geoInfo.City)) | Out-Null }
if (-not [string]::IsNullOrWhiteSpace([string]`$geoInfo.Org)) { `$lines.Add(('  ISP/Org: ' + [string]`$geoInfo.Org)) | Out-Null }
`$lines.Add(('  Source : ' + [string]`$geoInfo.Source)) | Out-Null
`$lines.Add('') | Out-Null
`$lines.Add('ENDPOINT DETAILS') | Out-Null
foreach (`$r in @(`$results.ToArray())) {
    `$ms = [Math]::Round([double]`$r.Ms, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
    if ([bool]`$r.Ok) {
        if ([bool]`$r.ValidIp) { `$validText = 'valid IP' } else { `$validText = 'not a clean IP response' }
        `$lines.Add(('  ' + [string]`$r.Family + ' | ' + [string]`$r.Name + ' | ' + [string]`$r.Value + ' | ' + `$validText + ' | ' + `$ms + ' ms')) | Out-Null
    }
    else {
        `$lines.Add(('  ' + [string]`$r.Family + ' | ' + [string]`$r.Name + ' | failed | ' + `$ms + ' ms | ' + [string]`$r.Error)) | Out-Null
    }
}
`$ipv4AutoOk = @(`$validItems | Where-Object { [string]`$_.Family -in @('IPv4','Auto') })
`$ipv6Ok = @(`$validItems | Where-Object { [string]`$_.Family -eq 'IPv6' })
`$lines.Add('') | Out-Null
`$lines.Add('  Public IP replies : ' + [string]`$validItems.Count + ' valid replies') | Out-Null
if (`$ipv4AutoOk.Count -gt 0) { `$lines.Add('  IPv4/Auto status  : reachable') | Out-Null } else { `$lines.Add('  IPv4/Auto status  : not detected') | Out-Null }
if (`$ipv6Ok.Count -gt 0) { `$lines.Add('  IPv6 status       : reachable') | Out-Null } else { `$lines.Add('  IPv6 status       : not detected') | Out-Null }
`$lines.Add('') | Out-Null
`$lines.Add('INTERPRETATION') | Out-Null
if ([string]::IsNullOrWhiteSpace([string]`$payload.CurrentExitNode)) {
    `$lines.Add('  No exit node is active. The public IP should be your current network, ISP, or external VPN outside Tailscale.') | Out-Null
}
else {
    `$lines.Add('  Exit node is active. The public IP should belong to the exit node network, not this local network.') | Out-Null
}
if (`$ipv4Values.Count -gt 1) { `$lines.Add('  Warning: multiple different IPv4 public IPs were reported. This can happen with proxying, VPNs, or endpoint issues.') | Out-Null }
if (`$ipv6Values.Count -gt 1) { `$lines.Add('  Warning: multiple different IPv6 public IPs were reported. This can happen with proxying, VPNs, or endpoint issues.') | Out-Null }
[Console]::WriteLine(([string]::Join([Environment]::NewLine, `$lines)).TrimEnd())
if (`$okItems.Count -gt 0) { exit 0 } else { exit 2 }
"@
    $runnerPath = Join-Path $env:TEMP ('tailscale-control-publicip-' + [guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($runnerPath, $childScript, [Text.Encoding]::Unicode)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:PowerShellExe
    $quotedRunnerPath = '"' + ($runnerPath -replace '"','\"') + '"'
    $psi.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ' + $quotedRunnerPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    try { $psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    try { $psi.StandardErrorEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    return [pscustomobject]@{ Process = $proc; StartedUtc = [datetime]::UtcNow; ScriptPath = $runnerPath }
}

function Set-DnsResolveBusyState {
    param([bool]$Busy)
    try {
        if ($null -eq $script:btnDnsResolveRun) { return }
        if ($Busy) {
            $script:btnDnsResolveRun.Enabled = $false
            $script:btnDnsResolveRun.Text = 'Resolving...'
            if ($null -ne $script:txtDnsResolveDomain) { $script:txtDnsResolveDomain.Enabled = $false }
            if ($null -ne $script:cmbDnsResolveResolver) { $script:cmbDnsResolveResolver.Enabled = $false }
            if ($null -ne $script:txtDnsResolveOtherServer) { $script:txtDnsResolveOtherServer.Enabled = $false }
            if ($null -ne $script:radDnsResolveNoCache) { $script:radDnsResolveNoCache.Enabled = $false }
            if ($null -ne $script:radDnsResolveUseCache) { $script:radDnsResolveUseCache.Enabled = $false }
        }
        else {
            $script:btnDnsResolveRun.Text = 'Resolve'
            $script:btnDnsResolveRun.Enabled = $true
            if ($null -ne $script:txtDnsResolveDomain) { $script:txtDnsResolveDomain.Enabled = $true }
            if ($null -ne $script:cmbDnsResolveResolver) { $script:cmbDnsResolveResolver.Enabled = $true }
            if ($null -ne $script:radDnsResolveNoCache) { $script:radDnsResolveNoCache.Enabled = $true }
            if ($null -ne $script:radDnsResolveUseCache) { $script:radDnsResolveUseCache.Enabled = $true }
            Update-DnsResolveOtherState
        }
    }
    catch { Write-LogException -Context 'Set DNS resolve busy state' -ErrorRecord $_ }
}

function Set-PublicIpBusyState {
    param([bool]$Busy)
    try {
        if ($null -eq $script:btnPublicIpRun) { return }
        if ($Busy) {
            $script:btnPublicIpRun.Enabled = $false
            $script:btnPublicIpRun.Text = 'Testing...'
            if ($null -ne $script:radPublicIpFast) { $script:radPublicIpFast.Enabled = $false }
            if ($null -ne $script:radPublicIpDetailed) { $script:radPublicIpDetailed.Enabled = $false }
        }
        else {
            $script:btnPublicIpRun.Text = 'Test public IP'
            $script:btnPublicIpRun.Enabled = $true
            if ($null -ne $script:radPublicIpFast) { $script:radPublicIpFast.Enabled = $true }
            if ($null -ne $script:radPublicIpDetailed) { $script:radPublicIpDetailed.Enabled = $true }
        }
    }
    catch { Write-LogException -Context 'Set Public IP busy state' -ErrorRecord $_ }
}

function Remove-ChildProcessNoise {
    param([string]$Text)
    if ([string]::IsNullOrEmpty([string]$Text)) { return '' }
    $value = [string]$Text
    $value = $value -replace '(?s)#< CLIXML\s*<Objs.*?</Objs>\s*',''
    $lines = New-Object System.Collections.Generic.List[string]
    $insideClixml = $false
    foreach ($line in @($value -split "`r?`n")) {
        $l = [string]$line
        if ($l -match '^#< CLIXML') { $insideClixml = $true; continue }
        if ($insideClixml) {
            if ($l -match '</Objs>') { $insideClixml = $false }
            continue
        }
        if ($l -match 'Preparing modules for first use') { continue }
        [void]$lines.Add($l)
    }
    return (($lines | ForEach-Object { [string]$_ }) -join [Environment]::NewLine).Trim()
}

function Complete-DnsResolveTask {
    $taskRef = $script:DnsResolveTask
    if ($null -eq $taskRef) { return }
    try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
    try {
        $proc = $taskRef.Process
        $stdout = ''
        $stderr = ''
        try { $stdout = [string]$proc.StandardOutput.ReadToEnd() } catch { }
        try { $stderr = [string]$proc.StandardError.ReadToEnd() } catch { }
        $stdout = Remove-ChildProcessNoise -Text $stdout
        $stderr = Remove-ChildProcessNoise -Text $stderr
        $output = (($stdout,$stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
        $exit = 1
        try { $exit = [int]$proc.ExitCode } catch { }
        try { $proc.Dispose() } catch { }
        if ([string]::IsNullOrWhiteSpace($output)) { $output = 'DNS RESOLVE TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Status       : no output returned' }
        $text = ConvertTo-DiagnosticText -Text $output.TrimEnd()
        Set-DnsResolveOutput -Text $text
        $elapsed = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
        $commandText = if ([string]::IsNullOrWhiteSpace([string]$taskRef.Server)) { 'Resolve-DnsName ' + [string]$taskRef.Domain + ' using Windows DNS client' } else { 'Resolve-DnsName ' + [string]$taskRef.Domain + ' -Server ' + [string]$taskRef.Server }
        Write-ActivityCommandBlock -Title 'DNS Resolve Test' -CommandText $commandText -ExitCode $exit -Output $text -DurationMs ([double]$elapsed)
    }
    catch {
        Set-DnsResolveOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'DNS Resolve Test failed')
        Write-ActivityFailureBlock -Title 'DNS Resolve Test failed' -CommandText 'nslookup.exe' -Message $_.Exception.Message
    }
    finally {
        try {
            if ($null -ne $taskRef -and -not [string]::IsNullOrWhiteSpace([string]$taskRef.ScriptPath) -and (Test-Path -LiteralPath ([string]$taskRef.ScriptPath))) {
                Remove-Item -LiteralPath ([string]$taskRef.ScriptPath) -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        $script:IsDnsResolveTaskRunning = $false
        $script:DnsResolveTask = $null
        Set-DnsResolveBusyState -Busy $false
    }
}

function Complete-PublicIpTask {
    $taskRef = $script:PublicIpTask
    if ($null -eq $taskRef) { return }
    try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
    try {
        $proc = $taskRef.Process
        $stdout = ''
        $stderr = ''
        try { $stdout = [string]$proc.StandardOutput.ReadToEnd() } catch { }
        try { $stderr = [string]$proc.StandardError.ReadToEnd() } catch { }
        $stdout = Remove-ChildProcessNoise -Text $stdout
        $stderr = Remove-ChildProcessNoise -Text $stderr
        $output = (($stdout,$stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
        $exit = 1
        try { $exit = [int]$proc.ExitCode } catch { }
        try { $proc.Dispose() } catch { }
        if ([string]::IsNullOrWhiteSpace($output)) { $output = 'PUBLIC IP TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Status           : no output returned' }
        $text = ConvertTo-DiagnosticText -Text $output.TrimEnd()
        Set-PublicIpOutput -Text $text
        $elapsed = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
        Write-ActivityCommandBlock -Title 'Public IP Test' -CommandText 'Public IP endpoints' -ExitCode $exit -Output $text -DurationMs ([double]$elapsed)
    }
    catch {
        Set-PublicIpOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Public IP Test failed')
        Write-ActivityFailureBlock -Title 'Public IP Test failed' -CommandText 'Invoke-WebRequest public IP endpoints' -Message $_.Exception.Message
    }
    finally {
        try {
            if ($null -ne $taskRef -and -not [string]::IsNullOrWhiteSpace([string]$taskRef.ScriptPath) -and (Test-Path -LiteralPath ([string]$taskRef.ScriptPath))) {
                Remove-Item -LiteralPath ([string]$taskRef.ScriptPath) -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        $script:IsPublicIpTaskRunning = $false
        $script:PublicIpTask = $null
        Set-PublicIpBusyState -Busy $false
    }
}

function Start-DnsResolveTestAsync {
    if ($script:IsDnsResolveTaskRunning) { return }
    $domain = ''
    try { $domain = ([string]$script:txtDnsResolveDomain.Text).Trim() } catch { }
    if ([string]::IsNullOrWhiteSpace($domain)) {
        Set-DnsResolveOutput -Text 'DNS RESOLVE TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Domain is empty.'
        return
    }
    if ($domain -match '\s') {
        Set-DnsResolveOutput -Text 'DNS RESOLVE TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Domain cannot contain spaces.'
        return
    }
    $mode = Get-DnsResolveModeText
    $cacheMode = Get-DnsResolveCacheMode
    $dnsSnapshot = $null
    try {
        Reset-SlowSnapshotCache
        $dnsSnapshot = Get-TailscaleSnapshot
    }
    catch {
        try { $dnsSnapshot = Get-CurrentSnapshot } catch { $dnsSnapshot = $null }
    }
    $serverInfo = Get-DnsResolveServerInfoForMode -Mode $mode -Snapshot $dnsSnapshot
    $server = [string]$serverInfo.QueryServer
    $serverDisplay = [string]$serverInfo.Display
    if ($mode -eq 'Other' -and [string]::IsNullOrWhiteSpace([string]$server)) {
        Set-DnsResolveOutput -Text 'DNS RESOLVE TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Other DNS server is empty.'
        return
    }
    try {
        $started = Start-DnsResolveProcess -Domain $domain -Mode $mode -Server $server -ServerDisplay $serverDisplay -CacheMode $cacheMode
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $script:IsDnsResolveTaskRunning = $true
        $script:DnsResolveTask = [pscustomobject]@{
            Process = $started.Process
            Timer = $timer
            StartedUtc = [datetime]$started.StartedUtc
            Domain = [string]$domain
            Mode = [string]$mode
            Server = [string]$server
            ServerDisplay = [string]$serverDisplay
            CacheMode = [string]$cacheMode
        }
        Set-DnsResolveBusyState -Busy $true
        $timer.add_Tick({
            try {
                $taskRef = $script:DnsResolveTask
                if ($null -eq $taskRef) { return }
                if ($null -ne $taskRef.Process -and -not $taskRef.Process.HasExited) {
                    $elapsedMs = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
                    if ($elapsedMs -lt 20000) { return }
                    try { $taskRef.Process.Kill() } catch { }
                }
                Complete-DnsResolveTask
            }
            catch {
                Set-DnsResolveOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'DNS Resolve Test failed')
                $script:IsDnsResolveTaskRunning = $false
                $script:DnsResolveTask = $null
                Set-DnsResolveBusyState -Busy $false
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsDnsResolveTaskRunning = $false
        $script:DnsResolveTask = $null
        Set-DnsResolveBusyState -Busy $false
        Set-DnsResolveOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'DNS Resolve Test failed')
    }
}

function Start-PublicIpTestAsync {
    if ($script:IsPublicIpTaskRunning) { return }
    $mode = 'Fast'
    try { if ($null -ne $script:radPublicIpDetailed -and [bool]$script:radPublicIpDetailed.Checked) { $mode = 'Detailed' } } catch { }
    try {
        $started = Start-PublicIpTestProcess -Mode $mode
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $script:IsPublicIpTaskRunning = $true
        $script:PublicIpTask = [pscustomobject]@{
            Process = $started.Process
            Timer = $timer
            StartedUtc = [datetime]$started.StartedUtc
            Mode = [string]$mode
        }
        Set-PublicIpBusyState -Busy $true
        $timer.add_Tick({
            try {
                $taskRef = $script:PublicIpTask
                if ($null -eq $taskRef) { return }
                if ($null -ne $taskRef.Process -and -not $taskRef.Process.HasExited) { return }
                Complete-PublicIpTask
            }
            catch {
                Set-PublicIpOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Public IP Test failed')
                $script:IsPublicIpTaskRunning = $false
                $script:PublicIpTask = $null
                Set-PublicIpBusyState -Busy $false
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsPublicIpTaskRunning = $false
        $script:PublicIpTask = $null
        Set-PublicIpBusyState -Busy $false
        Set-PublicIpOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Public IP Test failed')
    }
}
function Convert-ToDiagnosticsDouble {
    param([string]$Text)
    $value = 0.0
    [void][double]::TryParse([string]$Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)
    return $value
}

function Convert-ToHumanBytes {
    param([double]$Bytes)
    if ($Bytes -lt 0) { $Bytes = 0 }
    $units = @('B','KB','MB','GB','TB')
    $value = [double]$Bytes
    $index = 0
    while ($value -ge 1024 -and $index -lt ($units.Count - 1)) {
        $value = $value / 1024
        $index++
    }
    if ($index -eq 0) { return ('{0:N0} {1}' -f [Math]::Round($value,0), $units[$index]) }
    if ($value -ge 100) { return ('{0:N0} {1}' -f $value, $units[$index]) }
    if ($value -ge 10) { return ('{0:N1} {1}' -f $value, $units[$index]) }
    return ('{0:N2} {1}' -f $value, $units[$index])
}

function Get-MetricSamples {
    param([string]$Text,[string]$MetricName)

    $samples = New-Object System.Collections.Generic.List[object]
    $linePattern = '^' + [regex]::Escape($MetricName) + '(?:\{([^}]*)\})?\s+([^\s]+)\s*$'
    foreach ($line in @([string]$Text -split "`r?`n")) {
        $trim = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trim) -or $trim.StartsWith('#')) { continue }

        $lineMatch = [regex]::Match($trim, $linePattern)
        if (-not $lineMatch.Success) { continue }

        $labelMap = @{}
        $labelText = [string]$lineMatch.Groups[1].Value
        if (-not [string]::IsNullOrWhiteSpace($labelText)) {
            foreach ($pair in ($labelText -split ',')) {
                $pairMatch = [regex]::Match($pair, '^\s*([^=]+)="(.*)"\s*$')
                if ($pairMatch.Success) {
                    $labelMap[[string]$pairMatch.Groups[1].Value] = ([string]$pairMatch.Groups[2].Value -replace '\"','"')
                }
            }
        }

        $valueText = [string]$lineMatch.Groups[2].Value
        $number = Convert-ToDiagnosticsDouble $valueText
        [void]$samples.Add([pscustomobject]@{
            Labels = $labelMap
            Value = $number
            ValueText = $valueText
        })
    }

    return @($samples.ToArray())
}

function Get-MetricSampleValue {
    param([string]$Text,[string]$MetricName,[hashtable]$Labels = @{})

    foreach ($sample in @(Get-MetricSamples -Text $Text -MetricName $MetricName)) {
        $matched = $true
        foreach ($key in $Labels.Keys) {
            if (-not $sample.Labels.ContainsKey($key) -or [string]$sample.Labels[$key] -ne [string]$Labels[$key]) {
                $matched = $false
                break
            }
        }
        if ($matched) { return $sample.Value }
    }

    return $null
}

function Get-DiagnosticsPathLabel {
    param([string]$Path)
    switch ([string]$Path) {
        'derp' { return 'DERP' }
        'direct_ipv4' { return 'Direct IPv4' }
        'direct_ipv6' { return 'Direct IPv6' }
        'peer_relay_ipv4' { return 'Peer Relay IPv4' }
        'peer_relay_ipv6' { return 'Peer Relay IPv6' }
        default { return $Path }
    }
}

function Get-MachineRoleSummary {
    param($Machine)
    if ($null -eq $Machine) { return '-' }
    $roles = New-Object System.Collections.Generic.List[string]
    if ([string](Get-PropertyValue $Machine @('ExitNodeOptionText')) -eq 'True') { [void]$roles.Add('Exit node') }
    $prCountText = [string](Get-PropertyValue $Machine @('PrimaryRoutesCount'))
    if (-not [string]::IsNullOrWhiteSpace($prCountText)) {
        $prCount = 0
        [void][int]::TryParse($prCountText,[ref]$prCount)
        if ($prCount -gt 0) { [void]$roles.Add('Subnet router') }
    }
    if (@($roles).Count -eq 0) { return 'None' }
    return ($roles -join ', ')
}

function Get-ExceptionDiagnosticText {
    param([Parameter(Mandatory=$true)]$ErrorRecord,[string]$Prefix='Error')
    $parts = @()
    try { $parts += ($Prefix + ': ' + [string]$ErrorRecord.Exception.Message) } catch {}
    try { if ($ErrorRecord.Exception.GetType().FullName) { $parts += ('Type: ' + [string]$ErrorRecord.Exception.GetType().FullName) } } catch {}
    try { if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) { $parts += ('Position: ' + ($ErrorRecord.InvocationInfo.PositionMessage -replace "`r?`n", ' | ')) } } catch {}
    try { if ($ErrorRecord.ScriptStackTrace) { $parts += ('Stack: ' + ($ErrorRecord.ScriptStackTrace -replace "`r?`n", ' | ')) } } catch {}
    try { if ($ErrorRecord.Exception.InnerException -and $ErrorRecord.Exception.InnerException.Message) { $parts += ('Inner: ' + [string]$ErrorRecord.Exception.InnerException.Message) } } catch {}
    return ($parts -join [Environment]::NewLine)
}

function Show-MetricsDiagnostics {
    param([switch]$ReturnText)

    $snapshot = Get-CurrentSnapshot
    if (-not $snapshot.Found) { throw 'tailscale.exe was not detected.' }

    $result = Invoke-TailscaleCommand -Exe $snapshot.Exe -Arguments @('metrics')
    $raw = [string]$result.Output
    $lines = @()

    try {
        $adv = Get-MetricSampleValue -Text $raw -MetricName 'tailscaled_advertised_routes'
        $appr = Get-MetricSampleValue -Text $raw -MetricName 'tailscaled_approved_routes'
        $warn = Get-MetricSampleValue -Text $raw -MetricName 'tailscaled_health_messages' -Labels @{ type = 'warning' }
        $homeDerp = Get-MetricSampleValue -Text $raw -MetricName 'tailscaled_home_derp_region_id'

        $lines += 'METRICS'
        $lines += ''
        $lines += 'OVERVIEW'
        $lines += ('  Advertised routes: ' + $(if ($null -eq $adv) { '-' } else { [string][int64][math]::Round([double]$adv,0) }))
        $lines += ('  Approved routes: ' + $(if ($null -eq $appr) { '-' } else { [string][int64][math]::Round([double]$appr,0) }))
        $lines += ('  Health warnings: ' + $(if ($null -eq $warn) { '-' } else { [string][int64][math]::Round([double]$warn,0) }))
        $lines += ('  Home DERP region ID: ' + $(if ($null -eq $homeDerp) { '-' } else { [string][int64][math]::Round([double]$homeDerp,0) }))
        $lines += ''

        $sectionDefs = @(
            @{ Title = 'INBOUND'; Bytes='tailscaled_inbound_bytes_total'; Packets='tailscaled_inbound_packets_total' },
            @{ Title = 'OUTBOUND'; Bytes='tailscaled_outbound_bytes_total'; Packets='tailscaled_outbound_packets_total' }
        )
        $pathOrder = @('direct_ipv4','direct_ipv6','derp','peer_relay_ipv4','peer_relay_ipv6')
        foreach ($section in $sectionDefs) {
            $lines += [string]$section.Title
            $byteEntries = @(Get-MetricSamples -Text $raw -MetricName ([string]$section.Bytes))
            $packetEntries = @(Get-MetricSamples -Text $raw -MetricName ([string]$section.Packets))
            $totalBytes = 0.0
            foreach ($entry in $byteEntries) { $totalBytes += [double]$entry.Value }
            $totalPackets = 0.0
            foreach ($entry in $packetEntries) { $totalPackets += [double]$entry.Value }
            $lines += ('  Total: ' + (Convert-ToHumanBytes ([double]$totalBytes)) + ' | ' + ('{0:N0}' -f ([double]$totalPackets)) + ' packets')
            foreach ($path in $pathOrder) {
                $bytesValue = Get-MetricSampleValue -Text $raw -MetricName ([string]$section.Bytes) -Labels @{ path = [string]$path }
                $packetsValue = Get-MetricSampleValue -Text $raw -MetricName ([string]$section.Packets) -Labels @{ path = [string]$path }
                if ($null -eq $bytesValue -and $null -eq $packetsValue) { continue }
                $bytesSafe = if ($null -eq $bytesValue) { 0.0 } else { [double]$bytesValue }
                $packetsSafe = if ($null -eq $packetsValue) { 0.0 } else { [double]$packetsValue }
                $lines += ('  ' + (Get-DiagnosticsPathLabel ([string]$path)) + ': ' + (Convert-ToHumanBytes $bytesSafe) + ' | ' + ('{0:N0}' -f $packetsSafe) + ' packets')
            }
            $lines += ''
        }

        $dropDefs = @(
            @{ Title='OUTBOUND DROPPED'; Metric='tailscaled_outbound_dropped_packets_total'; Label='reason' },
            @{ Title='INBOUND DROPPED'; Metric='tailscaled_inbound_dropped_packets_total'; Label='reason' }
        )
        foreach ($drop in $dropDefs) {
            $entries = @(Get-MetricSamples -Text $raw -MetricName ([string]$drop.Metric))
            if (@($entries).Count -le 0) { continue }
            $lines += [string]$drop.Title
            foreach ($entry in $entries) {
                $key = 'unknown'
                if ($entry.Labels -is [hashtable] -and $entry.Labels.ContainsKey([string]$drop.Label)) {
                    $key = [string]$entry.Labels[[string]$drop.Label]
                }
                elseif ($entry.Labels -is [hashtable] -and $entry.Labels.ContainsKey('path')) {
                    $key = [string]$entry.Labels['path']
                }
                $lines += ('  ' + $key + ': ' + ('{0:N0}' -f ([double]$entry.Value)))
            }
            $lines += ''
        }

        if ($result.ExitCode -ne 0) {
            $lines += 'NOTE'
            $lines += '  The metrics command returned a non-zero exit code. Parsed values may be incomplete.'
            $lines += ''
        }
    }
    catch {
        $diag = Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Metrics parsing failed'
        $debugLines = @(
            $diag,
            '',
            'PROCESS',
            ('Command: ' + [string]$snapshot.Exe + ' metrics'),
            ('Exit code: ' + [string]$result.ExitCode),
            '',
            'RAW OUTPUT',
            $raw
        )
        $text = ($debugLines -join [Environment]::NewLine)
        try { Write-Log $text } catch { }
        if ($ReturnText) { return $text }
        Set-DiagnosticsOutput -Text $text
        return
    }

    $text = ($lines -join [Environment]::NewLine)
    if ($ReturnText) { return $text }
    Set-DiagnosticsOutput -Text $text
}

function Get-SelectedMachineTargetValue {
    param($Machine,[ValidateSet('DNS','IPv4','IPv6')] [string]$Kind = 'DNS')
    if ($null -eq $Machine) { return '' }
    switch ($Kind) {
        'DNS' { return ConvertTo-DnsName ([string](Get-PropertyValue $Machine @('DNSName'))) }
        'IPv4' { return [string](Get-PropertyValue $Machine @('IPv4')) }
        'IPv6' { return [string](Get-PropertyValue $Machine @('IPv6')) }
    }
    return ''
}

function Set-PingDiagnosticsBusyState {
    param([bool]$Busy,$Button = $null)
    try {
        $pingButtons = @($script:btnCmdPingAll,$script:btnCmdPingDns,$script:btnCmdPingIPv4,$script:btnCmdPingIPv6)
        if ($Busy) {
            $isAll = $false
            try { $isAll = ($Button -eq $script:btnCmdPingAll) } catch { }
            foreach ($btn in @($pingButtons)) {
                if ($null -eq $btn) { continue }
                try {
                    $allowValue = $true
                    try { if ($null -ne $btn.Tag -and $btn.Tag.PSObject.Properties['AllowPing']) { $allowValue = [bool]$btn.Tag.AllowPing } } catch { }
                    $original = [string]$btn.Text
                    try { if ($null -ne $btn.Tag -and $btn.Tag.PSObject.Properties['OriginalText']) { $original = [string]$btn.Tag.OriginalText } } catch { }
                    $btn.Tag = [pscustomobject]@{ AllowPing = $allowValue; OriginalText = $original }
                    if ($isAll -or $btn -eq $Button) { $btn.Text = 'Pinging...' }
                } catch { }
            }
            $script:DiagnosticsBusyFocusedButton = $Button
            Set-NetworkDeviceActionsBusyState -Busy $true
        }
        else {
            foreach ($btn in @($pingButtons)) {
                try {
                    if ($null -ne $btn -and $null -ne $btn.Tag -and $btn.Tag.PSObject.Properties['OriginalText']) {
                        $btn.Text = [string]$btn.Tag.OriginalText
                        $allowValue = $true
                        try { if ($btn.Tag.PSObject.Properties['AllowPing']) { $allowValue = [bool]$btn.Tag.AllowPing } } catch { }
                        $btn.Tag = [pscustomobject]@{ AllowPing = $allowValue }
                    }
                } catch { }
            }
            try { if ($null -ne $script:btnCmdPingAll) { $script:btnCmdPingAll.Text = 'Ping all' } } catch { }
            try { if ($null -ne $script:btnCmdPingDns) { $script:btnCmdPingDns.Text = 'Ping DNS' } } catch { }
            try { if ($null -ne $script:btnCmdPingIPv4) { $script:btnCmdPingIPv4.Text = 'Ping IPv4' } } catch { }
            try { if ($null -ne $script:btnCmdPingIPv6) { $script:btnCmdPingIPv6.Text = 'Ping IPv6' } } catch { }
            Set-NetworkDeviceActionsBusyState -Busy $false
        }
    }
    catch { Write-LogException -Context 'Set ping diagnostics busy state' -ErrorRecord $_ }
}

function Convert-TailscalePeerPingResult {
    param([string]$Target,$Machine,[int]$ExitCode,[string]$Output)
    $baseResult = [ordered]@{
        Target = [string]$Target
        Success = $false
        Conn = ''
        Derp = ''
        AvgLatency = '-'
        LastLatency = '-'
        ExitCode = [int]$ExitCode
        Output = [string]$Output
    }
    if ([string]::IsNullOrWhiteSpace([string]$Target)) { return [pscustomobject]$baseResult }
    $latencyValues = @()
    $successRows = @()
    foreach ($line in @([string]$Output -split "`r?`n")) {
        $trim = ([string]$line).Trim()
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        if ($trim -match 'pong .* via DERP\(([^)]+)\) in ([0-9]+(?:\.[0-9]+)?)ms') {
            $latencyValues += [double]$Matches[2]
            $successRows += [pscustomobject]@{ Conn='Relay'; Derp=(ConvertTo-DiagnosticText -Text ([string]$Matches[1])); Latency=([string]$Matches[2] + 'ms') }
            continue
        }
        if ($trim -match 'pong .* via ([^ ]+) in ([0-9]+(?:\.[0-9]+)?)ms') {
            $via = [string]$Matches[1]
            $conn = 'Direct'
            $derp = ''
            if ($via -match '^DERP\(([^)]+)\)$') { $conn = 'Relay'; $derp = [string]$Matches[1] }
            elseif ($via -match '^(peer-relay|relay)') { $conn = 'Relay'; $derp = $via }
            $latencyValues += [double]$Matches[2]
            $successRows += [pscustomobject]@{ Conn=$conn; Derp=$derp; Latency=([string]$Matches[2] + 'ms') }
            continue
        }
        if ($trim -match 'pong .* in ([0-9]+(?:\.[0-9]+)?)ms') {
            $latencyValues += [double]$Matches[1]
            $successRows += [pscustomobject]@{ Conn=''; Derp=''; Latency=([string]$Matches[1] + 'ms') }
            continue
        }
    }
    if (@($successRows).Count -gt 0) {
        $last = $successRows[@($successRows).Count - 1]
        $baseResult.Success = $true
        $baseResult.Conn = [string]$last.Conn
        $baseResult.Derp = $(if ([string]$last.Conn -eq 'Relay') { [string]$last.Derp } else { '' })
        $baseResult.LastLatency = [string]$last.Latency
        $avg = ($latencyValues | Measure-Object -Average).Average
        if ($null -ne $avg) { $baseResult.AvgLatency = ([Math]::Round([double]$avg,1).ToString([System.Globalization.CultureInfo]::InvariantCulture) + 'ms') }
    }
    else {
        $baseResult.Conn = [string](Get-PropertyValue $Machine @('Connection'))
        if ([string]$baseResult.Conn -eq 'Relay') { $baseResult.Derp = [string](Get-PropertyValue $Machine @('Relay')) }
    }
    return [pscustomobject]$baseResult
}

function Start-TailscalePingProcess {
    param([string]$Exe,[string]$Target,[int]$Count = 2)
    $pingProcessArgs = @('ping',('--c=' + [Math]::Max(1,$Count)),'--verbose','--timeout=4s',$Target)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = ConvertTo-ProcessArgumentString -Arguments $pingProcessArgs
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    try { $psi.StandardOutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    try { $psi.StandardErrorEncoding = [Text.UTF8Encoding]::new($false) } catch { }
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    return [pscustomobject]@{ Process = $proc; Arguments = $pingProcessArgs; StartedUtc = [datetime]::UtcNow }
}

function Complete-PingDiagnosticsTask {
    $taskRef = $script:PingDiagnosticsWorker
    if ($null -eq $taskRef) { return }
    try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
    $machine = $taskRef.Machine
    $kind = [string]$taskRef.Kind
    try {
        $dnsPing = $null
        $ipv4Ping = $null
        $ipv6Ping = $null
        $commandParts = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($taskRef.Items)) {
            $proc = $item.Process
            $stdout = ''
            $stderr = ''
            try { $stdout = [string]$proc.StandardOutput.ReadToEnd() } catch { }
            try { $stderr = [string]$proc.StandardError.ReadToEnd() } catch { }
            $output = (($stdout,$stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
            $exit = 1
            try { $exit = [int]$proc.ExitCode } catch { }
            try { $proc.Dispose() } catch { }
            $parsed = Convert-TailscalePeerPingResult -Target ([string]$item.Target) -Machine $machine -ExitCode $exit -Output $output
            if ([string]$item.Kind -eq 'DNS') { $dnsPing = $parsed }
            elseif ([string]$item.Kind -eq 'IPv4') { $ipv4Ping = $parsed }
            elseif ([string]$item.Kind -eq 'IPv6') { $ipv6Ping = $parsed }
            [void]$commandParts.Add('tailscale ' + (($item.Arguments | ForEach-Object { [string]$_ }) -join ' '))
        }
        $path = [string](Get-PropertyValue $machine @('Connection'))
        $derp = [string](Get-PropertyValue $machine @('Relay'))
        foreach ($candidate in @($dnsPing,$ipv4Ping,$ipv6Ping)) {
            if ($null -eq $candidate -or -not [bool]$candidate.Success) { continue }
            if ([string]$candidate.Conn -eq 'Direct') { $path = 'Direct'; $derp = ''; break }
            if ([string]$candidate.Conn -eq 'Relay') {
                $path = 'Relay'
                if (-not [string]::IsNullOrWhiteSpace([string]$candidate.Derp)) { $derp = [string]$candidate.Derp }
            }
        }
        $elapsed = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
        $modes = $taskRef.Modes
        $dnsTarget = ConvertTo-DnsName ([string](Get-PropertyValue $machine @('DNSName','DNS')))
        $ipv4Target = [string](Get-PropertyValue $machine @('IPv4'))
        $ipv6Target = [string](Get-PropertyValue $machine @('IPv6'))
        $state = [pscustomobject]@{
            MachineKey = Get-MachineKey -Machine $machine
            Path = $path
            Derp = $derp
            DNSResult = $dnsPing
            IPv4Result = $ipv4Ping
            IPv6Result = $ipv6Ping
            DNSMs = $(if ([bool]$modes.DNS) { $(if ([string]::IsNullOrWhiteSpace($dnsTarget)) { 'n/a' } else { Get-PingLatencyText -Ping $dnsPing }) } else { '-' })
            IPv4Ms = $(if ([bool]$modes.IPv4) { $(if ([string]::IsNullOrWhiteSpace($ipv4Target)) { 'n/a' } else { Get-PingLatencyText -Ping $ipv4Ping }) } else { '-' })
            IPv6Ms = $(if ([bool]$modes.IPv6) { $(if ([string]::IsNullOrWhiteSpace($ipv6Target)) { 'n/a' } else { Get-PingLatencyText -Ping $ipv6Ping }) } else { '-' })
            DurationMs = [double]$elapsed
        }
        $state | Add-Member -NotePropertyName Details -NotePropertyValue (Build-PingDetailText -Machine $machine -State $state) -Force
        Set-PingStateForMachine -MachineKey (Get-MachineKey -Machine $machine) -State $state
        if ($null -ne $script:gridPing) {
            foreach ($row in @($script:gridPing.Rows | Where-Object { $null -ne $_ -and $null -ne $_.Tag })) {
                if ((Get-MachineKey -Machine $row.Tag.Machine) -eq (Get-MachineKey -Machine $machine)) { Set-PingGridRowValues -Row $row -Machine $machine -State $state; break }
            }
        }
        try { Update-MachineDetailsView -Machine $machine } catch { Write-LogException -Context 'Render machine details from ping result' -ErrorRecord $_ }
        try { Update-DiagnosticsSelectionSummary -Machine $machine } catch { Write-LogException -Context 'Update diagnostics summary from ping result' -ErrorRecord $_ }
        $detailText = ConvertTo-DiagnosticText -Text ([string]$state.Details)
        Set-DiagnosticsOutput -Text $detailText -Mode 'Selection'
        $targetMachine = [string](Get-PropertyValue $machine @('Machine'))
        $targetKind = if ([string]::IsNullOrWhiteSpace([string]$kind)) { 'Ping' } else { 'Ping ' + [string]$kind }
        Write-ActivityCommandBlock -Title ($targetKind + ' for ' + $targetMachine) -CommandText (($commandParts.ToArray()) -join [Environment]::NewLine) -ExitCode 0 -Output $detailText -DurationMs ([double]$state.DurationMs)
    }
    catch {
        $message = [string]$_.Exception.Message
        Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Ping failed')
        Write-ActivityFailureBlock -Title 'Ping failed' -CommandText 'tailscale ping' -Message $message
    }
    finally {
        $script:IsPingDiagnosticsTaskRunning = $false
        $script:PingDiagnosticsWorker = $null
        $script:DiagnosticsBusyFocusedButton = $null
        Set-PingDiagnosticsBusyState -Busy $false
    }
}

function Start-SelectedPingDiagnosticsAsync {
    param([ValidateSet('DNS','IPv4','IPv6','All')] [string]$Kind = 'DNS',$SourceButton = $null)
    if ($script:IsPingDiagnosticsTaskRunning -or $script:IsDiagnosticsCommandTaskRunning) { return }
    $machine = Get-SelectedMachine
    if ($null -eq $machine) { throw 'Select a device first.' }
    if ([bool](Get-PropertyValue $machine @('IsLocal'))) { Set-DiagnosticsOutput -Text 'The local device should not be pinged from here.' -Mode 'Selection'; return }
    if ([string](Get-PropertyValue $machine @('Status')) -eq 'Offline') { Set-DiagnosticsOutput -Text 'The selected device is offline. Ping is disabled until it comes back online.' -Mode 'Selection'; return }
    $modes = [pscustomobject]@{ DNS = $false; IPv4 = $false; IPv6 = $false }
    switch ($Kind) {
        'DNS'  { $modes.DNS = $true }
        'IPv4' { $modes.IPv4 = $true }
        'IPv6' { $modes.IPv6 = $true }
        'All'  { $modes.DNS = $true; $modes.IPv4 = $true; $modes.IPv6 = $true }
    }
    $exe = ''
    try { if ($null -ne $script:Snapshot -and -not [string]::IsNullOrWhiteSpace([string]$script:Snapshot.Exe)) { $exe = [string]$script:Snapshot.Exe } } catch { }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { $exe = Find-TailscaleExe }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { throw 'tailscale.exe not found.' }
    $targets = New-Object System.Collections.Generic.List[object]
    $dnsTarget = ConvertTo-DnsName ([string](Get-PropertyValue $machine @('DNSName','DNS')))
    $ipv4Target = [string](Get-PropertyValue $machine @('IPv4'))
    $ipv6Target = [string](Get-PropertyValue $machine @('IPv6'))
    if ([bool]$modes.DNS -and -not [string]::IsNullOrWhiteSpace($dnsTarget)) { [void]$targets.Add([pscustomobject]@{ Kind='DNS'; Target=$dnsTarget }) }
    if ([bool]$modes.IPv4 -and -not [string]::IsNullOrWhiteSpace($ipv4Target)) { [void]$targets.Add([pscustomobject]@{ Kind='IPv4'; Target=$ipv4Target }) }
    if ([bool]$modes.IPv6 -and -not [string]::IsNullOrWhiteSpace($ipv6Target)) { [void]$targets.Add([pscustomobject]@{ Kind='IPv6'; Target=$ipv6Target }) }
    if ($targets.Count -eq 0) { throw 'No ping target is available for this device.' }
    $items = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($target in @($targets.ToArray())) {
            $started = Start-TailscalePingProcess -Exe $exe -Target ([string]$target.Target) -Count 2
            [void]$items.Add([pscustomobject]@{ Kind=[string]$target.Kind; Target=[string]$target.Target; Process=$started.Process; Arguments=$started.Arguments; StartedUtc=$started.StartedUtc })
        }
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $task = [pscustomobject]@{ Kind=[string]$Kind; Machine=$machine; Modes=$modes; Items=$items.ToArray(); Timer=$timer; StartedUtc=[datetime]::UtcNow }
        $script:IsPingDiagnosticsTaskRunning = $true
        $script:PingDiagnosticsWorker = $task
        $script:DiagnosticsBusyFocusedButton = $SourceButton
        Set-PingDiagnosticsBusyState -Busy $true -Button $SourceButton
        $timer.add_Tick({
            try {
                $taskRef = $script:PingDiagnosticsWorker
                if ($null -eq $taskRef) { return }
                $allDone = $true
                foreach ($item in @($taskRef.Items)) {
                    try { if ($null -ne $item.Process -and -not $item.Process.HasExited) { $allDone = $false; break } } catch { }
                }
                if (-not $allDone) { return }
                Complete-PingDiagnosticsTask
            }
            catch {
                Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Ping failed')
                $script:IsPingDiagnosticsTaskRunning = $false
                $script:PingDiagnosticsWorker = $null
                Set-PingDiagnosticsBusyState -Busy $false
            }
        })
        $timer.Start()
    }
    catch {
        foreach ($item in @($items.ToArray())) {
            try { if ($null -ne $item.Process -and -not $item.Process.HasExited) { $item.Process.Kill() } } catch { }
            try { if ($null -ne $item.Process) { $item.Process.Dispose() } } catch { }
        }
        $script:IsPingDiagnosticsTaskRunning = $false
        $script:PingDiagnosticsWorker = $null
        $script:DiagnosticsBusyFocusedButton = $null
        Set-PingDiagnosticsBusyState -Busy $false
        throw
    }
}

function Show-SelectedPingDiagnostics {
    param([ValidateSet('DNS','IPv4','IPv6','All')] [string]$Kind = 'DNS',$SourceButton = $null)
    Start-SelectedPingDiagnosticsAsync -Kind $Kind -SourceButton $SourceButton
}

function Get-PingLatencyText {
    param($Ping)
    if ($null -eq $Ping) { return '-' }
    if (-not [bool]$Ping.Success) { return 'Fail' }
    if (-not [string]::IsNullOrWhiteSpace([string]$Ping.AvgLatency) -and [string]$Ping.AvgLatency -ne '-') { return [string]$Ping.AvgLatency }
    if (-not [string]::IsNullOrWhiteSpace([string]$Ping.LastLatency) -and [string]$Ping.LastLatency -ne '-') { return [string]$Ping.LastLatency }
    return 'OK'
}
function Build-PingDetailText {
    param($Machine,$State)
    if ($null -eq $Machine) {
        return 'Select a device in Machines, then run a selected-device command to review the latest results.'
    }

    $overviewPath = $(if ([bool](Get-PropertyValue $Machine @('IsLocal'))) { 'Local' } elseif ($null -ne $State -and -not [string]::IsNullOrWhiteSpace([string]$State.Path)) { [string]$State.Path } else { [string]$Machine.Connection })
    $overviewDerp = ''
    if ($overviewPath -eq 'Relay') {
        if ($null -ne $State -and -not [string]::IsNullOrWhiteSpace([string]$State.Derp)) { $overviewDerp = [string]$State.Derp }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$Machine.Relay)) { $overviewDerp = [string]$Machine.Relay }
    }

    $lines = @(
        'OVERVIEW',
        ('  Device      ' + [string]$Machine.Machine),
        ('  Owner       ' + [string]$Machine.Owner),
        ('  Online      ' + $(if ([string]$Machine.Status -eq 'Offline' -or [string]$Machine.OnlineText -eq 'False') { 'No' } else { 'Yes' })),
        ('  Last seen   ' + [string]$Machine.LastSeen),
        ('  Conn        ' + $overviewPath),
        ('  DERP        ' + $(if ([string]::IsNullOrWhiteSpace($overviewDerp)) { '-' } else { $overviewDerp })),
        '',
        'TARGETS',
        ('  DNS         ' + (ConvertTo-DnsName ([string]$Machine.DNSName))),
        ('  IPv4        ' + [string]$Machine.IPv4),
        ('  IPv6        ' + [string]$Machine.IPv6)
    )

    if ($null -eq $State) {
        $lines += ''
        $lines += 'RESULTS'
        $lines += '  No ping has been run for this device yet.'
        return ($lines -join [Environment]::NewLine)
    }

    $lines += ''
    $lines += 'RESULTS'
    foreach ($entry in @(
        @{ Name='DNS'; Result=$State.DNSResult },
        @{ Name='IPv4'; Result=$State.IPv4Result },
        @{ Name='IPv6'; Result=$State.IPv6Result }
    )) {
        $r = $entry.Result
        if ($null -eq $r) {
            $lines += ('  {0,-11} Not tested' -f $entry.Name)
            continue
        }

        $statusText = if ([bool]$r.Success) { 'OK' } else { 'Fail' }
        $pathText = if ([string]::IsNullOrWhiteSpace([string]$r.Conn)) { '-' } else { [string]$r.Conn }
        $derpText = if ([string]$r.Conn -eq 'Relay' -and -not [string]::IsNullOrWhiteSpace([string]$r.Derp)) { [string]$r.Derp } else { '-' }
        $lastText = if ([string]::IsNullOrWhiteSpace([string]$r.LastLatency)) { '-' } else { [string]$r.LastLatency }

        $lines += ('  {0,-11} {1,-5}  {2,-6}  last {3,-8}  derp {4}' -f $entry.Name, $statusText, $pathText, $lastText, $derpText)
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-PingStateForMachine {
    param($MachineKey)
    if ([string]::IsNullOrWhiteSpace([string]$MachineKey)) { return $null }
    if ($script:PingState.ContainsKey($MachineKey)) { return $script:PingState[$MachineKey] }
    return $null
}

function Set-PingStateForMachine {
    param($MachineKey,$State)
    if ([string]::IsNullOrWhiteSpace([string]$MachineKey)) { return }
    $script:PingState[$MachineKey] = $State
}

function Set-PingGridRowValues {
    param($Row,$Machine,$State)
    if ($null -eq $Row -or $null -eq $Machine) { return }
    $onlineText = if ([string]$Machine.Status -eq 'Offline' -or [string]$Machine.OnlineText -eq 'False') { 'No' } else { 'Yes' }
    $pathText = if ($null -ne $State -and -not [string]::IsNullOrWhiteSpace([string]$State.Path)) { [string]$State.Path } else { [string]$Machine.Connection }
    $derpText = if ([string]$pathText -eq 'Relay') {
        if ($null -ne $State -and -not [string]::IsNullOrWhiteSpace([string]$State.Derp)) { [string]$State.Derp }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$Machine.Relay)) { [string]$Machine.Relay }
        else { '-' }
    }
    else { '' }

    $Row.Cells['PingDevice'].Value = [string]$Machine.Machine
    $Row.Cells['PingOwner'].Value = [string]$Machine.Owner
    $Row.Cells['PingOnline'].Value = $onlineText
    $Row.Cells['PingLastSeen'].Value = [string]$Machine.LastSeen
    $Row.Cells['PingDnsTarget'].Value = ConvertTo-DnsName ([string]$Machine.DNSName)
    $Row.Cells['PingIPv4Target'].Value = [string]$Machine.IPv4
    $Row.Cells['PingIPv6Target'].Value = [string]$Machine.IPv6
    $Row.Cells['PingPath'].Value = $pathText
    $Row.Cells['PingDerp'].Value = $derpText
    $Row.Cells['PingDnsMs'].Value = $(if ($null -ne $State) { [string]$State.DNSMs } else { '-' })
    $Row.Cells['PingIPv4Ms'].Value = $(if ($null -ne $State) { [string]$State.IPv4Ms } else { '-' })
    $Row.Cells['PingIPv6Ms'].Value = $(if ($null -ne $State) { [string]$State.IPv6Ms } else { '-' })
    $Row.Tag = [pscustomobject]@{ Machine = $Machine; PingState = $State }
}

function Update-PingTargets {
    param($Snapshot,$PreferredMachine = $null)
    if ($null -eq $script:gridPing) { return }

    $selectedKey = ''
    try {
        if ($null -ne $script:gridPing.CurrentRow -and $null -ne $script:gridPing.CurrentRow.Tag) {
            $selectedKey = Get-MachineKey -Machine $script:gridPing.CurrentRow.Tag.Machine
        }
    }
    catch { Write-LogException -Context 'Update ping button state' -ErrorRecord $_ }

    $preferredKey = Get-MachineKey -Machine $PreferredMachine
    $machines = @(
        Convert-ToObjectArray $Snapshot.Machines |
        Where-Object { $null -ne $_ -and [string](Get-PropertyValue $_ @('Status')) -ne 'This device' } |
        Sort-Object @{ Expression = { [string](Get-PropertyValue $_ @('Machine')) } }, @{ Expression = { [string](Get-PropertyValue $_ @('Owner')) } }
    )

    $script:gridPing.Rows.Clear()
    foreach ($machine in $machines) {
        $rowIndex = $script:gridPing.Rows.Add()
        $row = $script:gridPing.Rows[$rowIndex]
        $state = Get-PingStateForMachine -MachineKey (Get-MachineKey -Machine $machine)
        Set-PingGridRowValues -Row $row -Machine $machine -State $state
    }

    $pick = $null
    if (-not [string]::IsNullOrWhiteSpace($preferredKey)) { $pick = $preferredKey }
    elseif (-not [string]::IsNullOrWhiteSpace($selectedKey)) { $pick = $selectedKey }

    if (-not [string]::IsNullOrWhiteSpace($pick)) {
        for ($i = 0; $i -lt $script:gridPing.Rows.Count; $i++) {
            $row = $script:gridPing.Rows[$i]
            if ($null -ne $row.Tag -and (Get-MachineKey -Machine $row.Tag.Machine) -eq $pick) {
                $row.Selected = $true
                $script:gridPing.CurrentCell = $row.Cells[0]
                break
            }
        }
    }

    Update-PingSelection -Machine $PreferredMachine
    Update-DiagnosticsSelectionSummary -Machine $PreferredMachine
}

function Update-PingSelection {
    param($Machine = $null)
    if ($null -eq $script:gridPing) { return }

    if ($null -ne $Machine) {
        $targetKey = Get-MachineKey -Machine $Machine
        for ($i = 0; $i -lt $script:gridPing.Rows.Count; $i++) {
            $row = $script:gridPing.Rows[$i]
            if ($null -ne $row.Tag -and (Get-MachineKey -Machine $row.Tag.Machine) -eq $targetKey) {
                $script:gridPing.ClearSelection()
                $row.Selected = $true
                $script:gridPing.CurrentCell = $row.Cells[0]
                break
            }
        }
    }

    if ($null -eq $script:txtPingDetails) { return }
    $selectedDiagMachine = $null
    try { if ($null -ne $script:gridPing.CurrentRow -and $null -ne $script:gridPing.CurrentRow.Tag) { $selectedDiagMachine = $script:gridPing.CurrentRow.Tag.Machine } } catch { Write-LogException -Context 'Read selected diagnostics machine from grid' -ErrorRecord $_ }
    Update-DiagnosticsSelectionSummary -Machine $selectedDiagMachine
    Update-SelectedDeviceActionButtons
    if ([string]$script:DiagnosticsContentMode -eq 'Command' -or [string]$script:DiagnosticsContentMode -eq 'PingAll') { return }
    if ($null -eq $script:gridPing.CurrentRow -or $null -eq $script:gridPing.CurrentRow.Tag) {
        $script:DiagnosticsContentMode = 'Selection'
        Set-DiagnosticsOutput -Text 'Select a device in Machines, then run a local or selected-device command from the Commands tab.' -Mode 'Selection'
        return
    }

    $tag = $script:gridPing.CurrentRow.Tag
    $state = $null
    try { $state = $tag.PingState } catch { Write-LogException -Context 'Read ping state from machine tag' -ErrorRecord $_ }
    Update-DiagnosticsSelectionSummary -Machine $tag.Machine
    Set-DiagnosticsOutput -Text (Build-PingDetailText -Machine $tag.Machine -State $state) -Mode 'Selection'
}

function Update-UiFromSnapshot {
    param($Snapshot)
    $script:Snapshot = $Snapshot
    $backendState = [string](Get-ObjectPropertyOrDefault $Snapshot 'BackendState' '')
    $versionText = [string](Get-ObjectPropertyOrDefault $Snapshot 'Version' '')
    $userText = [string](Get-ObjectPropertyOrDefault $Snapshot 'User' '')
    $userEmailText = [string](Get-ObjectPropertyOrDefault $Snapshot 'UserEmail' '')
    $tailnetText = [string](Get-ObjectPropertyOrDefault $Snapshot 'Tailnet' '')
    $shortNameText = [string](Get-ObjectPropertyOrDefault $Snapshot 'ShortName' '')
    $dnsNameText = [string](Get-ObjectPropertyOrDefault $Snapshot 'DNSName' '')
    $ipv4Text = [string](Get-ObjectPropertyOrDefault $Snapshot 'IPv4' '')
    $ipv6Text = [string](Get-ObjectPropertyOrDefault $Snapshot 'IPv6' '')
    $mtuIPv4Text = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuIPv4' '')
    $mtuIPv6Text = [string](Get-ObjectPropertyOrDefault $Snapshot 'MtuIPv6' '')
    $corpDnsValue = Get-ObjectPropertyOrDefault $Snapshot 'CorpDNS' $null
    Set-UiValue $script:lblBackend $backendState
    Set-UiValue $script:lblVersion $versionText
    $fullVersionText = [string](Get-ObjectPropertyOrDefault $Snapshot 'FullVersion' '')
    if ([string]::IsNullOrWhiteSpace($fullVersionText)) { $fullVersionText = $versionText }
    $buildText = '-'
    if (-not [string]::IsNullOrWhiteSpace($fullVersionText) -and $fullVersionText -match '^[^-]+-(.+)$') { $buildText = [string]$Matches[1] }
    $shortVersionText = $(if ([string]::IsNullOrWhiteSpace($versionText)) { '-' } else { $versionText })
    $fullVersionDisplay = $(if ([string]::IsNullOrWhiteSpace($fullVersionText)) { '-' } else { $fullVersionText })
    Set-AppToolTip -Control $script:lblVersion -Text ("Tailscale client version detected on this device.`r`n`r`nShort version: " + $shortVersionText + "`r`nFull version: " + $fullVersionDisplay + "`r`nBuild/commit: " + $buildText)
    Set-UiValue $script:lblUser $userText
    Set-UiValue $script:lblUserEmail $userEmailText
    Set-UiValue $script:lblTailnet (ConvertTo-DnsName $tailnetText)
    Set-UiValue $script:lblDevice $shortNameText
    Set-UiValue $script:lblDnsName (ConvertTo-DnsName $dnsNameText)
    Set-UiValue $script:lblIPv4 $ipv4Text
    Set-UiValue $script:lblIPv6 $ipv6Text
    Set-UiValue $script:lblMtuIPv4 $mtuIPv4Text
    Set-UiValue $script:lblMtuIPv6 $mtuIPv6Text
    Set-UiValue $script:lblDnsInUse (Get-DnsResolverChainDisplay -Snapshot $Snapshot -PreferSystem:($backendState -ne 'Running'))
    try { Set-AppToolTip -Control $script:lblDnsInUse -Text 'Shows the active DNS path. Split DNS routes such as ts.net are shown in DNS diagnostics, not here.' } catch { }
    try { Update-DnsResolveServerPreview -Snapshot $Snapshot } catch { }
    $dnsMode = if ($null -eq $corpDnsValue) { 'Unknown' } elseif (Convert-ToNullableBool $corpDnsValue) { 'Tailscale' } else { 'System' }
    Set-UiValue $script:lblConnSummary $dnsMode
    $dnsKnown = ($null -ne $corpDnsValue)
    $dnsOn = [bool]($dnsKnown -and (Convert-ToNullableBool $corpDnsValue))
    Set-StateValue -ValueControl $script:lblDnsState -IndicatorControl $script:indDnsState -Value $(if (-not $dnsKnown) { '' } elseif ($dnsOn) { 'On' } else { 'Off' }) -Known:$dnsKnown -On:$dnsOn
    $routeValue = Get-ObjectPropertyOrDefault $Snapshot 'RouteAll' $null
    $routeKnown = ($null -ne $routeValue)
    $routeOn = [bool]($routeKnown -and (Convert-ToNullableBool $routeValue))
    Set-StateValue -ValueControl $script:lblRoutesState -IndicatorControl $script:indRoutesState -Value $(if (-not $routeKnown) { '' } elseif ($routeOn) { 'On' } else { 'Off' }) -Known:$routeKnown -On:$routeOn
    $incomingValue = Get-ObjectPropertyOrDefault $Snapshot 'IncomingAllowed' $null
    $incomingKnown = ($null -ne $incomingValue)
    $incomingOn = [bool]($incomingKnown -and (Convert-ToNullableBool $incomingValue))
    Set-StateValue -ValueControl $script:lblIncomingState -IndicatorControl $script:indIncomingState -Value $(if (-not $incomingKnown) { '' } elseif ($incomingOn) { 'Allowed' } else { 'Blocked' }) -Known:$incomingKnown -On:$incomingOn
    $exitNodeText = Resolve-ExitNodeDisplay $Snapshot
    $exitNodeOn = -not [string]::IsNullOrWhiteSpace([string](Get-ObjectPropertyOrDefault $Snapshot 'CurrentExitNode' ''))
    Set-StateValue -ValueControl $script:lblExitState -IndicatorControl $script:indExitState -Value $exitNodeText -Known:$true -On:$exitNodeOn
    Set-UiValue $script:lblMaintAutoUpdate (Get-TailscaleClientUpdateModeText)
    Update-TailscaleClientUpdateUi -Snapshot $Snapshot
    try {
        Update-ControlAppUpdateUi
    }
    catch { Write-LogException -Context 'Run selected command block' -ErrorRecord $_ }
    Set-UiValue $script:lblMaintMtuStatus ([string]$Snapshot.MtuStatus)
    Set-UiValue $script:lblMaintMtuVersion ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuVersion)) { '-' } else { [string]$Snapshot.MtuVersion }))
    Set-UiValue $script:lblMaintMtuService (([string]$Snapshot.MtuServiceState) + $(if (-not [string]::IsNullOrWhiteSpace([string]$Snapshot.MtuServiceName)) { ' (' + [string]$Snapshot.MtuServiceName + ')' } else { '' }))
    Set-UiValue $script:lblMaintMtuDesiredIPv4 ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuDesiredIPv4)) { '-' } else { [string]$Snapshot.MtuDesiredIPv4 }))
    Set-UiValue $script:lblMaintMtuDesiredIPv6 ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuDesiredIPv6)) { '-' } else { [string]$Snapshot.MtuDesiredIPv6 }))
    Set-UiValue $script:lblMaintMtuCheckInterval ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuCheckInterval)) { '-' } else { [string]$Snapshot.MtuCheckInterval + 's' }))
        Set-UiValue $script:lblMaintMtuLastResult ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuLastResult)) { '-' } else { [string]$Snapshot.MtuLastResult }))
    Set-UiValue $script:lblMaintMtuLastError ($(if ([string]::IsNullOrWhiteSpace([string]$Snapshot.MtuLastError)) { '-' } else { [string]$Snapshot.MtuLastError }))
    if ($null -ne $script:btnInstallMtu -and $null -ne $script:btnOpenMtu) {
        $installed = [bool]$Snapshot.MtuInstalled
        $script:btnInstallMtu.Visible = -not $installed
        $script:btnInstallMtu.Enabled = (-not $installed -and -not [bool]$script:IsMtuInstallRunning)
        $script:btnInstallMtu.Text = $(if (-not $installed -and [bool]$script:IsMtuInstallRunning) { 'Installing...' } else { 'Install Tailscale MTU' })
        $script:btnOpenMtu.Visible = $installed
        $script:btnOpenMtu.Enabled = $installed
    }
    if ($null -ne $script:txtLog) { Update-ActivityView -Text $Snapshot.LogTail }
    if ($null -ne $script:cmbExitNode) {
        $selected = if ($null -ne $script:cmbExitNode.SelectedItem) { ConvertTo-DnsName ([string]$script:cmbExitNode.SelectedItem) } else { '' }
        $preferred = Get-PreferredExitNodeLabel
        $script:cmbExitNode.BeginUpdate()
        $script:cmbExitNode.Items.Clear()
        foreach ($node in @($Snapshot.ExitNodes)) {
            $label = ConvertTo-DnsName ([string]$node.Name)
            if ([string]::IsNullOrWhiteSpace($label)) { $label = ConvertTo-DnsName ([string]$node.DNSName) }
            if (-not [string]::IsNullOrWhiteSpace($label)) { [void]$script:cmbExitNode.Items.Add($label) }
        }
        if (-not [string]::IsNullOrWhiteSpace($preferred) -and $script:cmbExitNode.Items.Contains($preferred)) { $script:cmbExitNode.SelectedItem = $preferred }
        elseif (-not [string]::IsNullOrWhiteSpace($selected) -and $script:cmbExitNode.Items.Contains($selected)) { $script:cmbExitNode.SelectedItem = $selected }
        elseif ($script:cmbExitNode.Items.Count -gt 0) { $script:cmbExitNode.SelectedIndex = 0 }
        $script:cmbExitNode.EndUpdate()
    }
    Update-ExitNodeActionAvailability -Snapshot $Snapshot
    Update-MachinesView -Snapshot $Snapshot
    try { Update-AccountView -Snapshot $Snapshot } catch { Write-LogException -Context 'Update account view from snapshot' -ErrorRecord $_ }
    $selectedMachineNow = Get-SelectedMachine
    Update-PingTargets -Snapshot $Snapshot -PreferredMachine $selectedMachineNow
    if ($null -ne $selectedMachineNow) { try { Update-MachineDetailsView -Machine $selectedMachineNow } catch { Write-LogException -Context 'Render machine details after selection restore' -ErrorRecord $_ } }
    Update-DiagnosticsSelectionSummary -Machine $selectedMachineNow
    Update-SelectedDeviceActionButtons
    Update-StatusBanner -Snapshot $Snapshot
    Update-TrayText -Snapshot $Snapshot
    Update-TrayMenuState -Snapshot $Snapshot
    if ($null -ne $script:toolStatusLabel) { $script:toolStatusLabel.Text = $(if ([bool](Get-ObjectPropertyOrDefault $Snapshot 'Found' $false)) { 'Last refresh: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } else { 'Tailscale not detected.' }) }
}

function Start-StatusRefreshAsync {
    param($Button = $null)
    if ($script:IsRefreshing) { return }
    $forceAccountRefresh = [bool]($null -ne $Button)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 50
    $timer.add_Tick({
        $this.Stop()
        $this.Dispose()
        Update-Status -RefreshAccounts:$forceAccountRefresh
    }.GetNewClosure())
    $timer.Start()
}

function Update-Status {
    param([switch]$RefreshAccounts)
    if ($script:IsRefreshing) { return }
    $script:IsRefreshing = $true
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $snapshot = Get-TailscaleSnapshot
        if ($RefreshAccounts) { Update-LoggedAccountsView -Force $true }
        Update-UiFromSnapshot -Snapshot $snapshot
        if (($null -ne $script:chkLogRefreshActivity -and $script:chkLogRefreshActivity.Checked) -or ($null -eq $script:chkLogRefreshActivity -and [bool](Get-ObjectPropertyOrDefault (Get-Config) 'log_refresh_activity' $false))) {
            try { $sw.Stop() } catch { }
            Write-Log ('Refresh completed. Duration: {0:N0} ms' -f [double]$sw.Elapsed.TotalMilliseconds)
            try { Update-ActivityView -Text (Get-ActivityLogTail) } catch { }
            try {
                $activityRefreshTimer = New-Object System.Windows.Forms.Timer
                $activityRefreshTimer.Interval = 150
                $activityRefreshTimer.add_Tick({
                    param($sender,$eventArgs)
                    try { $sender.Stop(); $sender.Dispose() } catch { }
                    try { Update-ActivityView -Text (Get-ActivityLogTail) } catch { }
                })
                $activityRefreshTimer.Start()
            } catch { }
        }
    }
    catch {
        try { $sw.Stop() } catch { }
        Write-Log ('Status refresh failed: ' + $_.Exception.Message)
        try { if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { Write-Log ('Status refresh position: ' + ($_.InvocationInfo.PositionMessage -replace "`r?`n", ' | ')) } } catch {}
        $message = [string]$_.Exception.Message
        if ($message -match '(?i)tailscale(\.exe)?\s+was\s+not\s+detected|tailscale\s+executable\s+was\s+not\s+found') {
            try {
                $snapshot = Get-TailscaleSnapshot
                Update-UiFromSnapshot -Snapshot $snapshot
            }
            catch { }
            if ($null -ne $script:toolStatusLabel) { $script:toolStatusLabel.Text = 'Tailscale not detected.' }
        }
        else {
            Show-Overlay -Title 'Refresh failed' -Message $message -ErrorStyle -Indicator 'Warn'
            if ($null -ne $script:toolStatusLabel) { $script:toolStatusLabel.Text = 'Refresh failed.' }
        }
    }
    finally {
        try { if ($sw.IsRunning) { $sw.Stop() } } catch { }
        $script:IsRefreshing = $false
        try { $script:RefreshCountSinceGc = [int]$script:RefreshCountSinceGc + 1; Invoke-MemoryCleanupThrottled } catch { }
    }
}

function Set-TailscaleClientMaintenanceBusyState {
    param([bool]$Busy,[string]$Operation)
    try {
        $taskReady = Test-TailscaleClientElevatedTasksReady
        if ($null -ne $script:btnInstallClientAutoUpdateTask) { $script:btnInstallClientAutoUpdateTask.Enabled = (-not $Busy -and -not $script:IsClientTaskSetupRunning) }
        if ($null -ne $script:btnCheckClientUpdate) {
            $script:btnCheckClientUpdate.Text = $(if ($Busy -and $Operation -eq 'Check') { 'Checking...' } else { 'Check Update' })
            $script:btnCheckClientUpdate.Enabled = $(if ($Busy) { $false } else { $taskReady })
            if ($Busy -and $Operation -eq 'Check') { Clear-UiFocusSoon }
        }
        if ($null -ne $script:btnRunClientUpdate) {
            if ($Busy) {
                $script:btnRunClientUpdate.Enabled = $false
                $script:btnRunClientUpdate.Text = $(if ($Operation -eq 'Update') { 'Updating...' } else { 'Update' })
                if ($Operation -eq 'Update') { Clear-UiFocusSoon }
            }
            else {
                $script:btnRunClientUpdate.Text = 'Update'
                $currentVersion = ''
                try { if ($null -ne $script:Snapshot) { $currentVersion = ConvertTo-PlainVersion ([string]$script:Snapshot.Version) } } catch { }
                $latestVersion = ''
                try { $latestVersion = [string](Get-ObjectPropertyOrDefault (Get-Config) 'last_client_update_latest_version' '') } catch { }
                $script:btnRunClientUpdate.Enabled = ((Test-VersionNewer -CurrentVersion $currentVersion -LatestVersion $latestVersion) -and (Test-TailscaleClientElevatedTasksReady))
            }
        }
    }
    catch {
        Write-LogException -Context 'Set Tailscale maintenance busy state' -ErrorRecord $_
    }
}

function Get-TailscaleClientElevatedTaskName {
    param([ValidateSet('Check','Update')][string]$Operation)
    if ($Operation -eq 'Update') { return [string]$script:TailscaleClientUpdateTaskName }
    return [string]$script:TailscaleClientCheckTaskName
}

function Get-TailscaleClientElevatedResultPath {
    param([ValidateSet('Check','Update')][string]$Operation)
    Initialize-AppRoot
    if ($Operation -eq 'Update') { return [string]$script:TailscaleClientUpdateResultPath }
    return [string]$script:TailscaleClientCheckResultPath
}

function Test-TailscaleClientElevatedTaskReadyDirect {
    param([ValidateSet('Check','Update')][string]$Operation)
    try {
        $taskName = Get-TailscaleClientElevatedTaskName -Operation $Operation
        $task = Get-ScheduledTask -TaskPath $script:TailscaleClientTaskPath -TaskName $taskName -ErrorAction Stop
        $action = @($task.Actions) | Select-Object -First 1
        $expectedExe = [IO.Path]::GetFullPath([string]$script:WScriptExe)
        $actualExe = ''
        if ($null -ne $action -and -not [string]::IsNullOrWhiteSpace([string]$action.Execute)) {
            $actualExe = [IO.Path]::GetFullPath([string]$action.Execute)
        }
        $usesHiddenLauncher = ($actualExe -ieq $expectedExe -and ([string]$action.Arguments) -like ('*' + [string]$script:TailscaleClientElevatedLauncherPath + '*'))
        return ($null -ne $task -and $usesHiddenLauncher -and (Test-Path -LiteralPath $script:TailscaleClientElevatedRunnerPath) -and (Test-Path -LiteralPath $script:TailscaleClientElevatedLauncherPath))
    }
    catch { return $false }
}

function Refresh-TailscaleClientElevatedTaskCache {
    try {
        $checkReady = Test-TailscaleClientElevatedTaskReadyDirect -Operation 'Check'
        $updateReady = Test-TailscaleClientElevatedTaskReadyDirect -Operation 'Update'
        $script:TailscaleClientElevatedTaskReadyCache = @{ Check = [bool]$checkReady; Update = [bool]$updateReady }
        $script:TailscaleClientElevatedTaskCacheInitialized = $true
        $script:TailscaleClientElevatedTaskCacheCheckedUtc = [datetime]::UtcNow
        return ([bool]$checkReady -and [bool]$updateReady)
    }
    catch {
        $script:TailscaleClientElevatedTaskReadyCache = @{ Check = $false; Update = $false }
        $script:TailscaleClientElevatedTaskCacheInitialized = $true
        $script:TailscaleClientElevatedTaskCacheCheckedUtc = [datetime]::UtcNow
        return $false
    }
}

function Test-TailscaleClientElevatedTaskReady {
    param([ValidateSet('Check','Update')][string]$Operation,[switch]$Refresh)
    if ($Refresh) { return (Test-TailscaleClientElevatedTaskReadyDirect -Operation $Operation) }
    try {
        if (-not $script:TailscaleClientElevatedTaskCacheInitialized) { return $false }
        if ($null -eq $script:TailscaleClientElevatedTaskReadyCache) { return $false }
        return [bool]$script:TailscaleClientElevatedTaskReadyCache[$Operation]
    }
    catch { return $false }
}

function Get-TailscaleClientUpdateModeText {
    try {
        if (-not $script:TailscaleClientElevatedTaskCacheInitialized) { return 'Elevated task status pending' }
        $checkReady = Test-TailscaleClientElevatedTaskReady -Operation 'Check'
        $updateReady = Test-TailscaleClientElevatedTaskReady -Operation 'Update'
        if ($checkReady -and $updateReady) { return 'Elevated task ready' }
        return 'Elevated task setup required'
    }
    catch { return 'Elevated task unknown' }
}

function Test-TailscaleClientElevatedTasksReady {
    param([switch]$Refresh)
    if ($Refresh) { return (Refresh-TailscaleClientElevatedTaskCache) }
    try {
        if (-not $script:TailscaleClientElevatedTaskCacheInitialized) { return $false }
        return ((Test-TailscaleClientElevatedTaskReady -Operation 'Check') -and (Test-TailscaleClientElevatedTaskReady -Operation 'Update'))
    }
    catch { return $false }
}

function Start-TailscaleClientTaskReadinessRefresh {
    if ($script:IsClientTaskReadinessRefreshRunning) { return }
    $script:IsClientTaskReadinessRefreshRunning = $true
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 50
    $timer.add_Tick({
        try {
            $this.Stop()
            $this.Dispose()
            [void](Refresh-TailscaleClientElevatedTaskCache)
            Update-TailscaleClientTaskSetupUi
            if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
        }
        catch { Write-LogException -Context 'Refresh Tailscale client elevated task cache' -ErrorRecord $_ }
        finally { $script:IsClientTaskReadinessRefreshRunning = $false }
    })
    $timer.Start()
}

function Update-TailscaleClientTaskSetupUi {
    try {
        $ready = Test-TailscaleClientElevatedTasksReady
        if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
        if ($null -ne $script:btnInstallClientAutoUpdateTask) {
            $script:btnInstallClientAutoUpdateTask.Text = $(if ($ready) { 'Uninstall Auto Update Task' } else { 'Install Auto Update Task' })
            $script:btnInstallClientAutoUpdateTask.Enabled = (-not $script:IsClientTaskSetupRunning -and -not $script:IsClientMaintenanceTaskRunning)
        }
        if ($null -ne $script:btnCheckClientUpdate) {
            $script:btnCheckClientUpdate.Enabled = ($ready -and -not $script:IsClientTaskSetupRunning -and -not $script:IsClientMaintenanceTaskRunning)
        }
        if ($null -ne $script:chkCheckUpdateEvery) {
            $script:chkCheckUpdateEvery.Enabled = ($ready -and -not $script:IsClientTaskSetupRunning)
        }
        if ($null -ne $script:numCheckUpdateHours) {
            $script:numCheckUpdateHours.Enabled = ($ready -and -not $script:IsClientTaskSetupRunning -and $null -ne $script:chkCheckUpdateEvery -and [bool]$script:chkCheckUpdateEvery.Checked)
        }
        if ($null -ne $script:lblCheckUpdateHours) {
            $script:lblCheckUpdateHours.Enabled = ($ready -and -not $script:IsClientTaskSetupRunning -and $null -ne $script:chkCheckUpdateEvery -and [bool]$script:chkCheckUpdateEvery.Checked)
        }
    }
    catch { Write-LogException -Context 'Update Tailscale client task setup UI' -ErrorRecord $_ }
}

function Set-TailscaleClientTaskSetupBusyState {
    param([bool]$Busy,[string]$Operation = '')
    try {
        if ([string]::IsNullOrWhiteSpace([string]$Operation)) { $Operation = [string]$script:ClientTaskSetupOperation }
        $ready = Test-TailscaleClientElevatedTasksReady
        if ($null -ne $script:btnInstallClientAutoUpdateTask) {
            if ($Busy) {
                $script:btnInstallClientAutoUpdateTask.Text = $(if ($Operation -eq 'Uninstall') { 'Uninstalling...' } else { 'Installing...' })
            }
            else {
                $script:btnInstallClientAutoUpdateTask.Text = $(if ($ready) { 'Uninstall Auto Update Task' } else { 'Install Auto Update Task' })
            }
            $script:btnInstallClientAutoUpdateTask.Enabled = -not $Busy
        }
        if ($null -ne $script:btnCheckClientUpdate) { $script:btnCheckClientUpdate.Enabled = (-not $Busy -and $ready) }
        if ($null -ne $script:btnRunClientUpdate) { $script:btnRunClientUpdate.Enabled = $false }
        if ($null -ne $script:chkCheckUpdateEvery) { $script:chkCheckUpdateEvery.Enabled = (-not $Busy -and $ready) }
        if ($null -ne $script:numCheckUpdateHours) { $script:numCheckUpdateHours.Enabled = (-not $Busy -and $ready -and $null -ne $script:chkCheckUpdateEvery -and [bool]$script:chkCheckUpdateEvery.Checked) }
        if (-not $Busy) { $script:ClientTaskSetupOperation = '' }
    }
    catch { Write-LogException -Context 'Set Tailscale client task setup busy state' -ErrorRecord $_ }
}

function Get-TailscaleClientElevatedRunnerScript {
@'
param(
    [ValidateSet('Check','Update')][string]$Operation,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-ResultFile {
    param($Object)
    $json = $Object | ConvertTo-Json -Depth 14 -Compress
    $dir = Split-Path -Parent $ResultPath
    if (-not [string]::IsNullOrWhiteSpace([string]$dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllText([string]$ResultPath, $json, [Text.Encoding]::UTF8)
}

function Join-ProcessArgumentsLocal {
    param([string[]]$Arguments)
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($arg in @($Arguments)) {
        $value = [string]$arg
        if ($value -match '[\s"]') { $value = '"' + ($value -replace '"','\"') + '"' }
        [void]$items.Add($value)
    }
    return ($items -join ' ')
}

function Invoke-ExternalLocal {
    param([string]$FilePath,[string[]]$Arguments)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $psi = New-Object Diagnostics.ProcessStartInfo
        $psi.FileName = $FilePath
        $psi.Arguments = Join-ProcessArgumentsLocal -Arguments $Arguments
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        try { if ($null -ne $script:Utf8NoBomEncoding) { $psi.StandardOutputEncoding = $script:Utf8NoBomEncoding } else { $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding -ArgumentList $false } } catch { }
        try { if ($null -ne $script:Utf8NoBomEncoding) { $psi.StandardErrorEncoding = $script:Utf8NoBomEncoding } else { $psi.StandardErrorEncoding = New-Object System.Text.UTF8Encoding -ArgumentList $false } } catch { }
        $p = New-Object Diagnostics.Process
        $p.StartInfo = $psi
        [void]$p.Start()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()
        $sw.Stop()
        $parts = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrWhiteSpace($stdout)) { [void]$parts.Add($stdout.TrimEnd()) }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) { [void]$parts.Add($stderr.TrimEnd()) }
        return [pscustomobject]@{ ExitCode = [int]$p.ExitCode; Output = [string](($parts -join [Environment]::NewLine).TrimEnd()); DurationMs = [double]$sw.Elapsed.TotalMilliseconds }
    }
    catch {
        try { $sw.Stop() } catch { }
        return [pscustomobject]@{ ExitCode = 1; Output = [string]$_.Exception.Message; DurationMs = [double]$sw.Elapsed.TotalMilliseconds }
    }
}

function Find-TailscaleExeLocal {
    try {
        $cmd = Get-Command tailscale.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
        if (-not [string]::IsNullOrWhiteSpace([string]$cmd)) { return [string]$cmd }
    }
    catch { }
    foreach ($candidate in @((Join-Path $env:ProgramFiles 'Tailscale\tailscale.exe'), (Join-Path ${env:ProgramFiles(x86)} 'Tailscale\tailscale.exe'))) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate)) { return [string]$candidate }
    }
    return ''
}

function Get-PlainVersionLocal {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return (($Text -split "`r?`n") | Select-Object -First 1).Trim()
}

function Convert-ToComparableVersionLocal {
    param([string]$VersionText)
    $clean = [string]$VersionText
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    if ($clean -match '(\d+(?:\.\d+){1,3})') { $clean = $Matches[1] }
    try { return [version]$clean } catch { return $null }
}

function Test-VersionNewerLocal {
    param([string]$CurrentVersion,[string]$LatestVersion)
    $cur = Convert-ToComparableVersionLocal $CurrentVersion
    $lat = Convert-ToComparableVersionLocal $LatestVersion
    if ($null -eq $cur -or $null -eq $lat) { return $false }
    return ($lat -gt $cur)
}

function Get-TailscaleLatestStableVersionLocal {
    try {
        $resp = Invoke-WebRequest -Uri 'https://pkgs.tailscale.com/stable/' -UseBasicParsing -TimeoutSec 12
        $html = [string]$resp.Content
        $windowsVersions = New-Object System.Collections.Generic.List[string]
        foreach ($m in [regex]::Matches($html, 'tailscale-setup(?:-full)?-([0-9]+(?:\.[0-9]+){2,3})(?:-(?:amd64|arm64|x86))?\.(?:exe|msi)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $value = [string]$m.Groups[1].Value
            if (-not [string]::IsNullOrWhiteSpace($value) -and -not $windowsVersions.Contains($value)) { [void]$windowsVersions.Add($value) }
        }
        if ($windowsVersions.Count -gt 0) {
            $bestText = ''
            $bestVersion = $null
            foreach ($v in $windowsVersions) {
                $parsed = Convert-ToComparableVersionLocal $v
                if ($null -eq $parsed) { continue }
                if ($null -eq $bestVersion -or $parsed -gt $bestVersion) {
                    $bestVersion = $parsed
                    $bestText = [string]$v
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($bestText)) { return $bestText }
        }
        if ($html -match 'View older version:\s*latest\s+([0-9]+(?:\.[0-9]+){2,3})') { return [string]$Matches[1] }
        if ($html -match '>\s*latest\s+([0-9]+(?:\.[0-9]+){2,3})\s*<') { return [string]$Matches[1] }
    }
    catch { }
    return ''
}

function Get-HighestVersionFromTextLocal {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $versionMatches = [regex]::Matches([string]$Text, '(?<!\d)(\d+(?:\.\d+){2,3})(?!\d)')
    $bestText = ''
    $bestVersion = $null
    foreach ($m in $versionMatches) {
        $value = [string]$m.Groups[1].Value
        $parsed = Convert-ToComparableVersionLocal $value
        if ($null -eq $parsed) { continue }
        if ($null -eq $bestVersion -or $parsed -gt $bestVersion) {
            $bestVersion = $parsed
            $bestText = $value
        }
    }
    return [string]$bestText
}

function Get-UpdateCheckStatusFromDryRunLocal {
    param([string]$Output,[string]$CurrentVersion,[string]$LatestVersion)
    $lower = ([string]$Output).ToLowerInvariant()
    $hasExplicitUpToDate = ($lower -match 'already\s+up\s+to\s+date|up[-\s]?to[-\s]?date|no\s+updates?\s+(are\s+)?available|already\s+(running|using|on)')
    $hasExplicitAvailable = ($lower -match 'update\s+(is\s+)?available|new\s+version|would\s+update|can\s+be\s+updated|available\s+version')
    $hasUpdate = $false
    if (-not [string]::IsNullOrWhiteSpace($LatestVersion)) { $hasUpdate = Test-VersionNewerLocal -CurrentVersion $CurrentVersion -LatestVersion $LatestVersion }
    if ((-not $hasUpdate) -and (-not $hasExplicitUpToDate) -and (-not $hasExplicitAvailable)) { $LatestVersion = '' }
    if ($hasUpdate -or $hasExplicitAvailable) {
        return [pscustomobject]@{ Status = 'New version available.'; HasUpdate = $true; LatestVersion = [string]$LatestVersion }
    }
    if ($hasExplicitUpToDate) {
        if ([string]::IsNullOrWhiteSpace($LatestVersion) -and -not [string]::IsNullOrWhiteSpace($CurrentVersion)) { $LatestVersion = $CurrentVersion }
        return [pscustomobject]@{ Status = 'Already up to date.'; HasUpdate = $false; LatestVersion = [string]$LatestVersion }
    }
    if ([string]::IsNullOrWhiteSpace([string]$Output)) {
        return [pscustomobject]@{ Status = 'Check completed, but Tailscale did not return update details.'; HasUpdate = $false; LatestVersion = [string]$LatestVersion }
    }
    return [pscustomobject]@{ Status = 'Check completed. Review Tailscale output.'; HasUpdate = $false; LatestVersion = [string]$LatestVersion }
}

function Get-UpdateCheckLocal {
    param([string]$Exe)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $current = ''
    $versionResult = Invoke-ExternalLocal -FilePath $Exe -Arguments @('version')
    if ($versionResult.ExitCode -eq 0) { $current = Get-PlainVersionLocal ([string]$versionResult.Output) }
    $latestRemote = Get-TailscaleLatestStableVersionLocal
    $dryRun = Invoke-ExternalLocal -FilePath $Exe -Arguments @('update','--dry-run')
    if ($dryRun.ExitCode -ne 0) {
        $message = [string]$dryRun.Output
        if ([string]::IsNullOrWhiteSpace($message)) { $message = 'Tailscale update check failed.' }
        throw $message
    }
    $output = [string]$dryRun.Output
    $latestFromOutput = Get-HighestVersionFromTextLocal -Text $output
    $latest = if (-not [string]::IsNullOrWhiteSpace($latestFromOutput)) { $latestFromOutput } else { $latestRemote }
    $parsed = Get-UpdateCheckStatusFromDryRunLocal -Output $output -CurrentVersion $current -LatestVersion $latest
    $sw.Stop()
    return [pscustomobject]@{
        CurrentVersion = [string]$current
        LatestVersion = [string]$parsed.LatestVersion
        LastCheckUtc = ([datetime]::UtcNow).ToString('o')
        HasUpdate = [bool]$parsed.HasUpdate
        Status = [string]$parsed.Status
        DurationMs = [double]$sw.Elapsed.TotalMilliseconds
        Command = 'tailscale update --dry-run (scheduled task)'
        Output = [string]$output
        WasElevated = $true
    }
}

try {
    if ([string]::IsNullOrWhiteSpace([string]$ResultPath)) { throw 'Result path is missing.' }
    $exe = Find-TailscaleExeLocal
    if ([string]::IsNullOrWhiteSpace($exe)) { throw 'tailscale.exe was not detected.' }
    if ($Operation -eq 'Check') {
        $check = Get-UpdateCheckLocal -Exe $exe
        Write-ResultFile ([pscustomobject]@{ Operation = 'Check'; Success = $true; Check = $check; ErrorMessage = '' })
    }
    else {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $update = Invoke-ExternalLocal -FilePath $exe -Arguments @('update','--yes')
        $sw.Stop()
        if ($update.ExitCode -ne 0) { throw ([string]$update.Output) }
        $postCheck = Get-UpdateCheckLocal -Exe $exe
        $statusText = 'Update command completed at ' + (Get-Date).ToString('HH:mm:ss')
        Write-ResultFile ([pscustomobject]@{ Operation = 'Update'; Success = $true; UpdateResult = [pscustomobject]@{ StatusText = $statusText; DurationMs = [double]$update.DurationMs; Output = [string]$update.Output }; PostCheck = $postCheck; ErrorMessage = '' })
    }
}
catch {
    try { Write-ResultFile ([pscustomobject]@{ Operation = [string]$Operation; Success = $false; ErrorMessage = [string]$_.Exception.Message }) } catch { }
    exit 1
}
'@
}

function Get-TailscaleClientElevatedLauncherScript {
    $powerShell = ([string]$script:PowerShellExe).Replace('"','""')
    $runner = ([string]$script:TailscaleClientElevatedRunnerPath).Replace('"','""')
@"
Option Explicit
Dim shell
Dim operation
Dim resultPath
Dim commandLine

If WScript.Arguments.Count < 2 Then
    WScript.Quit 2
End If

operation = WScript.Arguments.Item(0)
resultPath = WScript.Arguments.Item(1)

commandLine = Chr(34) & "$powerShell" & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & "$runner" & Chr(34) & " -Operation " & operation & " -ResultPath " & Chr(34) & resultPath & Chr(34)

Set shell = CreateObject("WScript.Shell")
WScript.Quit shell.Run(commandLine, 0, True)
"@
}

function Install-TailscaleClientElevatedTasksDirect {
    if (-not (Test-IsAdministrator)) { throw 'Administrator privileges are required to create the elevated scheduled tasks.' }
    $runnerDir = Split-Path -Parent $script:TailscaleClientElevatedRunnerPath
    New-Item -ItemType Directory -Path $runnerDir -Force | Out-Null
    Set-Content -LiteralPath $script:TailscaleClientElevatedRunnerPath -Value (Get-TailscaleClientElevatedRunnerScript) -Encoding UTF8 -Force
    Set-Content -LiteralPath $script:TailscaleClientElevatedLauncherPath -Value (Get-TailscaleClientElevatedLauncherScript) -Encoding ASCII -Force
    $userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    try {
        $admins = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
        $system = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
        $adminName = $admins.Translate([System.Security.Principal.NTAccount]).Value
        $systemName = $system.Translate([System.Security.Principal.NTAccount]).Value
        & icacls $runnerDir /inheritance:r /grant:r "${adminName}:(OI)(CI)(F)" "${systemName}:(OI)(CI)(F)" "${userId}:(OI)(CI)(RX)" | Out-Null
    }
    catch { Write-LogException -Context 'Harden elevated task runner ACL' -ErrorRecord $_ }

    $userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    foreach ($operation in @('Check','Update')) {
        $taskName = Get-TailscaleClientElevatedTaskName -Operation $operation
        $resultPath = Get-TailscaleClientElevatedResultPath -Operation $operation
        $args = '//B //NoLogo "' + $script:TailscaleClientElevatedLauncherPath + '" ' + $operation + ' "' + $resultPath + '"'
        $action = New-ScheduledTaskAction -Execute $script:WScriptExe -Argument $args
        $principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 20)
        Register-ScheduledTask -TaskName $taskName -TaskPath $script:TailscaleClientTaskPath -Action $action -Principal $principal -Settings $settings -Description 'Elevated Tailscale client update task used by Tailscale Control.' -Force | Out-Null
    }
}

function Uninstall-TailscaleClientElevatedTasksDirect {
    foreach ($operation in @('Check','Update')) {
        try {
            $taskName = Get-TailscaleClientElevatedTaskName -Operation $operation
            $task = Get-ScheduledTask -TaskPath $script:TailscaleClientTaskPath -TaskName $taskName -ErrorAction SilentlyContinue
            if ($null -ne $task) { Unregister-ScheduledTask -TaskPath $script:TailscaleClientTaskPath -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue }
        }
        catch { }
    }
    foreach ($path in @($script:TailscaleClientElevatedRunnerPath,$script:TailscaleClientElevatedLauncherPath,$script:TailscaleClientCheckResultPath,$script:TailscaleClientUpdateResultPath)) {
        try { if (-not [string]::IsNullOrWhiteSpace([string]$path) -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } catch { }
    }
    try {
        $root = [string]$script:ProgramDataRoot
        if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-Path -LiteralPath $root)) {
            $children = @(Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue)
            if ($children.Count -eq 0) { Remove-Item -LiteralPath $root -Force -ErrorAction SilentlyContinue }
        }
    }
    catch { }
    try {
        $folderName = ([string]$script:TailscaleClientTaskPath).Trim('\')
        if (-not [string]::IsNullOrWhiteSpace($folderName)) {
            $service = New-Object -ComObject Schedule.Service
            $service.Connect()
            $folder = $service.GetFolder('\' + $folderName)
            if ($folder.GetTasks(0).Count -eq 0 -and $folder.GetFolders(0).Count -eq 0) {
                $rootFolder = $service.GetFolder('\')
                $rootFolder.DeleteFolder($folderName, 0)
            }
        }
    }
    catch { }
    [void](Refresh-TailscaleClientElevatedTaskCache)
}

function New-TailscaleClientElevatedTaskUninstallScript {
    param([string]$ResultPath)
    $payload = [pscustomobject]@{
        ProgramDataRoot = [string]$script:ProgramDataRoot
        RunnerPath = [string]$script:TailscaleClientElevatedRunnerPath
        LauncherPath = [string]$script:TailscaleClientElevatedLauncherPath
        TaskPath = [string]$script:TailscaleClientTaskPath
        CheckTaskName = [string]$script:TailscaleClientCheckTaskName
        UpdateTaskName = [string]$script:TailscaleClientUpdateTaskName
        CheckResultPath = [string]$script:TailscaleClientCheckResultPath
        UpdateResultPath = [string]$script:TailscaleClientUpdateResultPath
        ResultPath = [string]$ResultPath
    }
    $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes(($payload | ConvertTo-Json -Depth 8 -Compress)))
@"
`$ErrorActionPreference = 'Stop'
try {
    `$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
    `$payload = `$payloadJson | ConvertFrom-Json
    foreach (`$item in @(@{Name=[string]`$payload.CheckTaskName}, @{Name=[string]`$payload.UpdateTaskName})) {
        try {
            `$task = Get-ScheduledTask -TaskPath ([string]`$payload.TaskPath) -TaskName ([string]`$item.Name) -ErrorAction SilentlyContinue
            if (`$null -ne `$task) { Unregister-ScheduledTask -TaskPath ([string]`$payload.TaskPath) -TaskName ([string]`$item.Name) -Confirm:`$false -ErrorAction SilentlyContinue }
        }
        catch { }
    }
    foreach (`$path in @([string]`$payload.RunnerPath,[string]`$payload.LauncherPath,[string]`$payload.CheckResultPath,[string]`$payload.UpdateResultPath)) {
        try { if (-not [string]::IsNullOrWhiteSpace(`$path) -and (Test-Path -LiteralPath `$path)) { Remove-Item -LiteralPath `$path -Force -ErrorAction SilentlyContinue } } catch { }
    }
    try {
        `$root = [string]`$payload.ProgramDataRoot
        if (-not [string]::IsNullOrWhiteSpace(`$root) -and (Test-Path -LiteralPath `$root)) {
            `$children = @(Get-ChildItem -LiteralPath `$root -Force -ErrorAction SilentlyContinue)
            if (`$children.Count -eq 0) { Remove-Item -LiteralPath `$root -Force -ErrorAction SilentlyContinue }
        }
    } catch { }
    try {
        `$folderName = ([string]`$payload.TaskPath).Trim('\')
        if (-not [string]::IsNullOrWhiteSpace(`$folderName)) {
            `$service = New-Object -ComObject Schedule.Service
            `$service.Connect()
            `$folder = `$service.GetFolder('\' + `$folderName)
            if (`$folder.GetTasks(0).Count -eq 0 -and `$folder.GetFolders(0).Count -eq 0) {
                `$rootFolder = `$service.GetFolder('\')
                `$rootFolder.DeleteFolder(`$folderName, 0)
            }
        }
    } catch { }
    [IO.File]::WriteAllText([string]`$payload.ResultPath, (@{Success=`$true; ErrorMessage=''} | ConvertTo-Json -Compress), [Text.Encoding]::UTF8)
}
catch {
    try { [IO.File]::WriteAllText([string]`$payload.ResultPath, (@{Success=`$false; ErrorMessage=[string]`$_.Exception.Message} | ConvertTo-Json -Compress), [Text.Encoding]::UTF8) } catch { }
    exit 1
}
"@
}

function Start-TailscaleClientElevatedTaskUninstallProcess {
    param([string]$SetupPath,[string]$ResultPath)
    Set-Content -LiteralPath $SetupPath -Value (New-TailscaleClientElevatedTaskUninstallScript -ResultPath $ResultPath) -Encoding UTF8 -Force
    $args = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $SetupPath + '"'
    return (Start-Process -FilePath $script:PowerShellExe -ArgumentList $args -Verb RunAs -WindowStyle Hidden -PassThru -ErrorAction Stop)
}

function Start-TailscaleClientElevatedTaskUninstall {
    if ($script:IsClientTaskSetupRunning -or $script:IsClientMaintenanceTaskRunning) { return }
    $script:IsClientTaskSetupRunning = $true
    $script:ClientTaskSetupOperation = 'Uninstall'
    Set-TailscaleClientTaskSetupBusyState -Busy $true -Operation 'Uninstall'
    if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = 'Uninstalling elevated task...' }
    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Approve UAC once to uninstall the elevated update task.' }

    if (Test-IsAdministrator) {
        try {
            Uninstall-TailscaleClientElevatedTasksDirect
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Elevated update task uninstalled.' }
            Write-ActivityCommandBlock -Title 'Uninstall Auto Update Task' -CommandText 'Remove elevated scheduled tasks' -ExitCode 0 -Output 'Elevated Tailscale client update tasks were removed.'
            Show-Overlay -Title 'Auto update task uninstalled' -Message 'Tailscale client checks and updates will no longer run through elevated scheduled tasks.' -Indicator 'Off'
        }
        catch {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task uninstall failed: ' + $_.Exception.Message }
            Write-ActivityFailureBlock -Title 'Uninstall Auto Update Task failed' -CommandText 'Remove elevated scheduled tasks' -Message $_.Exception.Message
            Show-Overlay -Title 'Task uninstall failed' -Message $_.Exception.Message -ErrorStyle
        }
        finally {
            $script:IsClientTaskSetupRunning = $false
            $script:ClientTaskSetupWorker = $null
            Set-TailscaleClientTaskSetupBusyState -Busy $false
            Update-TailscaleClientTaskSetupUi
            if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
        }
        return
    }

    try {
        $id = [guid]::NewGuid().ToString('N')
        $resultPath = Join-Path $env:TEMP ('tailscale-control-task-uninstall-' + $id + '.json')
        $setupPath = Join-Path $env:TEMP ('tailscale-control-task-uninstall-' + $id + '.ps1')
        $proc = Start-TailscaleClientElevatedTaskUninstallProcess -SetupPath $setupPath -ResultPath $resultPath
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $script:ClientTaskSetupWorker = [pscustomobject]@{
            Process = $proc
            ResultPath = [string]$resultPath
            SetupPath = [string]$setupPath
            Timer = $timer
            StartedUtc = [datetime]::UtcNow
        }
        $timer.add_Tick({
            try {
                $taskRef = $script:ClientTaskSetupWorker
                if ($null -eq $taskRef) { return }
                $path = [string]$taskRef.ResultPath
                $timedOut = (([datetime]::UtcNow - ([datetime]$taskRef.StartedUtc)).TotalMinutes -gt 5)
                $processExited = $false
                try { if ($null -ne $taskRef.Process) { $processExited = [bool]$taskRef.Process.HasExited } } catch { }
                if (-not $timedOut -and -not (Test-Path -LiteralPath $path) -and -not $processExited) { return }
                try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
                try {
                    if ($timedOut -and -not (Test-Path -LiteralPath $path)) { throw 'Elevated task uninstall timed out before writing a result.' }
                    if (-not (Test-Path -LiteralPath $path)) { throw 'Elevated task uninstall was canceled or did not return a result.' }
                    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
                    $setupResult = $raw | ConvertFrom-Json
                    if (-not [bool](Get-ObjectPropertyOrDefault $setupResult 'Success' $false)) { throw ([string](Get-ObjectPropertyOrDefault $setupResult 'ErrorMessage' 'Elevated task uninstall failed.')) }
                    [void](Refresh-TailscaleClientElevatedTaskCache)
                    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Elevated update task uninstalled.' }
                    Write-ActivityCommandBlock -Title 'Uninstall Auto Update Task' -CommandText 'Remove elevated scheduled tasks' -ExitCode 0 -Output 'Elevated Tailscale client update tasks were removed.'
                    Show-Overlay -Title 'Auto update task uninstalled' -Message 'Tailscale client checks and updates will no longer run through elevated scheduled tasks.' -Indicator 'Off'
                }
                catch {
                    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task uninstall failed: ' + $_.Exception.Message }
                    Write-ActivityFailureBlock -Title 'Uninstall Auto Update Task failed' -CommandText 'Remove elevated scheduled tasks' -Message $_.Exception.Message
                    Show-Overlay -Title 'Task uninstall failed' -Message $_.Exception.Message -ErrorStyle
                }
                finally {
                    try { if (Test-Path -LiteralPath ([string]$taskRef.SetupPath)) { Remove-Item -LiteralPath ([string]$taskRef.SetupPath) -Force -ErrorAction SilentlyContinue } } catch { }
                    try { if (Test-Path -LiteralPath ([string]$taskRef.ResultPath)) { Remove-Item -LiteralPath ([string]$taskRef.ResultPath) -Force -ErrorAction SilentlyContinue } } catch { }
                    $script:IsClientTaskSetupRunning = $false
                    $script:ClientTaskSetupWorker = $null
                    Set-TailscaleClientTaskSetupBusyState -Busy $false
                    Update-TailscaleClientTaskSetupUi
                    if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
                }
            }
            catch {
                $script:IsClientTaskSetupRunning = $false
                $script:ClientTaskSetupWorker = $null
                Set-TailscaleClientTaskSetupBusyState -Busy $false
                if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task uninstall failed: ' + $_.Exception.Message }
                Write-ActivityFailureBlock -Title 'Uninstall Auto Update Task failed' -CommandText 'Remove elevated scheduled tasks' -Message $_.Exception.Message
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsClientTaskSetupRunning = $false
        $script:ClientTaskSetupWorker = $null
        Set-TailscaleClientTaskSetupBusyState -Busy $false
        Update-TailscaleClientTaskSetupUi
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task uninstall failed: ' + $_.Exception.Message }
        Write-ActivityFailureBlock -Title 'Uninstall Auto Update Task failed' -CommandText 'Remove elevated scheduled tasks' -Message $_.Exception.Message
        Show-Overlay -Title 'Task uninstall failed' -Message $_.Exception.Message -ErrorStyle
    }
}

function Invoke-TailscaleClientAutoUpdateTaskButton {
    try {
        if (Test-TailscaleClientElevatedTasksReady -Refresh) { Start-TailscaleClientElevatedTaskUninstall }
        else { Start-TailscaleClientElevatedTaskSetup }
    }
    catch {
        Write-ActivityFailureBlock -Title 'Auto Update Task action failed' -CommandText 'Toggle elevated scheduled tasks' -Message $_.Exception.Message
        Show-Overlay -Title 'Task action failed' -Message $_.Exception.Message -ErrorStyle
    }
}

function Start-TailscaleClientElevatedTaskSetup {
    if ($script:IsClientTaskSetupRunning -or $script:IsClientMaintenanceTaskRunning) { return }
    $script:IsClientTaskSetupRunning = $true
    $script:ClientTaskSetupOperation = 'Install'
    Set-TailscaleClientTaskSetupBusyState -Busy $true -Operation 'Install'
    if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = 'Installing elevated task...' }
    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Approve UAC once to install the elevated update task.' }

    if (Test-IsAdministrator) {
        try {
            Install-TailscaleClientElevatedTasksDirect
            if (-not (Test-TailscaleClientElevatedTasksReady -Refresh)) { throw 'Elevated scheduled tasks could not be verified after setup.' }
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Elevated update task installed.' }
            Write-ActivityCommandBlock -Title 'Install Auto Update Task' -CommandText 'Register elevated scheduled tasks' -ExitCode 0 -Output 'Elevated Tailscale client update tasks were installed.'
            Show-Overlay -Title 'Auto update task installed' -Message 'Tailscale client checks and updates can now run without repeated UAC prompts.' -Indicator 'On'
        }
        catch {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task setup failed: ' + $_.Exception.Message }
            Write-ActivityFailureBlock -Title 'Install Auto Update Task failed' -CommandText 'Register elevated scheduled tasks' -Message $_.Exception.Message
            Show-Overlay -Title 'Task setup failed' -Message $_.Exception.Message -ErrorStyle
        }
        finally {
            $script:IsClientTaskSetupRunning = $false
            $script:ClientTaskSetupWorker = $null
            Set-TailscaleClientTaskSetupBusyState -Busy $false
            Update-TailscaleClientTaskSetupUi
            if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
        }
        return
    }

    try {
        $id = [guid]::NewGuid().ToString('N')
        $resultPath = Join-Path $env:TEMP ('tailscale-control-task-setup-' + $id + '.json')
        $setupPath = Join-Path $env:TEMP ('tailscale-control-task-setup-' + $id + '.ps1')
        $setupLauncherPath = Join-Path $env:TEMP ('tailscale-control-task-setup-' + $id + '.vbs')
        $payload = [pscustomobject]@{
            ProgramDataRoot = [string]$script:ProgramDataRoot
            RunnerPath = [string]$script:TailscaleClientElevatedRunnerPath
            RunnerContentBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((Get-TailscaleClientElevatedRunnerScript)))
            LauncherPath = [string]$script:TailscaleClientElevatedLauncherPath
            LauncherContentBase64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes((Get-TailscaleClientElevatedLauncherScript)))
            TaskPath = [string]$script:TailscaleClientTaskPath
            CheckTaskName = [string]$script:TailscaleClientCheckTaskName
            UpdateTaskName = [string]$script:TailscaleClientUpdateTaskName
            CheckResultPath = [string]$script:TailscaleClientCheckResultPath
            UpdateResultPath = [string]$script:TailscaleClientUpdateResultPath
            PowerShellExe = [string]$script:PowerShellExe
            WScriptExe = [string]$script:WScriptExe
            UserId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ResultPath = [string]$resultPath
        }
        $payloadEncoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes(($payload | ConvertTo-Json -Depth 8 -Compress)))
        $setup = @"
`$ErrorActionPreference = 'Stop'
try {
    `$payloadJson = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$payloadEncoded'))
    `$payload = `$payloadJson | ConvertFrom-Json
    New-Item -ItemType Directory -Path ([string]`$payload.ProgramDataRoot) -Force | Out-Null
    `$runnerContent = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]`$payload.RunnerContentBase64))
    [IO.File]::WriteAllText([string]`$payload.RunnerPath, `$runnerContent, [Text.Encoding]::UTF8)
    `$launcherContent = [Text.Encoding]::ASCII.GetString([Convert]::FromBase64String([string]`$payload.LauncherContentBase64))
    [IO.File]::WriteAllText([string]`$payload.LauncherPath, `$launcherContent, [Text.Encoding]::ASCII)
    try {
        `$admins = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, `$null)
        `$system = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, `$null)
        `$adminName = `$admins.Translate([System.Security.Principal.NTAccount]).Value
        `$systemName = `$system.Translate([System.Security.Principal.NTAccount]).Value
        `$userName = [string]`$payload.UserId
        & icacls ([string]`$payload.ProgramDataRoot) /inheritance:r /grant:r "`${adminName}:(OI)(CI)(F)" "`${systemName}:(OI)(CI)(F)" "`${userName}:(OI)(CI)(RX)" | Out-Null
    } catch { }
    foreach (`$item in @(@{Name=[string]`$payload.CheckTaskName; Operation='Check'; Result=[string]`$payload.CheckResultPath}, @{Name=[string]`$payload.UpdateTaskName; Operation='Update'; Result=[string]`$payload.UpdateResultPath})) {
        `$args = '//B //NoLogo "' + [string]`$payload.LauncherPath + '" ' + [string]`$item.Operation + ' "' + [string]`$item.Result + '"'
        `$action = New-ScheduledTaskAction -Execute ([string]`$payload.WScriptExe) -Argument `$args
        `$principal = New-ScheduledTaskPrincipal -UserId ([string]`$payload.UserId) -LogonType Interactive -RunLevel Highest
        `$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 20)
        Register-ScheduledTask -TaskName ([string]`$item.Name) -TaskPath ([string]`$payload.TaskPath) -Action `$action -Principal `$principal -Settings `$settings -Description 'Elevated Tailscale client update task used by Tailscale Control.' -Force | Out-Null
    }
    [IO.File]::WriteAllText([string]`$payload.ResultPath, (@{Success=`$true; ErrorMessage=''} | ConvertTo-Json -Compress), [Text.Encoding]::UTF8)
}
catch {
    try { [IO.File]::WriteAllText([string]`$payload.ResultPath, (@{Success=`$false; ErrorMessage=[string]`$_.Exception.Message} | ConvertTo-Json -Compress), [Text.Encoding]::UTF8) } catch { }
    exit 1
}
"@
        Set-Content -LiteralPath $setupPath -Value $setup -Encoding UTF8 -Force
        $ps = ([string]$script:PowerShellExe).Replace('"','""')
        $sp = ([string]$setupPath).Replace('"','""')
        $vbs = @"
Option Explicit
Dim shell
Dim commandLine
commandLine = Chr(34) & "$ps" & Chr(34) & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & "$sp" & Chr(34)
Set shell = CreateObject("WScript.Shell")
WScript.Quit shell.Run(commandLine, 0, True)
"@
        Set-Content -LiteralPath $setupLauncherPath -Value $vbs -Encoding ASCII -Force
        $proc = Start-Process -FilePath $script:WScriptExe -ArgumentList ('"' + $setupLauncherPath + '"') -Verb RunAs -PassThru -ErrorAction Stop
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 250
        $script:ClientTaskSetupWorker = [pscustomobject]@{
            Process = $proc
            ResultPath = [string]$resultPath
            SetupPath = [string]$setupPath
            LauncherPath = [string]$setupLauncherPath
            Timer = $timer
            StartedUtc = [datetime]::UtcNow
        }
        $timer.add_Tick({
            try {
                $taskRef = $script:ClientTaskSetupWorker
                if ($null -eq $taskRef) { return }
                $path = [string]$taskRef.ResultPath
                $timedOut = (([datetime]::UtcNow - ([datetime]$taskRef.StartedUtc)).TotalMinutes -gt 5)
                if (-not $timedOut -and -not (Test-Path -LiteralPath $path)) { return }
                try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
                try {
                    if ($timedOut -and -not (Test-Path -LiteralPath $path)) { throw 'Elevated task setup timed out before writing a result.' }
                    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
                    $setupResult = $raw | ConvertFrom-Json
                    if (-not [bool](Get-ObjectPropertyOrDefault $setupResult 'Success' $false)) { throw ([string](Get-ObjectPropertyOrDefault $setupResult 'ErrorMessage' 'Elevated task setup failed.')) }
                    if (-not (Test-TailscaleClientElevatedTasksReady -Refresh)) { throw 'Elevated scheduled tasks could not be verified after setup.' }
                    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Elevated update task installed.' }
                    Write-ActivityCommandBlock -Title 'Install Auto Update Task' -CommandText 'Register elevated scheduled tasks' -ExitCode 0 -Output 'Elevated Tailscale client update tasks were installed.'
                    Show-Overlay -Title 'Auto update task installed' -Message 'Tailscale client checks and updates can now run without repeated UAC prompts.' -Indicator 'On'
                }
                catch {
                    if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task setup failed: ' + $_.Exception.Message }
                    Write-ActivityFailureBlock -Title 'Install Auto Update Task failed' -CommandText 'Register elevated scheduled tasks' -Message $_.Exception.Message
                    Show-Overlay -Title 'Task setup failed' -Message $_.Exception.Message -ErrorStyle
                }
                finally {
                    try { if (Test-Path -LiteralPath ([string]$taskRef.SetupPath)) { Remove-Item -LiteralPath ([string]$taskRef.SetupPath) -Force -ErrorAction SilentlyContinue } } catch { }
                    try { if (Test-Path -LiteralPath ([string]$taskRef.LauncherPath)) { Remove-Item -LiteralPath ([string]$taskRef.LauncherPath) -Force -ErrorAction SilentlyContinue } } catch { }
                    try { if (Test-Path -LiteralPath ([string]$taskRef.ResultPath)) { Remove-Item -LiteralPath ([string]$taskRef.ResultPath) -Force -ErrorAction SilentlyContinue } } catch { }
                    $script:IsClientTaskSetupRunning = $false
                    $script:ClientTaskSetupWorker = $null
                    Set-TailscaleClientTaskSetupBusyState -Busy $false
                    Update-TailscaleClientTaskSetupUi
                    if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
                }
            }
            catch {
                $script:IsClientTaskSetupRunning = $false
                $script:ClientTaskSetupWorker = $null
                Set-TailscaleClientTaskSetupBusyState -Busy $false
                if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task setup failed: ' + $_.Exception.Message }
                Write-ActivityFailureBlock -Title 'Install Auto Update Task failed' -CommandText 'Register elevated scheduled tasks' -Message $_.Exception.Message
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsClientTaskSetupRunning = $false
        $script:ClientTaskSetupWorker = $null
        Set-TailscaleClientTaskSetupBusyState -Busy $false
        Update-TailscaleClientTaskSetupUi
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task setup failed: ' + $_.Exception.Message }
        Write-ActivityFailureBlock -Title 'Install Auto Update Task failed' -CommandText 'Register elevated scheduled tasks' -Message $_.Exception.Message
        Show-Overlay -Title 'Task setup failed' -Message $_.Exception.Message -ErrorStyle
    }
}

function Complete-TailscaleClientMaintenanceResult {
    param($Result,[string]$FallbackOperation)
    $operationName = [string](Get-ObjectPropertyOrDefault $Result 'Operation' $FallbackOperation)
    $success = [bool](Get-ObjectPropertyOrDefault $Result 'Success' $false)
    if (-not $success) {
        $message = [string](Get-ObjectPropertyOrDefault $Result 'ErrorMessage' 'Task failed.')
        if ($operationName -eq 'Update') {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Update failed: ' + $message }
            Write-ActivityFailureBlock -Title 'Tailscale Update failed' -CommandText 'tailscale update' -Message $message
        }
        else {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Check failed: ' + $message }
            Write-ActivityFailureBlock -Title 'Tailscale Check Update failed' -CommandText 'tailscale update --dry-run' -Message $message
        }
        return
    }

    if ($operationName -eq 'Check') {
        $check = Get-ObjectPropertyOrDefault $Result 'Check' $null
        if ($null -ne $check) {
            $latest = [string](Get-ObjectPropertyOrDefault $check 'LatestVersion' '')
            $lastCheck = [string](Get-ObjectPropertyOrDefault $check 'LastCheckUtc' '')
            $parsedDate = [datetime]::MinValue
            if ([datetime]::TryParse($lastCheck, [ref]$parsedDate)) { Set-LastClientUpdateCheckTimestamp -Timestamp $parsedDate.ToUniversalTime() }
            else { Set-LastClientUpdateCheckTimestamp }
            Set-LastClientUpdateLatestVersion -Version $latest
            Set-LastClientUpdateStatus -Status ([string](Get-ObjectPropertyOrDefault $check 'Status' ''))
            Write-VersionCheckActivity -Title 'Tailscale Check Update' -CommandText ([string](Get-ObjectPropertyOrDefault $check 'Command' 'tailscale update --dry-run')) -Check $check
        }
        if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
    }
    else {
        $updateResult = Get-ObjectPropertyOrDefault $Result 'UpdateResult' $null
        $statusText = if ($null -ne $updateResult) { [string](Get-ObjectPropertyOrDefault $updateResult 'StatusText' 'Update command completed.') } else { 'Update command completed.' }
        $duration = if ($null -ne $updateResult) { [double](Get-ObjectPropertyOrDefault $updateResult 'DurationMs' 0) } else { 0 }
        Write-ActivityCommandBlock -Title 'Tailscale Update' -CommandText 'tailscale update' -ExitCode 0 -Output [string]$statusText -DurationMs $duration
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = [string]$statusText }
        $postCheck = Get-ObjectPropertyOrDefault $Result 'PostCheck' $null
        if ($null -ne $postCheck) {
            $latest = [string](Get-ObjectPropertyOrDefault $postCheck 'LatestVersion' '')
            $lastCheck = [string](Get-ObjectPropertyOrDefault $postCheck 'LastCheckUtc' '')
            $parsedDate = [datetime]::MinValue
            if ([datetime]::TryParse($lastCheck, [ref]$parsedDate)) { Set-LastClientUpdateCheckTimestamp -Timestamp $parsedDate.ToUniversalTime() }
            else { Set-LastClientUpdateCheckTimestamp }
            Set-LastClientUpdateLatestVersion -Version $latest
            Set-LastClientUpdateStatus -Status ([string](Get-ObjectPropertyOrDefault $postCheck 'Status' ''))
            Write-VersionCheckActivity -Title 'Tailscale Check Update' -CommandText ([string](Get-ObjectPropertyOrDefault $postCheck 'Command' 'tailscale update --dry-run')) -Check $postCheck
        }
        if ($null -ne $script:Snapshot) { Update-TailscaleClientUpdateUi -Snapshot $script:Snapshot }
    }
}

function Start-TailscaleClientMaintenanceTask {
    param(
        [ValidateSet('Check','Update','AutoUpdate')][string]$Operation,
        [switch]$Scheduled
    )
    if ($script:IsClientMaintenanceTaskRunning) { return }
    $taskOperation = if ($Operation -eq 'AutoUpdate') { 'Update' } else { [string]$Operation }
    if (-not (Test-TailscaleClientElevatedTasksReady)) {
        if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Install the auto update task first.' }
        if (-not $Scheduled) { Show-Overlay -Title 'Task not installed' -Message 'Install the Auto Update Task first so checks and updates can run without repeated UAC prompts.' -Indicator 'Warn' }
        return
    }
    $script:IsClientMaintenanceTaskRunning = $true
    Set-TailscaleClientMaintenanceBusyState -Busy $true -Operation $taskOperation
    try {
        if ($taskOperation -eq 'Check') {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Checking for updates...' }
        }
        else {
            if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = $(if ($Scheduled) { 'Running scheduled update...' } else { 'Updating...' }) }
        }
        if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
        $resultPath = Get-TailscaleClientElevatedResultPath -Operation $taskOperation
        try { if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue } } catch { }
        $taskName = Get-TailscaleClientElevatedTaskName -Operation $taskOperation
        Start-ScheduledTask -TaskPath $script:TailscaleClientTaskPath -TaskName $taskName -ErrorAction Stop
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 500
        $task = [pscustomobject]@{
            Operation = [string]$taskOperation
            Scheduled = [bool]$Scheduled
            ResultPath = [string]$resultPath
            TaskName = [string]$taskName
            Timer = $timer
            StartedUtc = [datetime]::UtcNow
        }
        $script:ClientMaintenanceWorker = $task
        $timer.add_Tick({
            try {
                $taskRef = $script:ClientMaintenanceWorker
                if ($null -eq $taskRef) { return }
                $path = [string]$taskRef.ResultPath
                $timedOut = (([datetime]::UtcNow - ([datetime]$taskRef.StartedUtc)).TotalMinutes -gt 25)
                if (-not $timedOut -and -not (Test-Path -LiteralPath $path)) { return }
                try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
                $operationName = [string]$taskRef.Operation
                try {
                    if ($timedOut -and -not (Test-Path -LiteralPath $path)) { throw 'Elevated scheduled task timed out before writing a result.' }
                    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
                    $result = $raw | ConvertFrom-Json
                    Complete-TailscaleClientMaintenanceResult -Result $result -FallbackOperation $operationName
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if ($operationName -eq 'Update') {
                        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Update failed: ' + $message }
                        Write-ActivityFailureBlock -Title 'Tailscale Update failed' -CommandText 'tailscale update' -Message $message
                    }
                    else {
                        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Check failed: ' + $message }
                        Write-ActivityFailureBlock -Title 'Tailscale Check Update failed' -CommandText 'tailscale update --dry-run' -Message $message
                    }
                }
                finally {
                    try { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } catch { }
                    $script:IsClientMaintenanceTaskRunning = $false
                    $script:ClientMaintenanceWorker = $null
                    Set-TailscaleClientMaintenanceBusyState -Busy $false -Operation $operationName
                    if ($null -ne $script:btnCheckClientUpdate) {
                        $script:btnCheckClientUpdate.Enabled = $true
                        $script:btnCheckClientUpdate.Text = 'Check Update'
                    }
                    if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
                    Update-TailscaleClientTaskSetupUi
                }
            }
            catch {
                $operationName = 'Check'
                try { if ($null -ne $script:ClientMaintenanceWorker) { $operationName = [string]$script:ClientMaintenanceWorker.Operation } } catch { }
                $script:IsClientMaintenanceTaskRunning = $false
                $script:ClientMaintenanceWorker = $null
                Set-TailscaleClientMaintenanceBusyState -Busy $false -Operation $operationName
                if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task failed: ' + $_.Exception.Message }
                Write-ActivityFailureBlock -Title ('Tailscale ' + $operationName + ' failed') -CommandText ('tailscale ' + $(if ($operationName -eq 'Update') { 'update' } else { 'check update' })) -Message $_.Exception.Message
            }
        })
        $timer.Start()
    }
    catch {
        $script:IsClientMaintenanceTaskRunning = $false
        $script:ClientMaintenanceWorker = $null
        Set-TailscaleClientMaintenanceBusyState -Busy $false -Operation $taskOperation
        if ($null -ne $script:lblMaintAutoUpdate) { $script:lblMaintAutoUpdate.Text = Get-TailscaleClientUpdateModeText }
        Update-TailscaleClientTaskSetupUi
        if ($null -ne $script:lblMaintUpdateStatus) { $script:lblMaintUpdateStatus.Text = 'Task failed: ' + $_.Exception.Message }
        Write-ActivityFailureBlock -Title ('Tailscale ' + $taskOperation + ' failed') -CommandText ('tailscale ' + $(if ($taskOperation -eq 'Update') { 'update' } else { 'check update' })) -Message $_.Exception.Message
    }
}

function Set-MainActionButtonsBusyState {
    param([bool]$Busy,$Button = $null,[string]$BusyText = 'Working...')
    try {
        $buttons = @($script:btnRefresh,$script:btnToggleConnect,$script:btnToggleExit,$script:btnToggleDns,$script:btnToggleSubnets,$script:btnToggleIncoming)
        foreach ($btn in @($buttons)) {
            if ($null -eq $btn) { continue }
            if ($Busy) {
                if ($btn -eq $Button) {
                    if ($null -eq $btn.Tag -or -not $btn.Tag.PSObject.Properties['OriginalText']) { $btn.Tag = [pscustomobject]@{ OriginalText = [string]$btn.Text } }
                    $btn.Text = $BusyText
                    $btn.Enabled = $true
                    Set-ControlFocusSafe -Control $btn
                }
                else { $btn.Enabled = $false }
            }
            else {
                try { if ($null -ne $btn.Tag -and $btn.Tag.PSObject.Properties['OriginalText']) { $btn.Text = [string]$btn.Tag.OriginalText; $btn.Tag = $null } } catch { }
                $btn.Enabled = $true
            }
        }
        if (-not $Busy) {
            try { Update-ExitNodeActionAvailability -Snapshot $script:Snapshot } catch { }
            try { Update-TrayMenuState -Snapshot $script:Snapshot } catch { }
            try { Update-SelectedDeviceActionButtons } catch { }
        }
    }
    catch { Write-LogException -Context 'Set main action busy state' -ErrorRecord $_ }
}

function Complete-AsyncTailscaleActionTask {
    $taskRef = $script:AsyncActionTask
    if ($null -eq $taskRef) { return }
    try { $taskRef.Timer.Stop(); $taskRef.Timer.Dispose() } catch { }
    try {
        $proc = $taskRef.Process
        $stdout = ''
        $stderr = ''
        try { $stdout = [string]$proc.StandardOutput.ReadToEnd() } catch { }
        try { $stderr = [string]$proc.StandardError.ReadToEnd() } catch { }
        $output = (($stdout,$stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join [Environment]::NewLine
        $exit = 1
        try { $exit = [int]$proc.ExitCode } catch { }
        try { $proc.Dispose() } catch { }
        $elapsed = ([datetime]::UtcNow - [datetime]$taskRef.StartedUtc).TotalMilliseconds
        if ($exit -ne 0) { throw $output }
        $isSwitchAccountTask = ([string]$taskRef.Kind -eq 'SwitchAccount')
        if ($isSwitchAccountTask -and -not [string]::IsNullOrWhiteSpace([string]$output)) {
            $displayForOutput = ConvertTo-DiagnosticText -Text ([string]$script:PendingSwitchAccountDisplayIdentifier)
            $pendingSwitchId = ConvertTo-DiagnosticText -Text ([string]$script:PendingSwitchAccountIdentifier)
            if ([string]::IsNullOrWhiteSpace($displayForOutput) -or [string]::Equals($displayForOutput, $pendingSwitchId, [System.StringComparison]::OrdinalIgnoreCase)) {
                try {
                    $candidates = @(Get-QuickAccountSwitchLoggedAccounts -Force)
                    $matchedCandidate = $null
                    foreach ($candidate in $candidates) {
                        if ($null -eq $candidate) { continue }
                        $candidateSwitch = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $candidate 'SwitchIdentifier' ''))
                        if (-not [string]::IsNullOrWhiteSpace($candidateSwitch) -and [string]::Equals($candidateSwitch, $pendingSwitchId, [System.StringComparison]::OrdinalIgnoreCase)) { $matchedCandidate = $candidate; break }
                    }
                    if ($null -eq $matchedCandidate) {
                        foreach ($candidate in $candidates) {
                            if ($null -eq $candidate) { continue }
                            $candidateIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $candidate 'Identifier' ''))
                            if (-not [string]::IsNullOrWhiteSpace($candidateIdentifier) -and [string]::Equals($candidateIdentifier, $pendingSwitchId, [System.StringComparison]::OrdinalIgnoreCase)) { $matchedCandidate = $candidate; break }
                        }
                    }
                    if ($null -eq $matchedCandidate) {
                        foreach ($candidate in $candidates) {
                            if ($null -eq $candidate) { continue }
                            if ([bool](Get-ObjectPropertyOrDefault $candidate 'Active' $false)) { $matchedCandidate = $candidate; break }
                        }
                    }
                    if ($null -ne $matchedCandidate) {
                        $resolvedDisplay = Get-TailscaleAccountSwitchOutputIdentifier -Account $matchedCandidate
                        if (-not [string]::IsNullOrWhiteSpace($resolvedDisplay)) { $displayForOutput = $resolvedDisplay }
                    }
                } catch { }
            }
            if (-not [string]::IsNullOrWhiteSpace($displayForOutput)) {
                $outputLines = New-Object System.Collections.Generic.List[string]
                foreach ($line in ([string]$output -split "`r?`n")) {
                    if ([string]$line -match '^Switching to account\s+".*?"\s*$') { [void]$outputLines.Add('Switching to account "' + $displayForOutput + '"') }
                    else { [void]$outputLines.Add([string]$line) }
                }
                $output = ($outputLines -join [Environment]::NewLine)
            }
        }
        $activitySummary = [string]$taskRef.SuccessMessage
        if ($isSwitchAccountTask -and [string]::IsNullOrWhiteSpace($activitySummary)) { $activitySummary = 'Tailscale account switch completed.' }
        $activityOutput = Get-CommandActivityOutput -Summary $activitySummary -Output $output
        Write-ActivityCommandBlock -Title ([string]$taskRef.SuccessTitle) -CommandText ([string]$taskRef.CommandText) -ExitCode $exit -Output $activityOutput -DurationMs ([double]$elapsed)
        if (-not [bool]$taskRef.SuppressOverlay -and -not $isSwitchAccountTask) { Show-ToggleOverlay -Title ([string]$taskRef.SuccessTitle) -Message ([string]$taskRef.SuccessMessage) -Indicator ([string]$taskRef.Indicator) }
        if ([string]$taskRef.Kind -match 'Toggle|DNS|Subnets|Incoming|Exit|Connect') { Invoke-ToggleFeedbackSound -Enabled:([bool]([string]$taskRef.Indicator -eq 'On')) }
        try { Reset-SlowSnapshotCache } catch { }
        $refreshAccountsAfterAction = ([string]$taskRef.Kind -eq 'SwitchAccount' -or [string]$taskRef.Kind -eq 'AddAccount' -or [string]$taskRef.Kind -eq 'LogoutAccount')
        try { Update-Status -RefreshAccounts:$refreshAccountsAfterAction } catch { Write-LogException -Context 'Refresh after async action' -ErrorRecord $_ }
        if ($isSwitchAccountTask) {
            $overlayIdentifier = ConvertTo-DiagnosticText -Text ([string]$script:PendingSwitchAccountDisplayIdentifier)
            if ([string]::IsNullOrWhiteSpace($overlayIdentifier)) {
                try {
                    $activeAccount = @(Get-TailscaleSwitchAccounts | Where-Object { [bool]$_.Active } | Select-Object -First 1)
                    if ($activeAccount.Count -gt 0) { $overlayIdentifier = Get-TailscaleAccountOverlayIdentifier -Account $activeAccount[0] }
                } catch { }
            }
            if ([string]::IsNullOrWhiteSpace($overlayIdentifier)) { $overlayIdentifier = 'selected account' }
            if (-not [bool]$taskRef.SuppressOverlay) {
                Start-SwitchAccountCompletionOverlay -Identifier $overlayIdentifier -PreviousTailnet ([string]$script:PendingSwitchPreviousTailnet) -Title ([string]$taskRef.SuccessTitle) -Indicator ([string]$taskRef.Indicator)
            }
            else {
                $script:PendingSwitchAccountIdentifier = ''
                $script:PendingSwitchAccountDisplayIdentifier = ''
                $script:PendingSwitchPreviousTailnet = ''
                try { Update-LoggedAccountsView -Force $true } catch { }
            }
        }
    }
    catch {
        $message = [string]$_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($message)) { $message = [string]$_ }
        if ([string]$taskRef.Kind -eq 'SwitchAccount' -or [string]$taskRef.Kind -eq 'AddAccount' -or [string]$taskRef.Kind -eq 'LogoutAccount') {
            if ([string]$taskRef.Kind -eq 'SwitchAccount') {
                $script:PendingSwitchAccountIdentifier = ''
                $script:PendingSwitchAccountDisplayIdentifier = ''
                $script:PendingSwitchPreviousTailnet = ''
            }
            try { Update-LoggedAccountsView -Force $true } catch { }
        }
        Write-ActivityFailureBlock -Title (([string]$taskRef.Title) + ' failed') -CommandText ([string]$taskRef.CommandText) -Message $message
        Show-Overlay -Title (([string]$taskRef.Title) + ' failed') -Message $message -ErrorStyle
    }
    finally {
        $script:IsAsyncActionRunning = $false
        $script:AsyncActionTask = $null
        Set-MainActionButtonsBusyState -Busy $false
    }
}

function Start-TailscaleActionProcessAsync {
    param(
        [string]$Title,
        [string[]]$Arguments,
        [string]$SuccessTitle,
        [string]$SuccessMessage,
        [string]$Indicator = 'Info',
        $Button = $null,
        [string]$BusyText = 'Working...',
        [string]$Kind = 'Action'
    )
    if ($script:IsAsyncActionRunning) { return }
    $exe = ''
    try { if ($null -ne $script:Snapshot -and -not [string]::IsNullOrWhiteSpace([string]$script:Snapshot.Exe)) { $exe = [string]$script:Snapshot.Exe } } catch { }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { $exe = Find-TailscaleExe }
    if ([string]::IsNullOrWhiteSpace([string]$exe)) { throw 'tailscale.exe not found.' }
    $suppressOverlayForTask = [bool]$script:SuppressNextAsyncToggleOverlay
    $script:SuppressNextAsyncToggleOverlay = $false
    $started = Start-TailscaleProcess -Exe $exe -Arguments $Arguments
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250
    $script:IsAsyncActionRunning = $true
    $script:LastActionStartedAsync = $true
    $script:AsyncActionTask = [pscustomobject]@{
        Title = [string]$Title
        SuccessTitle = [string]$SuccessTitle
        SuccessMessage = [string]$SuccessMessage
        Indicator = [string]$Indicator
        Kind = [string]$Kind
        Process = $started.Process
        Arguments = $started.Arguments
        CommandText = 'tailscale ' + (($Arguments | ForEach-Object { [string]$_ }) -join ' ')
        Timer = $timer
        StartedUtc = [datetime]$started.StartedUtc
        Button = $Button
        SuppressOverlay = [bool]$suppressOverlayForTask
    }
    Set-MainActionButtonsBusyState -Busy $true -Button $Button -BusyText $BusyText
    $timer.add_Tick({
        try {
            $taskRef = $script:AsyncActionTask
            if ($null -eq $taskRef) { return }
            if ($null -ne $taskRef.Process -and -not $taskRef.Process.HasExited) { return }
            Complete-AsyncTailscaleActionTask
        }
        catch {
            Write-ActivityFailureBlock -Title 'Async action failed' -CommandText 'tailscale action' -Message $_.Exception.Message
            $script:IsAsyncActionRunning = $false
            $script:AsyncActionTask = $null
            Set-MainActionButtonsBusyState -Busy $false
        }
    })
    $timer.Start()
}

function Invoke-Action {
    param([string]$Title,[scriptblock]$Action)
    if ($script:IsBusy -or $script:IsAsyncActionRunning) { return }
    $script:IsBusy = $true
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $restoreBounds = $null
    try {
        if ($null -ne $script:MainForm -and $script:MainForm.WindowState -eq 'Normal') {
            $restoreBounds = $script:MainForm.Bounds
        }
    }
    catch { }
    try {
        $script:LastActionStartedAsync = $false
        & $Action
        if (-not $script:LastActionStartedAsync) {
            Reset-SlowSnapshotCache
            Update-Status
        }
    }
    catch {
        try { $sw.Stop() } catch { }
        Write-Log ($Title + ' failed: ' + $_.Exception.Message)
        Write-ActivityFailureBlock -Title ($Title + ' failed') -CommandText $Title -Message $_.Exception.Message -DurationMs ([double]$sw.Elapsed.TotalMilliseconds)
        Show-Overlay -Title ($Title + ' failed') -Message $_.Exception.Message -ErrorStyle
    }
    finally {
        try {
            if ($null -ne $restoreBounds -and $null -ne $script:MainForm -and $script:MainForm.IsHandleCreated -and $script:MainForm.WindowState -eq 'Normal') {
                $boundsToRestore = $restoreBounds
                $null = $script:MainForm.BeginInvoke([Action]{
                    try {
                        if ($null -ne $script:MainForm -and $script:MainForm.WindowState -eq 'Normal') {
                            $script:MainForm.Bounds = $boundsToRestore
                        }
                    }
                    catch { }
                })
            }
        }
        catch { }
        $script:IsBusy = $false
    }
}

function Set-TailscaleSetting {
    param([string[]]$SettingArgs,[string]$SuccessTitle,[string]$SuccessMessage,[string]$Indicator='Info',$Button = $null)
    $snapshot = Get-CurrentSnapshot
    if (-not $snapshot.Found) { throw 'tailscale.exe was not detected.' }
    Start-TailscaleActionProcessAsync -Title $SuccessTitle -Arguments (@('set') + $SettingArgs) -SuccessTitle $SuccessTitle -SuccessMessage $SuccessMessage -Indicator $Indicator -Button $Button -BusyText 'Updating...' -Kind 'Toggle'
}

function Switch-Connect {
    param($Button = $null)
    $snapshot = Get-CurrentSnapshot
    if (-not $snapshot.Found) { throw 'tailscale.exe was not detected.' }
    if ($snapshot.BackendState -eq 'Running') {
        Start-TailscaleActionProcessAsync -Title 'Toggle Connect' -Arguments @('down') -SuccessTitle 'Tailscale disconnected' -SuccessMessage 'The local device was disconnected from Tailscale.' -Indicator 'Off' -Button $Button -BusyText 'Disconnecting...' -Kind 'ToggleConnect'
    }
    else {
        Start-TailscaleActionProcessAsync -Title 'Toggle Connect' -Arguments @('up') -SuccessTitle 'Tailscale connected' -SuccessMessage 'The local device is connected.' -Indicator 'On' -Button $Button -BusyText 'Connecting...' -Kind 'ToggleConnect'
    }
}

function Switch-Dns {
    param($Button = $null)
    $snapshot = Get-CurrentSnapshot
    $nextBool = Convert-ToNullableBool $snapshot.CorpDNS
    if ($null -eq $nextBool) { $nextBool = $false }
    $next = -not $nextBool
    Set-TailscaleSetting -SettingArgs @('--accept-dns=' + $next.ToString().ToLowerInvariant()) -SuccessTitle 'DNS updated' -SuccessMessage ('Tailscale DNS is now ' + ($(if ($next) { 'On' } else { 'Off' })) + '.') -Indicator $(if ($next) { 'On' } else { 'Off' }) -Button $Button
}

function Switch-Subnets {
    param($Button = $null)
    $snapshot = Get-CurrentSnapshot
    $nextBool = Convert-ToNullableBool $snapshot.RouteAll
    if ($null -eq $nextBool) { $nextBool = $false }
    $next = -not $nextBool
    Set-TailscaleSetting -SettingArgs @('--accept-routes=' + $next.ToString().ToLowerInvariant()) -SuccessTitle 'Subnets updated' -SuccessMessage ('Subnet routes are now ' + ($(if ($next) { 'On' } else { 'Off' })) + '.') -Indicator $(if ($next) { 'On' } else { 'Off' }) -Button $Button
}

function Switch-Incoming {
    param($Button = $null)
    $snapshot = Get-CurrentSnapshot
    $allowedValue = Convert-ToNullableBool $snapshot.IncomingAllowed
    $allowed = if ($null -eq $allowedValue) { $true } else { $allowedValue }
    $nextAllowed = -not $allowed
    Set-TailscaleSetting -SettingArgs @('--shields-up=' + ((-not $nextAllowed).ToString().ToLowerInvariant())) -SuccessTitle ('Incoming ' + ($(if ($nextAllowed) { 'allowed' } else { 'blocked' }))) -SuccessMessage ('Incoming connections are now ' + ($(if ($nextAllowed) { 'Allowed' } else { 'Blocked' })) + '.') -Indicator $(if ($nextAllowed) { 'On' } else { 'Off' }) -Button $Button
}

function Switch-ExitNode {
    param($Button = $null)
    $snapshot = Get-CurrentSnapshot
    if (-not $snapshot.Found) { throw 'tailscale.exe was not detected.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$snapshot.CurrentExitNode)) {
        Start-TailscaleActionProcessAsync -Title 'Toggle Exit Node' -Arguments @('set','--exit-node=') -SuccessTitle 'Exit node cleared' -SuccessMessage 'Traffic is no longer using an exit node.' -Indicator 'Off' -Button $Button -BusyText 'Updating...' -Kind 'ToggleExitNode'
        return
    }
    $exitNodeAvailability = Get-ExitNodeToggleAvailability -Snapshot $snapshot
    if (-not [bool]$exitNodeAvailability.Enabled) {
        Write-Log ([string]$exitNodeAvailability.Reason)
        Show-Overlay -Title 'No exit node available' -Message ([string]$exitNodeAvailability.Reason) -Indicator 'Warn'
        return
    }
    $cfg = Get-Config
    $selectedLabel = if ($null -ne $script:cmbExitNode.SelectedItem) { ConvertTo-DnsName ([string]$script:cmbExitNode.SelectedItem) } else { Get-PreferredExitNodeLabel }
    $targetNode = $null
    foreach ($node in @(Convert-ToObjectArray $snapshot.ExitNodes)) {
        $nodeLabel = ConvertTo-DnsName ([string]$node.DNSName)
        $nodeName = ConvertTo-DnsName ([string]$node.Name)
        if ($selectedLabel -eq $nodeLabel -or $selectedLabel -eq $nodeName) { $targetNode = $node; break }
    }
    if ($null -eq $targetNode -and (Convert-ToObjectArray $snapshot.ExitNodes).Count -gt 0) { $targetNode = (Convert-ToObjectArray $snapshot.ExitNodes)[0] }
    if ($null -eq $targetNode) {
        Show-Overlay -Title 'No exit node available' -Message 'This tailnet does not currently expose any exit nodes.' -Indicator 'Warn'
        return
    }
    $target = ConvertTo-DnsName ([string]$targetNode.IPv4)
    if ([string]::IsNullOrWhiteSpace($target)) {
        foreach ($m in @(Convert-ToObjectArray $snapshot.Machines)) {
            if ($null -eq $m) { continue }
            if ($selectedLabel -eq (ConvertTo-DnsName ([string]$m.DNSName)) -or $selectedLabel -eq (ConvertTo-DnsName ([string]$m.Machine))) {
                $target = ConvertTo-DnsName ([string]$m.IPv4)
                if (-not [string]::IsNullOrWhiteSpace($target)) { break }
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($target)) { $target = ConvertTo-DnsName ([string]$targetNode.Name) }
    if ([string]::IsNullOrWhiteSpace($target)) { throw 'No preferred exit node is available.' }
    $tsArgs = @('set', ('--exit-node=' + $target), ('--exit-node-allow-lan-access=' + ($(if ([bool]$cfg.allow_lan_on_exit) { 'true' } else { 'false' }))))
    Start-TailscaleActionProcessAsync -Title 'Toggle Exit Node' -Arguments $tsArgs -SuccessTitle 'Exit node enabled' -SuccessMessage ('Using ' + $selectedLabel + ' as the exit node.') -Indicator 'On' -Button $Button -BusyText 'Updating...' -Kind 'ToggleExitNode'
}

function Send-ExistingInstanceSignal {
    try {
        $evt = [System.Threading.EventWaitHandle]::OpenExisting('Luiz.TailscaleControl.Activate')
        try { [void]$evt.Set() } finally { $evt.Dispose() }
    }
    catch { }
}

function Start-InstanceActivationMonitor {
    if ($null -eq $script:InstanceActivateEvent) { return }
    if ($null -ne $script:InstanceActivateTimer) { return }
    $script:InstanceActivateTimer = New-Object System.Windows.Forms.Timer
    $script:InstanceActivateTimer.Interval = 250
    $script:InstanceActivateTimer.add_Tick({
        try {
            if ($null -ne $script:InstanceActivateEvent -and $script:InstanceActivateEvent.WaitOne(0)) {
                Show-MainForm
            }
        }
        catch { Write-LogException -Context 'Instance activation monitor tick' -ErrorRecord $_ }
    })
    $script:InstanceActivateTimer.Start()
}

function Set-MainFormCenteredOnPrimaryScreen {
    if ($null -eq $script:MainForm -or $script:MainForm.IsDisposed) { return }
    try {
        $primary = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $bounds = $script:MainForm.Bounds
        $isOffscreen = ($bounds.Right -lt ($primary.Left + 80) -or $bounds.Bottom -lt ($primary.Top + 80) -or $bounds.Left -gt ($primary.Right - 80) -or $bounds.Top -gt ($primary.Bottom - 80))
        $wa = if ($script:MainForm.Visible -and -not $isOffscreen) { [System.Windows.Forms.Screen]::FromControl($script:MainForm).WorkingArea } else { $primary }
        if ($null -eq $wa -or $wa.Width -le 0 -or $wa.Height -le 0) { $wa = $primary }
        $targetX = [Math]::Max($wa.Left, [int]([math]::Floor($wa.Left + (($wa.Width - $script:MainForm.Width) / 2))))
        $targetY = [Math]::Max($wa.Top, [int]([math]::Floor($wa.Top + (($wa.Height - $script:MainForm.Height) / 2))))
        $script:MainForm.StartPosition = 'Manual'
        $script:MainForm.Location = New-Object System.Drawing.Point($targetX, $targetY)
    }
    catch { }
}

function Show-MainForm {
    if ($null -eq $script:MainForm) { return }
    if ($script:MainForm.InvokeRequired) {
        $null = $script:MainForm.BeginInvoke([Action]{ Show-MainForm })
        return
    }
    $script:StartupHidePending = $false
    $script:MainFormVisibilityToken = [int]$script:MainFormVisibilityToken + 1
    $wasVisible = $false
    try { $wasVisible = [bool]$script:MainForm.Visible -and $script:MainForm.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized -and [double]$script:MainForm.Opacity -gt 0.05 } catch { }
    if (-not $wasVisible) { try { $script:MainForm.Opacity = 0.0 } catch { } }
    try { Set-MainFormAppIcon } catch { }
    try { $script:MainForm.ShowInTaskbar = $true } catch { }
    try { Set-MainFormAppIcon } catch { }
    try { $script:MainForm.WindowState = 'Normal' } catch { }
    if (-not $wasVisible) { try { Set-MainFormCenteredOnPrimaryScreen } catch { } }
    try {
        $bounds = $script:MainForm.Bounds
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $isOffscreen = ($bounds.Right -lt ($wa.Left + 80) -or $bounds.Bottom -lt ($wa.Top + 80) -or $bounds.Left -gt ($wa.Right - 80) -or $bounds.Top -gt ($wa.Bottom - 80))
        if (-not $script:MainFormPresentedOnce -or $isOffscreen) {
            Set-MainFormCenteredOnPrimaryScreen
            $script:MainFormPresentedOnce = $true
        }
    } catch { }
    try { $script:MainForm.Show() } catch { }
    try { $script:MainForm.Visible = $true } catch { }
    if (-not $wasVisible) { try { Set-MainFormCenteredOnPrimaryScreen } catch { } }
    try { $script:MainForm.Opacity = 1.0 } catch { }
    try { Set-MainFormAppIcon } catch { }
    try { Set-BodySplitPreferredLayout } catch { }
    try {
        if ($script:MainForm.IsHandleCreated) {
            $null = $script:MainForm.BeginInvoke([Action]{ try { Set-BodySplitPreferredLayout } catch { } })
        }
    } catch { }
    try {
        $bounds = $script:MainForm.Bounds
        $wa = [System.Windows.Forms.Screen]::FromControl($script:MainForm).WorkingArea
        if ($null -eq $wa -or $wa.Width -le 0 -or $wa.Height -le 0) { $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea }
        $isOffscreen = ($bounds.Right -lt ($wa.Left + 80) -or $bounds.Bottom -lt ($wa.Top + 80) -or $bounds.Left -gt ($wa.Right - 80) -or $bounds.Top -gt ($wa.Bottom - 80))
        if ($isOffscreen) {
            Set-MainFormCenteredOnPrimaryScreen
            $script:MainFormPresentedOnce = $true
        }
    } catch { }
    $script:MainForm.TopMost = $true
    $script:MainForm.Activate()
    $script:MainForm.BringToFront()
    $script:MainForm.Focus() | Out-Null
    $script:MainForm.TopMost = $false
}

function Invoke-HotkeyAction {
    param([string]$Name)
    if ($script:IsCapturingHotkey) { return }
    if ($script:HotkeyExecutionLock) { return }
    $now = Get-Date
    if ($script:LastHotkeyAt.ContainsKey($Name)) {
        $delta = ($now - [datetime]$script:LastHotkeyAt[$Name]).TotalMilliseconds
        $minimumDelay = if ($Name -eq 'ShowSettings') { 70 } else { 220 }
        if ($delta -lt $minimumDelay) { return }
    }
    $script:LastHotkeyAt[$Name] = $now
    Start-HotkeyReleaseMonitor -Name $Name
    try {
        switch ($Name) {
            'ToggleConnect' { Invoke-Action 'Toggle Connect' { Switch-Connect } }
            'ToggleExitNode' { Invoke-Action 'Toggle Exit Node' { Switch-ExitNode } }
            'ToggleDns' { Invoke-Action 'Toggle DNS' { Switch-Dns } }
            'ToggleSubnets' { Invoke-Action 'Toggle Subnets' { Switch-Subnets } }
            'ToggleIncoming' { Invoke-Action 'Toggle Incoming' { Switch-Incoming } }
            'ShowSettings' { Switch-MainFormVisibility }
            default { if ((Get-QuickAccountSwitchIndex -Name $Name) -gt 0) { Invoke-QuickAccountSwitchHotkey -Name $Name } }
        }
    }
    catch {
        Clear-HotkeyExecutionLock
        throw
    }
}

function Invoke-TrayToggleAction {
    param([string]$Title,[scriptblock]$Action)
    $previousOverlay = $script:SuppressToggleOverlay
    $previousAsyncOverlay = $script:SuppressNextAsyncToggleOverlay
    try {
        $script:SuppressToggleOverlay = $true
        $script:SuppressNextAsyncToggleOverlay = $true
        Invoke-Action $Title $Action
        if (-not $script:LastActionStartedAsync) { $script:SuppressNextAsyncToggleOverlay = $previousAsyncOverlay }
    }
    finally {
        $script:SuppressToggleOverlay = $previousOverlay
    }
}

function New-ActionButton {
    param([string]$Text,[int]$Left,[int]$Top,[int]$Width)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($Left,$Top)
    $btn.Size = New-Object System.Drawing.Size($Width,34)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = 'Standard'
    $btn.UseVisualStyleBackColor = $true
    return $btn
}

function New-FlowButtonRow {
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Dock = 'Top'
    $panel.AutoSize = $true
    $panel.FlowDirection = 'LeftToRight'
    $panel.WrapContents = $false
    $panel.AutoScroll = $false
    $panel.Padding = New-Object System.Windows.Forms.Padding(0)
    $panel.Margin = New-Object System.Windows.Forms.Padding(0)
    $panel.Anchor = 'Left'
    $panel.MinimumSize = New-Object System.Drawing.Size(0, 40)
    return $panel
}

function New-HotkeyRow {
    param([System.Windows.Forms.Control]$Parent,[string]$Name,[int]$Top,[string]$LabelText)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16,$Top)
    $lbl.Size = New-Object System.Drawing.Size(158,24)
    $lbl.Text = $LabelText
    $lbl.Font = $script:SummaryLabelFont
    $Parent.Controls.Add($lbl) | Out-Null

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Location = New-Object System.Drawing.Point(188,$Top)
    $chk.Size = New-Object System.Drawing.Size(78,24)
    $chk.Text = 'Enable'
    $Parent.Controls.Add($chk) | Out-Null

    $cmbMod = New-Object System.Windows.Forms.ComboBox
    $cmbMod.DropDownStyle = 'DropDownList'
    $cmbMod.Location = New-Object System.Drawing.Point(272,$Top)
    $cmbMod.Size = New-Object System.Drawing.Size(1,1)
    $cmbMod.Visible = $false
    [void]$cmbMod.Items.AddRange(@('Ctrl+Alt','Ctrl+Shift','Alt+Shift','Ctrl+Alt+Shift','Alt','Ctrl','Shift','None'))
    $Parent.Controls.Add($cmbMod) | Out-Null

    $cmbKey = New-Object System.Windows.Forms.ComboBox
    $cmbKey.DropDownStyle = 'DropDownList'
    $cmbKey.Location = New-Object System.Drawing.Point(272,$Top)
    $cmbKey.Size = New-Object System.Drawing.Size(1,1)
    $cmbKey.Visible = $false
    [void]$cmbKey.Items.AddRange(@('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','F1','F2','F3','F4','F5','F6','F7','F8','F9','F10','F11','F12'))
    $Parent.Controls.Add($cmbKey) | Out-Null

    $txtHotkey = New-Object System.Windows.Forms.TextBox
    $txtHotkey.Location = New-Object System.Drawing.Point(272,$Top)
    $txtHotkey.Size = New-Object System.Drawing.Size(148,24)
    $txtHotkey.ReadOnly = $true
    $txtHotkey.Text = ''
    $txtHotkey.Tag = [pscustomobject]@{ Modifiers = $cmbMod; Key = $cmbKey }
    $txtHotkey.add_Enter({
        param($controlSender,$e)
        $script:IsCapturingHotkey = $true
        try { Unregister-Hotkeys | Out-Null } catch { Write-LogException -Context 'Unregister hotkeys during uninstall' -ErrorRecord $_ }
        $controlSender.SelectAll()
    })
    $txtHotkey.add_Leave({
        param($controlSender,$e)
        if ($script:IsCapturingHotkey) {
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 350
            $timer.add_Tick({
                param($ts,$te)
                try { $ts.Stop() } catch { Write-LogException -Context 'Stop hotkey re-register timer' -ErrorRecord $_ }
                try { $ts.Dispose() } catch { Write-LogException -Context 'Dispose hotkey re-register timer' -ErrorRecord $_ }
                $script:IsCapturingHotkey = $false
                try { [void](Register-Hotkeys -Config (Get-Config)) } catch { Write-LogException -Context 'Re-register hotkeys after capture delay' -ErrorRecord $_ }
            })
            $timer.Start()
        }
        else {
            try { [void](Register-Hotkeys -Config (Get-Config)) } catch { Write-LogException -Context 'Re-register hotkeys after capture cancel' -ErrorRecord $_ }
        }
    })
    $txtHotkey.add_KeyDown({
        param($controlSender,$e)
        $meta = $controlSender.Tag
        $cmbModRef = $meta.Modifiers
        $cmbKeyRef = $meta.Key
        if ($e.KeyCode -in @([System.Windows.Forms.Keys]::ControlKey,[System.Windows.Forms.Keys]::ShiftKey,[System.Windows.Forms.Keys]::Menu)) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            return
        }
        if ($e.KeyCode -in @([System.Windows.Forms.Keys]::Back,[System.Windows.Forms.Keys]::Delete,[System.Windows.Forms.Keys]::Escape)) {
            $cmbModRef.SelectedItem = 'None'
            $cmbKeyRef.SelectedItem = $null
            $cmbKeyRef.Text = ''
            $controlSender.Text = ''
            try { Save-UiSettings -Silent } catch { }
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $controlSender.Parent.Focus()
            return
        }
        $mods = Get-ModifierTextFromKeyEvent -keyEventArgs $e
        $keyText = Get-KeyTextFromKeyCode -KeyCode $e.KeyCode
        if (-not [string]::IsNullOrWhiteSpace($keyText)) {
            if ($cmbModRef.Items.Contains($mods)) { $cmbModRef.SelectedItem = $mods } else { $cmbModRef.SelectedItem = 'None' }
            if ($cmbKeyRef.Items.Contains($keyText)) { $cmbKeyRef.SelectedItem = $keyText } else { $cmbKeyRef.Text = $keyText }
            if ([string]::IsNullOrWhiteSpace($mods) -or $mods -eq 'None') { $controlSender.Text = $keyText } else { $controlSender.Text = $mods + '+' + $keyText }
            try { Save-UiSettings -Silent } catch { }
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $controlSender.Parent.Focus()
        }
    })
    $Parent.Controls.Add($txtHotkey) | Out-Null

    $script:HotkeyControls[$Name] = [pscustomobject]@{ Label = $lbl; Enabled = $chk; Modifiers = $cmbMod; Key = $cmbKey; Capture = $txtHotkey; Hint = $null }
}

function Get-QuickAccountSwitchDisplay {
    param($Account,[int]$Index)
    if ($null -eq $Account) { return '' }
    $emailPattern = '[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}'
    $quickDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickDisplay' ''))
    if (-not [string]::IsNullOrWhiteSpace($quickDisplay)) { return $quickDisplay }
    $tailnetEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickTailnetEmail' ''))
    $userEmail = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'QuickUserEmail' ''))
    if ([string]::IsNullOrWhiteSpace($tailnetEmail)) {
        $raw = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'Raw' ''))
        $matches = @([regex]::Matches($raw, $emailPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))
        if ($matches.Count -gt 1) {
            $tailnetEmail = [string]$matches[0].Value
            if ([string]::IsNullOrWhiteSpace($userEmail)) { $userEmail = [string]$matches[$matches.Count - 1].Value }
        }
    }
    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        $userValue = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Account 'User' ''))
        if ($userValue -match $emailPattern) { $userEmail = [string]$Matches[0] }
    }
    if ($tailnetEmail -match $emailPattern) { $tailnetEmail = [string]$Matches[0] } else { $tailnetEmail = '' }
    if ($userEmail -match $emailPattern) { $userEmail = [string]$Matches[0] } else { $userEmail = '' }
    if (-not [string]::IsNullOrWhiteSpace($tailnetEmail) -and -not [string]::IsNullOrWhiteSpace($userEmail)) { return ($tailnetEmail + ' | ' + $userEmail) }
    if (-not [string]::IsNullOrWhiteSpace($tailnetEmail) -and [string]::IsNullOrWhiteSpace($userEmail)) { return ($tailnetEmail + ' | ' + $tailnetEmail) }
    if (-not [string]::IsNullOrWhiteSpace($userEmail)) { return ($userEmail + ' | ' + $userEmail) }
    return ('Account ' + [string]$Index)
}

function New-QuickAccountSwitchRow {
    param([System.Windows.Forms.Control]$Parent,[string]$Name,[int]$Top,[string]$LabelText)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16,$Top)
    $lbl.Size = New-Object System.Drawing.Size(82,24)
    $lbl.Text = $LabelText
    $lbl.TextAlign = 'MiddleLeft'
    $lbl.Font = $script:SummaryLabelFont
    $Parent.Controls.Add($lbl) | Out-Null

    $cmbAccount = New-Object System.Windows.Forms.ComboBox
    $cmbAccount.DropDownStyle = 'DropDownList'
    $cmbAccount.Location = New-Object System.Drawing.Point(104,$Top)
    $cmbAccount.Size = New-Object System.Drawing.Size(238,24)
    $cmbAccount.DropDownWidth = 500
    $cmbAccount.DisplayMember = 'Display'
    $Parent.Controls.Add($cmbAccount) | Out-Null

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Location = New-Object System.Drawing.Point(354,$Top)
    $chk.Size = New-Object System.Drawing.Size(76,24)
    $chk.Text = 'Enable'
    $Parent.Controls.Add($chk) | Out-Null

    $cmbMod = New-Object System.Windows.Forms.ComboBox
    $cmbMod.DropDownStyle = 'DropDownList'
    $cmbMod.Location = New-Object System.Drawing.Point(436,$Top)
    $cmbMod.Size = New-Object System.Drawing.Size(1,1)
    $cmbMod.Visible = $false
    [void]$cmbMod.Items.AddRange(@('Ctrl+Alt','Ctrl+Shift','Alt+Shift','Ctrl+Alt+Shift','Alt','Ctrl','Shift','None'))
    $Parent.Controls.Add($cmbMod) | Out-Null

    $cmbKey = New-Object System.Windows.Forms.ComboBox
    $cmbKey.DropDownStyle = 'DropDownList'
    $cmbKey.Location = New-Object System.Drawing.Point(436,$Top)
    $cmbKey.Size = New-Object System.Drawing.Size(1,1)
    $cmbKey.Visible = $false
    [void]$cmbKey.Items.AddRange(@('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','F1','F2','F3','F4','F5','F6','F7','F8','F9','F10','F11','F12'))
    $Parent.Controls.Add($cmbKey) | Out-Null

    $txtHotkey = New-Object System.Windows.Forms.TextBox
    $txtHotkey.Location = New-Object System.Drawing.Point(436,$Top)
    $txtHotkey.Size = New-Object System.Drawing.Size(166,24)
    $txtHotkey.ReadOnly = $true
    $txtHotkey.Text = ''
    $txtHotkey.Tag = [pscustomobject]@{ Modifiers = $cmbMod; Key = $cmbKey }
    $txtHotkey.add_Enter({
        param($controlSender,$e)
        $script:IsCapturingHotkey = $true
        try { Unregister-Hotkeys | Out-Null } catch { Write-LogException -Context 'Unregister hotkeys during uninstall' -ErrorRecord $_ }
        $controlSender.SelectAll()
    })
    $txtHotkey.add_Leave({
        param($controlSender,$e)
        if ($script:IsCapturingHotkey) {
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 350
            $timer.add_Tick({
                param($ts,$te)
                try { $ts.Stop() } catch { Write-LogException -Context 'Stop hotkey re-register timer' -ErrorRecord $_ }
                try { $ts.Dispose() } catch { Write-LogException -Context 'Dispose hotkey re-register timer' -ErrorRecord $_ }
                $script:IsCapturingHotkey = $false
                try { [void](Register-Hotkeys -Config (Get-Config)) } catch { Write-LogException -Context 'Re-register hotkeys after capture delay' -ErrorRecord $_ }
            })
            $timer.Start()
        }
        else {
            try { [void](Register-Hotkeys -Config (Get-Config)) } catch { Write-LogException -Context 'Re-register hotkeys after capture cancel' -ErrorRecord $_ }
        }
    })
    $txtHotkey.add_KeyDown({
        param($controlSender,$e)
        $meta = $controlSender.Tag
        $cmbModRef = $meta.Modifiers
        $cmbKeyRef = $meta.Key
        if ($e.KeyCode -in @([System.Windows.Forms.Keys]::ControlKey,[System.Windows.Forms.Keys]::ShiftKey,[System.Windows.Forms.Keys]::Menu)) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            return
        }
        if ($e.KeyCode -in @([System.Windows.Forms.Keys]::Back,[System.Windows.Forms.Keys]::Delete,[System.Windows.Forms.Keys]::Escape)) {
            $cmbModRef.SelectedItem = 'None'
            $cmbKeyRef.SelectedItem = $null
            $cmbKeyRef.Text = ''
            $controlSender.Text = ''
            try { Save-UiSettings -Silent } catch { }
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $controlSender.Parent.Focus()
            return
        }
        $mods = Get-ModifierTextFromKeyEvent -keyEventArgs $e
        $keyText = Get-KeyTextFromKeyCode -KeyCode $e.KeyCode
        if (-not [string]::IsNullOrWhiteSpace($keyText)) {
            if ($cmbModRef.Items.Contains($mods)) { $cmbModRef.SelectedItem = $mods } else { $cmbModRef.SelectedItem = 'None' }
            if ($cmbKeyRef.Items.Contains($keyText)) { $cmbKeyRef.SelectedItem = $keyText } else { $cmbKeyRef.Text = $keyText }
            if ([string]::IsNullOrWhiteSpace($mods) -or $mods -eq 'None') { $controlSender.Text = $keyText } else { $controlSender.Text = $mods + '+' + $keyText }
            try { Save-UiSettings -Silent } catch { }
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            $controlSender.Parent.Focus()
        }
    })
    $Parent.Controls.Add($txtHotkey) | Out-Null

    $cmbAccount.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } })
    $chk.add_CheckedChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } })
    $cmbMod.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } })
    $cmbKey.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } })
    Set-AppToolTip -Control $cmbAccount -Text ''
    Set-AppToolTip -Control $chk -Text 'Default: Off'
    Set-AppToolTip -Control $txtHotkey -Text 'Default: Ctrl+Alt+Shift+1'
    $script:HotkeyControls[$Name] = [pscustomobject]@{ Label = $lbl; Enabled = $chk; Modifiers = $cmbMod; Key = $cmbKey; Capture = $txtHotkey; Account = $cmbAccount; Hint = $null }
    $script:QuickAccountSwitchRows[$Name] = [pscustomobject]@{ Label = $lbl; Account = $cmbAccount; Enabled = $chk; Modifiers = $cmbMod; Key = $cmbKey; Capture = $txtHotkey }
}

function New-QuickAccountSwitchSection {
    param([System.Windows.Forms.Control]$Parent,[int]$Top)
    $separator = New-Object System.Windows.Forms.Label
    $separator.Location = New-Object System.Drawing.Point(16,$Top)
    $separatorWidth = [Math]::Max(300, $Parent.ClientSize.Width - 32)
    $separator.Size = New-Object System.Drawing.Size -ArgumentList $separatorWidth,1
    $separator.Anchor = 'Top,Left,Right'
    $separator.BorderStyle = 'Fixed3D'
    $Parent.Controls.Add($separator) | Out-Null
    $script:QuickAccountSwitchSeparator = $separator
    $script:QuickAccountSwitchHeader = $null
    $script:QuickAccountSwitchSectionPanel = $Parent
    try {
        $Parent.add_SizeChanged({
            try {
                if ($null -ne $script:QuickAccountSwitchSeparator -and $null -ne $script:QuickAccountSwitchSectionPanel) {
                    $script:QuickAccountSwitchSeparator.Width = [Math]::Max(300, $script:QuickAccountSwitchSectionPanel.ClientSize.Width - 32)
                }
            } catch { }
        })
    } catch { }
    Update-QuickAccountSwitchRows -Accounts @()
}

function Update-QuickAccountSwitchRows {
    param($Accounts)
    if ($null -eq $script:QuickAccountSwitchSectionPanel) { return }
    $accountList = @($Accounts)
    $quickSwitchAvailable = ([int]$accountList.Count -gt 1)
    $script:QuickAccountSwitchAvailable = [bool]$quickSwitchAvailable
    $cfg = Get-Config
    $count = [Math]::Max([int]$script:QuickAccountSwitchMinimumRows, [int]$accountList.Count)
    if ($count -gt [int]$script:QuickAccountSwitchMaximumRows) { $count = [int]$script:QuickAccountSwitchMaximumRows }
    Ensure-QuickAccountSwitchDefinitions -Count $count
    Ensure-ConfigHotkeyEntries -Config $cfg -Count $count
    $previousLoading = $script:IsLoadingConfig
    $script:IsLoadingConfig = $true
    try {
        for ($i = 1; $i -le $count; $i++) {
            $name = Get-QuickAccountSwitchName -Index $i
            if (-not $script:QuickAccountSwitchRows.ContainsKey($name)) {
                New-QuickAccountSwitchRow -Parent $script:QuickAccountSwitchSectionPanel -Name $name -Top (320 + (($i - 1) * 40)) -LabelText ('Account ' + [string]$i)
            }
            $row = $script:QuickAccountSwitchRows[$name]
            foreach ($control in @($row.Label,$row.Account,$row.Enabled,$row.Capture)) { if ($null -ne $control) { $control.Visible = $true } }
            foreach ($control in @($row.Modifiers,$row.Key)) { if ($null -ne $control) { $control.Visible = $false } }

            $combo = $row.Account
            $combo.Items.Clear()
            $entry = $cfg.hotkeys.PSObject.Properties[$name].Value
            $selectedIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $entry 'account_identifier' ''))
            $selectedDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $entry 'account_display' ''))
            $defaultIdentifier = ''
            for ($a = 0; $a -lt $accountList.Count; $a++) {
                $account = $accountList[$a]
                $switchIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'SwitchIdentifier' ''))
                $identifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'Identifier' ''))
                $tailnet = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'QuickTailnetEmail' ''))
                $user = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $account 'User' ''))
                if ([string]::IsNullOrWhiteSpace($switchIdentifier)) { $switchIdentifier = $identifier }
                if ([string]::IsNullOrWhiteSpace($switchIdentifier)) { $switchIdentifier = $tailnet }
                if ($a -eq ($i - 1)) { $defaultIdentifier = $switchIdentifier }
                $display = Get-QuickAccountSwitchDisplay -Account $account -Index ($a + 1)
                [void]$combo.Items.Add([pscustomobject]@{ Display = $display; SwitchIdentifier = $switchIdentifier; Identifier = $identifier; Tailnet = $tailnet; User = $user; Account = $account })
            }
            $combo.Enabled = ($quickSwitchAvailable -and $combo.Items.Count -gt 0)
            $target = if (-not [string]::IsNullOrWhiteSpace($selectedIdentifier)) { $selectedIdentifier } else { $defaultIdentifier }
            $combo.SelectedIndex = -1
            if (-not [string]::IsNullOrWhiteSpace($selectedDisplay)) {
                for ($x = 0; $x -lt $combo.Items.Count; $x++) {
                    $item = $combo.Items[$x]
                    if ([string]::Equals([string]$item.Display, $selectedDisplay, [System.StringComparison]::OrdinalIgnoreCase)) { $combo.SelectedIndex = $x; break }
                }
            }
            if ($combo.SelectedIndex -lt 0 -and -not [string]::IsNullOrWhiteSpace($target)) {
                for ($x = 0; $x -lt $combo.Items.Count; $x++) {
                    $item = $combo.Items[$x]
                    if (
                        [string]::Equals([string]$item.SwitchIdentifier, $target, [System.StringComparison]::OrdinalIgnoreCase) -or
                        [string]::Equals([string]$item.Identifier, $target, [System.StringComparison]::OrdinalIgnoreCase) -or
                        [string]::Equals([string]$item.Tailnet, $target, [System.StringComparison]::OrdinalIgnoreCase)
                    ) { $combo.SelectedIndex = $x; break }
                }
            }
            if ($combo.SelectedIndex -lt 0 -and $combo.Items.Count -gt 0 -and ($i - 1) -lt $combo.Items.Count) { $combo.SelectedIndex = ($i - 1) }
            $controls = $script:HotkeyControls[$name]
            if ($null -ne $controls) {
                $entry = $cfg.hotkeys.PSObject.Properties[$name].Value
                if (-not $quickSwitchAvailable) {
                    try { $entry.enabled = $false } catch { }
                    $controls.Enabled.Checked = $false
                }
                else {
                    $controls.Enabled.Checked = [bool](Get-ObjectPropertyOrDefault $entry 'enabled' $false)
                }
                if ($null -ne $controls.Label) { $controls.Label.Enabled = [bool]$quickSwitchAvailable }
                $controls.Enabled.Enabled = [bool]$quickSwitchAvailable
                $controls.Capture.Enabled = [bool]$quickSwitchAvailable
                if ($null -ne $controls.Account) { $controls.Account.Enabled = ($quickSwitchAvailable -and $combo.Items.Count -gt 0) }
                $controls.Modifiers.SelectedItem = [string](Get-ObjectPropertyOrDefault $entry 'modifiers' 'Ctrl+Alt+Shift')
                $keyText = [string](Get-ObjectPropertyOrDefault $entry 'key' '')
                if ($controls.Key.Items.Contains($keyText)) { $controls.Key.SelectedItem = $keyText } else { $controls.Key.Text = $keyText }
                $modsText = [string](Get-ObjectPropertyOrDefault $entry 'modifiers' 'Ctrl+Alt+Shift')
                if ([string]::IsNullOrWhiteSpace($keyText)) { $controls.Capture.Text = '' }
                elseif ([string]::IsNullOrWhiteSpace($modsText) -or $modsText -eq 'None') { $controls.Capture.Text = $keyText }
                else { $controls.Capture.Text = $modsText + '+' + $keyText }
            }
        }
        foreach ($name in @($script:QuickAccountSwitchRows.Keys)) {
            $idx = Get-QuickAccountSwitchIndex -Name $name
            if ($idx -le 0 -or $idx -le $count) { continue }
            $row = $script:QuickAccountSwitchRows[$name]
            foreach ($control in @($row.Label,$row.Account,$row.Enabled,$row.Modifiers,$row.Key,$row.Capture)) { if ($null -ne $control) { $control.Visible = $false } }
        }
        if ($null -ne $script:QuickAccountSwitchSectionPanel) {
            $bottom = 320 + ($count * 40) + 18
            try { $script:QuickAccountSwitchSectionPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0,$bottom) } catch { }
            try { if ($null -ne $script:grpHotkeys) { $script:grpHotkeys.Height = [Math]::Min(660, [Math]::Max(430, ($bottom + 44))) } } catch { }
        }
    }
    finally {
        $script:IsLoadingConfig = $previousLoading
    }
    try { Save-Config -Config $cfg } catch { }
    try { Update-HotkeyToolTips } catch { }
    try { [void](Register-Hotkeys -Config $cfg) } catch { Write-LogException -Context 'Register quick account switch hotkeys' -ErrorRecord $_ }
}

function Invoke-QuickAccountSwitchHotkey {
    param([string]$Name)
    $idx = Get-QuickAccountSwitchIndex -Name $Name
    if ($idx -le 0) { return }
    if (-not [bool]$script:QuickAccountSwitchAvailable) { return }

    $cfg = Get-Config
    Ensure-ConfigHotkeyEntries -Config $cfg -Count ([Math]::Max($idx, [int]$script:QuickAccountSwitchNames.Count))
    $entry = $cfg.hotkeys.PSObject.Properties[$Name].Value

    $targetSwitch = [string](Get-ObjectPropertyOrDefault $entry 'account_identifier' '')
    $targetDisplay = [string](Get-ObjectPropertyOrDefault $entry 'account_display' '')
    $selectedUiAccount = $null
    $selectedUiItem = $null

    try {
        if ($script:HotkeyControls.ContainsKey($Name)) {
            $controls = $script:HotkeyControls[$Name]
            if ($null -ne $controls -and $null -ne $controls.Account -and $null -ne $controls.Account.SelectedItem) {
                $selectedUiItem = $controls.Account.SelectedItem
                $uiSwitch = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $selectedUiItem 'SwitchIdentifier' ''))
                $uiIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $selectedUiItem 'Identifier' ''))
                $uiUser = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $selectedUiItem 'User' ''))
                $uiDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $selectedUiItem 'Display' ''))
                $selectedUiAccount = Get-ObjectPropertyOrDefault $selectedUiItem 'Account' $null
                if (-not [string]::IsNullOrWhiteSpace($uiSwitch)) { $targetSwitch = $uiSwitch }
                elseif (-not [string]::IsNullOrWhiteSpace($uiIdentifier)) { $targetSwitch = $uiIdentifier }
                elseif (-not [string]::IsNullOrWhiteSpace($uiUser)) { $targetSwitch = $uiUser }
                if (-not [string]::IsNullOrWhiteSpace($uiDisplay)) { $targetDisplay = $uiDisplay }
            }
        }
    } catch { }

    $matchesTarget = {
        param($Candidate,[string]$Switch,[string]$Display,$UiItem)
        if ($null -eq $Candidate) { return $false }
        $candidateSwitch = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Candidate 'SwitchIdentifier' ''))
        $candidateIdentifier = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Candidate 'Identifier' ''))
        $candidateTailnet = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Candidate 'QuickTailnetEmail' ''))
        $candidateDisplay = ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $Candidate 'QuickDisplay' ''))
        if ([string]::IsNullOrWhiteSpace($candidateDisplay)) { $candidateDisplay = Get-QuickAccountSwitchDisplay -Account $Candidate -Index 0 }
        $uiIdentifier = if ($null -ne $UiItem) { ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $UiItem 'Identifier' '')) } else { '' }
        $uiTailnet = if ($null -ne $UiItem) { ConvertTo-DiagnosticText -Text ([string](Get-ObjectPropertyOrDefault $UiItem 'Tailnet' '')) } else { '' }
        $preciseValues = @($Display,$Switch,$uiIdentifier,$uiTailnet) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
        foreach ($value in $preciseValues) {
            $v = ConvertTo-DiagnosticText -Text ([string]$value)
            if ([string]::IsNullOrWhiteSpace($v)) { continue }
            if (-not [string]::IsNullOrWhiteSpace($candidateDisplay) -and [string]::Equals($candidateDisplay, $v, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            if (-not [string]::IsNullOrWhiteSpace($candidateSwitch) -and [string]::Equals($candidateSwitch, $v, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            if (-not [string]::IsNullOrWhiteSpace($candidateIdentifier) -and [string]::Equals($candidateIdentifier, $v, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            if (-not [string]::IsNullOrWhiteSpace($candidateTailnet) -and [string]::Equals($candidateTailnet, $v, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
        return $false
    }

    $accounts = @(Get-QuickAccountSwitchLoggedAccounts)
    if ($accounts.Count -le 0) { $accounts = @(Get-QuickAccountSwitchLoggedAccounts -Force) }
    if ($accounts.Count -le 0 -and $null -eq $selectedUiAccount) {
        Show-Overlay -Title 'No account available' -Message 'No logged account was returned by tailscale switch --list.' -ErrorStyle
        return
    }

    $account = $selectedUiAccount
    if ($null -eq $account) {
        foreach ($candidate in $accounts) {
            if (& $matchesTarget $candidate $targetSwitch $targetDisplay $selectedUiItem) { $account = $candidate; break }
        }
    }
    if ($null -eq $account -and [string]::IsNullOrWhiteSpace($targetSwitch) -and $idx -le $accounts.Count) { $account = $accounts[$idx - 1] }
    if ($null -eq $account) {
        $freshAccounts = @(Get-QuickAccountSwitchLoggedAccounts -Force)
        foreach ($candidate in $freshAccounts) {
            if (& $matchesTarget $candidate $targetSwitch $targetDisplay $selectedUiItem) { $account = $candidate; break }
        }
        if ($null -eq $account -and [string]::IsNullOrWhiteSpace($targetSwitch) -and $idx -le $freshAccounts.Count) { $account = $freshAccounts[$idx - 1] }
    }

    if ($null -eq $account) {
        Show-Overlay -Title 'Account slot empty' -Message ('Account ' + [string]$idx + ' has no account selected.') -ErrorStyle
        return
    }

    if (-not [bool](Get-ObjectPropertyOrDefault $account 'Active' $false)) {
        $label = Get-TailscaleAccountOverlayIdentifier -Account $account
        if ([string]::IsNullOrWhiteSpace([string]$label)) { $label = if (-not [string]::IsNullOrWhiteSpace($targetDisplay)) { $targetDisplay } else { [string](Get-ObjectPropertyOrDefault $account 'Identifier' 'selected account') } }
        Show-ToggleOverlay -Title 'Switching account' -Message ('Switching to "' + [string]$label + '"...') -Indicator 'Info'
    }
    Invoke-TailscaleAccountSwitch -Account $account
}

function Set-AppTheme {
    $script:Palette = Get-ThemePalette
    $p = $script:Palette
    $script:MainForm.BackColor = $p.FormBack
    $script:MainForm.ForeColor = $p.Text
    $script:headerPanel.BackColor = $p.HeaderBack
    $script:lblTitle.ForeColor = $p.HeaderText
    $script:lblIntro.ForeColor = [System.Drawing.Color]::FromArgb(234,239,248)
    foreach ($g in @($script:grpSummary,$script:grpActions,$script:grpMachines,$script:grpHotkeys,$script:grpPreferences,$script:grpActivity,$script:grpMaintenance)) {
        $g.BackColor = $p.PanelBack
        $g.ForeColor = $p.Text
    }
    foreach ($c in @($script:txtMachineFilter,$script:txtLog,$script:txtPingDetails,$script:txtDnsResolveOutput,$script:txtDnsResolveDomain,$script:txtDnsResolveOtherServer,$script:txtPublicIpOutput,$script:txtMachineDetails,$script:txtMetricsSummary,$script:cmbExitNode,$script:cmbDnsResolveResolver,$script:trkOverlay,$script:trkOverlayOpacity,$script:trkRefresh,$script:trkToggleSoundVolume,$script:numCheckUpdateHours,$script:lblMaintLocalVersion,$script:lblMaintAutoUpdate,$script:lblMaintUpdateStatus,$script:lblMaintMtuStatus,$script:lblMaintMtuVersion,$script:lblMaintMtuService,$script:lblMaintMtuDesiredIPv4,$script:lblMaintMtuDesiredIPv6,$script:lblMaintMtuCheckInterval,$script:lblMaintMtuLastResult,$script:lblMaintMtuLastError,$script:lblMaintMtuRepo,$script:lblMaintMtuAuthor)) {
        if ($null -ne $c) { $c.BackColor = $p.InputBack; $c.ForeColor = $p.Text }
    }
    foreach ($c in @($script:chkStartup,$script:chkStartMinimized,$script:chkCloseToBackground,$script:chkShowTrayIcon,$script:chkAllowLan,$script:chkTogglePopups,$script:chkToggleSounds,$script:chkCheckUpdateEvery,$script:chkControlCheckUpdateEvery,$script:chkExportRedactSensitive,$script:radPublicIpFast,$script:radPublicIpDetailed)) {
        if ($null -eq $c) { continue }
        $c.BackColor = $p.PanelBack; $c.ForeColor = $p.Text
    }
    foreach ($hc in @($script:HotkeyControls.Values)) {
        try {
            $captureControl = Get-ObjectPropertyOrDefault $hc 'Capture' $null
            $accountControl = Get-ObjectPropertyOrDefault $hc 'Account' $null
            $enabledControl = Get-ObjectPropertyOrDefault $hc 'Enabled' $null
            $labelControl = Get-ObjectPropertyOrDefault $hc 'Label' $null
            if ($null -ne $captureControl) { $captureControl.BackColor = $p.InputBack; $captureControl.ForeColor = $p.Text }
            if ($null -ne $accountControl) { $accountControl.BackColor = $p.InputBack; $accountControl.ForeColor = $p.Text }
            if ($null -ne $enabledControl) { $enabledControl.BackColor = $p.PanelBack; $enabledControl.ForeColor = $p.Text }
            if ($null -ne $labelControl) { $labelControl.ForeColor = $p.Text }
        } catch { Write-LogException -Context 'Apply hotkey theme' -ErrorRecord $_ }
    }
    foreach ($label in @($script:QuickAccountSwitchHeader,$script:QuickAccountSwitchSeparator)) {
        try { if ($null -ne $label) { $label.ForeColor = $p.Text } } catch { }
    }
    foreach ($b in @($script:btnRefresh,$script:btnToggleConnect,$script:btnToggleExit,$script:btnToggleDns,$script:btnToggleSubnets,$script:btnToggleIncoming,$script:btnUpdate,$script:btnUninstall,$script:btnCheckControlUpdate,$script:btnCheckClientUpdate,$script:btnRunClientUpdate,$script:btnAdminPanel,$script:btnInstallMtu,$script:btnOpenMtu,$script:btnCheckMtuRepo,$script:btnDetailMetrics,$script:btnDetailClearMetrics,$script:btnDiagStatus,$script:btnDiagNetcheck,$script:btnDiagDns,$script:btnDiagIPs,$script:btnDiagMetrics,$script:btnDiagClear,$script:btnCmdPingAll,$script:btnCmdPingDns,$script:btnCmdPingIPv4,$script:btnCmdPingIPv6,$script:btnCmdWhois,$script:btnDnsResolveRun,$script:btnPublicIpRun,$script:radDnsResolveNoCache,$script:radDnsResolveUseCache,$script:btnCheckControlRepo,$script:btnOpenControlPath,$script:btnExportDiagnostics,$script:btnCheckControlUpdate)) {
        if ($null -eq $b) { continue }
        $b.FlatStyle = 'Standard'
        $b.UseVisualStyleBackColor = $true
    }
    foreach ($label in @($script:lblBackend,$script:lblVersion,$script:lblUser,$script:lblTailnet,$script:lblDevice,$script:lblDnsName,$script:lblIPv4,$script:lblIPv6,$script:lblMtuIPv4,$script:lblMtuIPv6,$script:lblConnSummary,$script:lblDnsInUse,$script:lblDnsState,$script:lblRoutesState,$script:lblIncomingState,$script:lblExitState,$script:lblFooter,$script:lblMachineHelp,$script:lblPingHelp,$script:lblPingNote,$script:lblControlVersion,$script:lblControlLatestVersion,$script:lblControlLastCheck,$script:lblControlAutoUpdate,$script:lblControlPath,$script:lblControlRepo,$script:lblControlAuthor,$script:lblMaintMtuVersion,$script:lblMaintMtuRepo,$script:lblMaintMtuAuthor,$script:lblDiagSelection,$script:lblDnsResolveHelp,$script:lblDnsResolveOther,$script:lblDnsResolveServerPreview,$script:lblPublicIpHelp,$script:lblMetricsInfo,$script:lblCheckUpdateHours,$script:lblControlCheckUpdateHours,$script:lblMaintLatestVersion,$script:lblMaintLastCheck)) {
        if ($null -ne $label) { $label.ForeColor = if ($label -eq $script:lblFooter -or $label -eq $script:lblMachineHelp) { $p.MutedText } else { $p.Text } }
    }
    $script:statusStrip.BackColor = $p.PanelBack
    $script:statusStrip.ForeColor = $p.MutedText
    $script:gridMachines.BackgroundColor = $p.InputBack
    $script:gridMachines.GridColor = $p.Border
    $script:gridMachines.DefaultCellStyle.BackColor = $p.InputBack
    $script:gridMachines.DefaultCellStyle.ForeColor = $p.Text
    $script:gridMachines.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(232,241,255)
    $script:gridMachines.DefaultCellStyle.SelectionForeColor = $p.Text
    $script:gridMachines.ColumnHeadersDefaultCellStyle.BackColor = $p.PanelBack
    $script:gridMachines.ColumnHeadersDefaultCellStyle.ForeColor = $p.MutedText
    $script:gridMachines.EnableHeadersVisualStyles = $false
    $script:gridMachines.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,250,252)
    if ($null -ne $script:gridAccounts) {
        $script:gridAccounts.BackgroundColor = $p.InputBack
        $script:gridAccounts.GridColor = $p.Border
        $script:gridAccounts.DefaultCellStyle.BackColor = $p.InputBack
        $script:gridAccounts.DefaultCellStyle.ForeColor = $p.Text
        $script:gridAccounts.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(232,241,255)
        $script:gridAccounts.DefaultCellStyle.SelectionForeColor = $p.Text
        $script:gridAccounts.ColumnHeadersDefaultCellStyle.BackColor = $p.PanelBack
        $script:gridAccounts.ColumnHeadersDefaultCellStyle.ForeColor = $p.MutedText
        $script:gridAccounts.EnableHeadersVisualStyles = $false
        $script:gridAccounts.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,250,252)
    }
$script:gridPing.BackgroundColor = $p.InputBack
$script:gridPing.GridColor = $p.Border
$script:gridPing.DefaultCellStyle.BackColor = $p.InputBack
$script:gridPing.DefaultCellStyle.ForeColor = $p.Text
$script:gridPing.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(232,241,255)
$script:gridPing.DefaultCellStyle.SelectionForeColor = $p.Text
$script:gridPing.ColumnHeadersDefaultCellStyle.BackColor = $p.PanelBack
$script:gridPing.ColumnHeadersDefaultCellStyle.ForeColor = $p.MutedText
$script:gridPing.EnableHeadersVisualStyles = $false
$script:gridPing.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,250,252)
    foreach ($group in @($grpDiagLocal,$grpDiagDevice)) { if ($null -ne $group) { $group.BackColor = $p.PanelBack; $group.ForeColor = $p.Text } }
    foreach ($tab in @($leftTabs,$machinesTabs,$maintTabs)) { if ($null -ne $tab) { $tab.BackColor = $p.PanelBack; $tab.ForeColor = $p.Text } }
    foreach ($tp in @($tabPrefs,$tabHotkeys,$tabActivity,$tabMaint,$tabMachinesList,$tabMachinesDetails,$tabMachinesPing,$tabMachinesDnsResolve,$tabMachinesPublicIp,$tabUpdateInner,$tabMtuInner,$tabControlInner)) { if ($null -ne $tp) { $tp.BackColor = $p.PanelBack; $tp.ForeColor = $p.Text } }
    if ($null -ne $script:Snapshot) { Update-UiFromSnapshot -Snapshot $script:Snapshot }
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($controlSender,$e)
    try { Write-Log ('UI thread exception: ' + $e.Exception.Message) } catch {}
    try { [System.Windows.Forms.MessageBox]::Show($e.Exception.Message, $script:AppName, 'OK', 'Error') | Out-Null } catch {}
})
Initialize-AppRoot
Remove-StaleTailscaleLoginScripts -OlderThanMinutes 15
try { [void][TailscaleControlTaskbar]::SetCurrentProcessExplicitAppUserModelID($script:AppUserModelId) } catch { }

$sourcePath = [System.IO.Path]::GetFullPath($script:ScriptPath)
$installedPathResolved = [System.IO.Path]::GetFullPath($script:InstalledScriptPath)
if ($sourcePath -ne $installedPathResolved) {
    $null = Install-TailscaleControl
    Write-Log 'Relaunching installed copy without console.'
    Start-Process -FilePath $script:WScriptExe -ArgumentList @($script:LauncherVbsPath)
    return
}

Initialize-StartMenuShortcut
Write-Log ('Launching ' + $script:AppName + ' v' + $script:AppVersion + $(if ($Background) { ' in background mode.' } else { ' in UI mode.' }))

$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, 'Luiz.TailscaleControl.Singleton', [ref]$createdNew)
if (-not $createdNew) {
    Send-ExistingInstanceSignal
    return
}
$activationEventCreated = $false
$script:InstanceActivateEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Luiz.TailscaleControl.Activate', [ref]$activationEventCreated)
$config = Get-Config
$script:Palette = Get-ThemePalette

Write-Log 'UI initialization started.'
try {
$script:MainForm = New-Object System.Windows.Forms.Form
$script:MainForm.Text = $script:AppName
$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$script:MainForm.StartPosition = $(if ($Background) { 'Manual' } else { 'CenterScreen' })
try { $script:MainForm.Opacity = 0.0 } catch { }
if ($Background) {
    $script:MainForm.Location = New-Object System.Drawing.Point(-32000,-32000)
}
$script:IsCompactLayout = ($workingArea.Width -lt 1320 -or $workingArea.Height -lt 780)
$initialWidth = [Math]::Min([int]$workingArea.Width, [Math]::Max(1120, [Math]::Min([int]([math]::Floor($workingArea.Width * 0.96)), 1480)))
$initialHeight = [Math]::Min([int]$workingArea.Height, [Math]::Max(660, [Math]::Min([int]([math]::Floor($workingArea.Height * 0.96)), 940)))
$minWidth = [Math]::Min([int]$workingArea.Width, $(if ($script:IsCompactLayout) { 1040 } else { 1180 }))
$minHeight = [Math]::Min([int]$workingArea.Height, $(if ($script:IsCompactLayout) { 660 } else { 760 }))
$rootPadding = $(if ($script:IsCompactLayout) { New-Object System.Windows.Forms.Padding(8,8,8,6) } else { New-Object System.Windows.Forms.Padding(12,12,12,8) })
$headerRowHeight = $(if ($script:IsCompactLayout) { 56 } else { 64 })
$bannerRowHeight = $(if ($script:IsCompactLayout) { 24 } else { 28 })
$statusRowHeight = $(if ($script:IsCompactLayout) { 22 } else { 24 })
$leftSummaryHeight = $(if ($script:IsCompactLayout) { 228 } else { 254 })
$actionStateRowHeight = $(if ($script:IsCompactLayout) { 36 } else { 40 })
$script:MainForm.Size = New-Object System.Drawing.Size($initialWidth,$initialHeight)
$script:MainForm.MinimumSize = New-Object System.Drawing.Size($minWidth,$minHeight)
$script:MainForm.MaximumSize = New-Object System.Drawing.Size($workingArea.Width,$workingArea.Height)
if (-not $Background) {
    $initialX = [Math]::Max($workingArea.Left, [int]([math]::Floor($workingArea.Left + (($workingArea.Width - $initialWidth) / 2))))
    $initialY = [Math]::Max($workingArea.Top, [int]([math]::Floor($workingArea.Top + (($workingArea.Height - $initialHeight) / 2))))
    $script:MainForm.Location = New-Object System.Drawing.Point($initialX,$initialY)
}
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
Set-MainFormAppIcon
$script:MainForm.Add_HandleCreated({ try { Set-MainFormAppIcon } catch { } })
$script:StartupHidePending = [bool]$Background
if ($Background) {
    try { $script:MainForm.ShowInTaskbar = $false } catch { }
    try { $script:MainForm.Opacity = 0.0 } catch { }
    try { $script:MainForm.WindowState = 'Minimized' } catch { }
}

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 4
$rootLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$rootLayout.Padding = $rootPadding
$null = $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $headerRowHeight)))
$null = $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $bannerRowHeight)))
$null = $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $statusRowHeight)))
$script:MainForm.Controls.Add($rootLayout)

$script:headerPanel = New-Object System.Windows.Forms.Panel
$script:headerPanel.Dock = 'Fill'
$script:headerPanel.Margin = New-Object System.Windows.Forms.Padding(0,0,0,5)
$rootLayout.Controls.Add($script:headerPanel,0,0)

$script:lblTitle = New-Object System.Windows.Forms.Label
$script:lblTitle.Location = New-Object System.Drawing.Point(18,$(if ($script:IsCompactLayout) { 6 } else { 8 }))
$script:lblTitle.Size = New-Object System.Drawing.Size(420,$(if ($script:IsCompactLayout) { 24 } else { 26 }))
$script:lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$script:lblTitle.Text = 'Tailscale Control'
$script:headerPanel.Controls.Add($script:lblTitle) | Out-Null

$script:lblIntro = New-Object System.Windows.Forms.Label
$script:lblIntro.Location = New-Object System.Drawing.Point(20,$(if ($script:IsCompactLayout) { 29 } else { 34 }))
$script:lblIntro.Size = New-Object System.Drawing.Size(1600,$(if ($script:IsCompactLayout) { 16 } else { 18 }))
$script:lblIntro.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$script:lblIntro.Text = 'Fast control for connect, exit node, DNS, subnets, and peer visibility.'
$script:headerPanel.Controls.Add($script:lblIntro) | Out-Null

$script:lblBanner = New-Object System.Windows.Forms.Label
$script:lblBanner.Dock = 'Fill'
$script:lblBanner.Margin = New-Object System.Windows.Forms.Padding(0,0,0,5)
$script:lblBanner.Padding = New-Object System.Windows.Forms.Padding(8,2,8,2)
$script:lblBanner.TextAlign = 'MiddleLeft'
$rootLayout.Controls.Add($script:lblBanner,0,1)

$bodySplit = New-Object System.Windows.Forms.SplitContainer
$script:BodySplit = $bodySplit
$bodySplit.Dock = 'Fill'
$bodySplit.IsSplitterFixed = $false
$bodySplit.FixedPanel = 'Panel1'
$bodySplit.SplitterWidth = 6
$bodySplit.SplitterDistance = [Math]::Max($(if ($script:IsCompactLayout) { 500 } else { 548 }), [int]([math]::Floor($initialWidth * 0.468)))
$bodySplit.Panel1MinSize = $(if ($script:IsCompactLayout) { 360 } else { 390 })
$bodySplit.Margin = New-Object System.Windows.Forms.Padding(0)
$rootLayout.Controls.Add($bodySplit,0,2)

$leftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$leftLayout.Dock = 'Fill'
$leftLayout.ColumnCount = 1
$leftLayout.RowCount = 3
$leftLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $leftSummaryHeight)))
$null = $leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))
$bodySplit.Panel1.Controls.Add($leftLayout)

$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = 'Fill'
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 2
$rightLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $(if ($script:IsCompactLayout) { 224 } else { 236 }))))
$bodySplit.Panel2.Controls.Add($rightLayout)

$script:RightLayoutPreferredActionsRatio = $(if ($script:IsCompactLayout) { 0.31 } else { 0.29 })
$script:RightLayoutPreferredActionsMin = $(if ($script:IsCompactLayout) { 208 } else { 220 })
$script:RightLayoutPreferredActionsMax = $(if ($script:IsCompactLayout) { 272 } else { 300 })
$script:RightLayoutAbsoluteActionsFloor = $(if ($script:IsCompactLayout) { 196 } else { 208 })
$script:RightLayoutMachineFloor = $(if ($script:IsCompactLayout) { 150 } else { 170 })

function Update-RightLayoutSizing {
    try {
        if ($null -eq $rightLayout -or $rightLayout.IsDisposed) { return }
        if ($rightLayout.RowStyles.Count -lt 2) { return }

        $availableHeight = [int]$rightLayout.ClientSize.Height
        if ($availableHeight -le 0) { return }

        $preferredActionsHeight = [int]([math]::Floor($availableHeight * [double]$script:RightLayoutPreferredActionsRatio))
        $minActionsHeight = [int]$script:RightLayoutPreferredActionsMin
        $maxActionsHeight = [int]$script:RightLayoutPreferredActionsMax
        $absoluteActionsFloor = [int]$script:RightLayoutAbsoluteActionsFloor
        $machineFloor = [int]$script:RightLayoutMachineFloor

        if ($preferredActionsHeight -lt $minActionsHeight) { $preferredActionsHeight = $minActionsHeight }
        if ($preferredActionsHeight -gt $maxActionsHeight) { $preferredActionsHeight = $maxActionsHeight }

        $maxActionsByMachineFloor = [int]($availableHeight - $machineFloor)
        if ($maxActionsByMachineFloor -lt $absoluteActionsFloor) {
            $maxActionsByMachineFloor = $absoluteActionsFloor
        }
        if ($preferredActionsHeight -gt $maxActionsByMachineFloor) { $preferredActionsHeight = $maxActionsByMachineFloor }
        if ($preferredActionsHeight -lt $absoluteActionsFloor) { $preferredActionsHeight = $absoluteActionsFloor }

        if ($preferredActionsHeight -gt $availableHeight) { $preferredActionsHeight = $availableHeight }
        if ($preferredActionsHeight -lt 0) { $preferredActionsHeight = 0 }

        $topHeight = [int]($availableHeight - $preferredActionsHeight)
        if ($topHeight -lt $machineFloor) {
            $topHeight = $machineFloor
            $preferredActionsHeight = [int]($availableHeight - $topHeight)
            if ($preferredActionsHeight -lt $absoluteActionsFloor) {
                $preferredActionsHeight = $absoluteActionsFloor
                $topHeight = [int]($availableHeight - $preferredActionsHeight)
            }
        }
        if ($topHeight -lt 0) { $topHeight = 0 }

        if ($rightLayout.RowStyles[0].SizeType -ne [System.Windows.Forms.SizeType]::Absolute) {
            $rightLayout.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
        }
        if ($rightLayout.RowStyles[1].SizeType -ne [System.Windows.Forms.SizeType]::Absolute) {
            $rightLayout.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Absolute
        }
        if ([int]$rightLayout.RowStyles[0].Height -ne $topHeight) {
            $rightLayout.RowStyles[0].Height = $topHeight
        }
        if ([int]$rightLayout.RowStyles[1].Height -ne $preferredActionsHeight) {
            $rightLayout.RowStyles[1].Height = $preferredActionsHeight
        }
    }
    catch {
        Write-Log ('Right layout sizing failed: ' + $_.Exception.Message)
    }
}

$bodySplit.add_SizeChanged({
    try { Update-RightLayoutSizing } catch { Write-Log ('Split layout size change failed: ' + $_.Exception.Message) }
})
$rightLayout.add_SizeChanged({
    try { Update-RightLayoutSizing } catch { Write-Log ('Right layout size change failed: ' + $_.Exception.Message) }
})

$script:grpSummary = New-Object System.Windows.Forms.GroupBox
$script:grpSummary.Text = 'Current Status'
$script:grpSummary.Dock = 'Fill'
$script:grpSummary.Margin = New-Object System.Windows.Forms.Padding(0,0,0,7)
$leftLayout.Controls.Add($script:grpSummary,0,0)

$script:SummaryLabelFont = New-Object System.Drawing.Font('Segoe UI', 9)
$script:SummaryValueFont = New-Object System.Drawing.Font('Segoe UI', 9)

$summaryTable = New-Object System.Windows.Forms.TableLayoutPanel
$summaryTable.Dock = 'Fill'
$summaryTable.ColumnCount = 4
$summaryTable.RowCount = 7
$summaryTable.Padding = New-Object System.Windows.Forms.Padding(12,14,12,10)
$summaryTable.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
$null = $summaryTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138)))
$null = $summaryTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$null = $summaryTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 124)))
$null = $summaryTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($r = 0; $r -lt 7; $r++) { $null = $summaryTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 14.2857))) }
$script:grpSummary.Controls.Add($summaryTable)

function Set-SummarySingleLineLabel {
    param(
        [System.Windows.Forms.Label]$Label,
        [bool]$Ellipsis = $false
    )
    $Label.AutoSize = $false
    $Label.Dock = 'Fill'
    $Label.Anchor = ([System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
    $Label.Height = 24
    $Label.MinimumSize = New-Object System.Drawing.Size(0,24)
    $Label.MaximumSize = New-Object System.Drawing.Size(0,24)
    $Label.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
    $Label.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
    $Label.TextAlign = 'MiddleLeft'
    $Label.UseCompatibleTextRendering = $false
    $Label.UseMnemonic = $false
    $Label.AutoEllipsis = $false
    try { if ($Label -is [TailscaleControlSummaryLabel]) { $Label.UseEndEllipsis = $Ellipsis } } catch { }
}

function Add-SummaryPair {
    param([int]$Row,[int]$LabelCol,[string]$LabelText,[string]$ValueName)
    $lbl = New-Object TailscaleControlSummaryLabel
    $lbl.Text = $LabelText
    $lbl.Font = $script:SummaryLabelFont
    Set-SummarySingleLineLabel -Label $lbl -Ellipsis $false
    $summaryTable.Controls.Add($lbl,$LabelCol,$Row) | Out-Null

    $val = New-Object TailscaleControlSummaryLabel
    $val.Text = '-'
    $val.Font = $script:SummaryValueFont
    Set-SummarySingleLineLabel -Label $val -Ellipsis $true
    $summaryTable.Controls.Add($val,($LabelCol + 1),$Row) | Out-Null
    Set-Variable -Scope Script -Name $ValueName -Value $val
}
Add-SummaryPair -Row 0 -LabelCol 0 -LabelText 'Backend Status:' -ValueName 'lblBackend'
Add-SummaryPair -Row 1 -LabelCol 0 -LabelText 'Current User:' -ValueName 'lblUser'
Add-SummaryPair -Row 2 -LabelCol 0 -LabelText 'Account Email:' -ValueName 'lblUserEmail'
Add-SummaryPair -Row 3 -LabelCol 0 -LabelText 'Device Name:' -ValueName 'lblDevice'
Add-SummaryPair -Row 4 -LabelCol 0 -LabelText 'MagicDNS Name:' -ValueName 'lblDnsName'
Add-SummaryPair -Row 5 -LabelCol 0 -LabelText 'DNS Mode:' -ValueName 'lblConnSummary'

$lblDnsCaption = New-Object TailscaleControlSummaryLabel
$lblDnsCaption.Text = 'DNS Resolver:'
$lblDnsCaption.Font = $script:SummaryLabelFont
Set-SummarySingleLineLabel -Label $lblDnsCaption -Ellipsis $false
$summaryTable.Controls.Add($lblDnsCaption,0,6) | Out-Null

$script:lblDnsInUse = New-Object TailscaleControlSummaryLabel
$script:lblDnsInUse.Text = '-'
$script:lblDnsInUse.Font = $script:SummaryValueFont
Set-SummarySingleLineLabel -Label $script:lblDnsInUse -Ellipsis $true
$summaryTable.Controls.Add($script:lblDnsInUse,1,6) | Out-Null
$summaryTable.SetColumnSpan($script:lblDnsInUse,3)

Add-SummaryPair -Row 0 -LabelCol 2 -LabelText 'Version:' -ValueName 'lblVersion'
Add-SummaryPair -Row 1 -LabelCol 2 -LabelText 'Tailnet:' -ValueName 'lblTailnet'
Add-SummaryPair -Row 2 -LabelCol 2 -LabelText 'IPv4 Address:' -ValueName 'lblIPv4'
Add-SummaryPair -Row 3 -LabelCol 2 -LabelText 'IPv6 Address:' -ValueName 'lblIPv6'
$script:lblIPv6.Padding = New-Object System.Windows.Forms.Padding(0)
Add-SummaryPair -Row 4 -LabelCol 2 -LabelText 'IPv4 MTU:' -ValueName 'lblMtuIPv4'
Add-SummaryPair -Row 5 -LabelCol 2 -LabelText 'IPv6 MTU:' -ValueName 'lblMtuIPv6'


$script:btnUpdate = New-ActionButton -Text 'Update' -Left 0 -Top 0 -Width 96
$script:btnUninstall = New-ActionButton -Text 'Uninstall' -Left 0 -Top 0 -Width 96
foreach ($b in @($script:btnUpdate,$script:btnUninstall)) {
    $b.Dock = 'Fill'
    $b.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
}

$script:grpActions = New-Object System.Windows.Forms.GroupBox
$script:grpActions.Text = 'Quick Actions'
$script:grpActions.Dock = 'Fill'
$script:grpActions.Margin = New-Object System.Windows.Forms.Padding(0)
$rightLayout.Controls.Add($script:grpActions,0,1)

$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.AutoSize = $false
$actionsLayout.ColumnCount = 1
$actionsLayout.RowCount = 2
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(12,6,12,10)
$null = $actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $actionStateRowHeight)))
$null = $actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:grpActions.Controls.Add($actionsLayout)

$statesTable = New-Object System.Windows.Forms.TableLayoutPanel
$statesTable.Dock = 'Fill'
$statesTable.ColumnCount = 4
$statesTable.RowCount = 2
$statesTable.Margin = New-Object System.Windows.Forms.Padding(0)
for ($i = 0; $i -lt 4; $i++) { $null = $statesTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) }
$null = $statesTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
$null = $statesTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
$actionsLayout.Controls.Add($statesTable,0,0)

function Add-StatePair {
    param([int]$Col,[string]$LabelText,[string]$ValueName)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Dock = 'Fill'
    $lbl.TextAlign = 'BottomLeft'
    $lbl.Font = $script:SummaryLabelFont
    $statesTable.Controls.Add($lbl,$Col,0) | Out-Null

    $valuePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $valuePanel.Dock = 'Fill'
    $valuePanel.Margin = New-Object System.Windows.Forms.Padding(0)
    $valuePanel.Padding = New-Object System.Windows.Forms.Padding(0)
    $valuePanel.ColumnCount = 2
    $valuePanel.RowCount = 1
    $null = $valuePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
    $null = $valuePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $dot = New-Object System.Windows.Forms.Label
    $dot.Text = ''
    $dot.Dock = 'Fill'
    $dot.TextAlign = 'MiddleCenter'
    $dot.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $dot.Image = New-MainStateDotBitmap -Known:$false -On:$false
    $dot.ForeColor = Get-StateDotColor -Known:$false -On:$false
    $dot.Margin = New-Object System.Windows.Forms.Padding(0)
    $dot.Padding = New-Object System.Windows.Forms.Padding(0)
    $valuePanel.Controls.Add($dot,0,0) | Out-Null

    $val = New-Object System.Windows.Forms.Label
    $val.Text = '-'
    $val.Dock = 'Fill'
    $val.TextAlign = 'MiddleLeft'
    $val.Font = $script:SummaryValueFont
    $val.Margin = New-Object System.Windows.Forms.Padding(0)
    $valuePanel.Controls.Add($val,1,0) | Out-Null

    $statesTable.Controls.Add($valuePanel,$Col,1) | Out-Null
    Set-Variable -Scope Script -Name $ValueName -Value $val
    $indicatorName = $ValueName -replace '^lbl','ind'
    Set-Variable -Scope Script -Name $indicatorName -Value $dot
}
Add-StatePair -Col 0 -LabelText 'Accept DNS' -ValueName 'lblDnsState'
Add-StatePair -Col 1 -LabelText 'Subnets' -ValueName 'lblRoutesState'
Add-StatePair -Col 2 -LabelText 'Incoming' -ValueName 'lblIncomingState'
Add-StatePair -Col 3 -LabelText 'Exit node' -ValueName 'lblExitState'

$buttonsGrid = New-Object System.Windows.Forms.TableLayoutPanel
$buttonsGrid.Dock = 'Fill'
$buttonsGrid.ColumnCount = 2
$buttonsGrid.RowCount = 3
$buttonsGrid.Margin = New-Object System.Windows.Forms.Padding(0)
$buttonsGrid.Padding = New-Object System.Windows.Forms.Padding(0,10,0,0)
$null = $buttonsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$null = $buttonsGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$null = $buttonsGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.3333)))
$null = $buttonsGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.3333)))
$null = $buttonsGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.3334)))
$actionsLayout.Controls.Add($buttonsGrid,0,1)

$script:btnRefresh = New-ActionButton -Text 'Refresh' -Left 0 -Top 0 -Width 100
$script:btnToggleConnect = New-ActionButton -Text 'Toggle Connect' -Left 0 -Top 0 -Width 100
$script:btnToggleExit = New-ActionButton -Text 'Toggle Exit Node' -Left 0 -Top 0 -Width 100
$script:btnToggleDns = New-ActionButton -Text 'Toggle DNS' -Left 0 -Top 0 -Width 100
$script:btnToggleSubnets = New-ActionButton -Text 'Toggle Subnets' -Left 0 -Top 0 -Width 100
$script:btnToggleIncoming = New-ActionButton -Text 'Toggle Incoming' -Left 0 -Top 0 -Width 100

$buttons = @($script:btnRefresh,$script:btnToggleConnect,$script:btnToggleExit,$script:btnToggleDns,$script:btnToggleSubnets,$script:btnToggleIncoming)
for ($i = 0; $i -lt $buttons.Count; $i++) {
    $b = $buttons[$i]
    $b.Dock = 'Fill'
    $row = [int]([math]::Floor($i / 2))
    $col = $i % 2
    $b.Margin = New-Object System.Windows.Forms.Padding(4,0,4,8)
    $b.MinimumSize = New-Object System.Drawing.Size(0,34)
    $buttonsGrid.Controls.Add($b,$col,$row) | Out-Null
}

$leftTabs = New-Object System.Windows.Forms.TabControl
$script:leftTabs = $leftTabs
$leftTabs.Dock = 'Fill'
$leftTabs.Margin = New-Object System.Windows.Forms.Padding(0)
$leftTabs.Padding = New-Object System.Drawing.Point(8,5)
$leftTabs.ItemSize = New-Object System.Drawing.Size(94,28)
$leftTabs.SizeMode = 'Fixed'
$leftTabs.DrawMode = 'Normal'
$leftTabs.Appearance = 'Normal'
$leftTabs.Multiline = $false
$tabMaint = New-Object System.Windows.Forms.TabPage
$tabMaint.Text = 'Maintenance'
$script:tabAccount = New-Object System.Windows.Forms.TabPage
$script:tabAccount.Text = 'Account'
$tabPrefs = New-Object System.Windows.Forms.TabPage
$tabPrefs.Text = 'Preferences'
$tabHotkeys = New-Object System.Windows.Forms.TabPage
$tabHotkeys.Text = 'Hotkeys'
$tabActivity = New-Object System.Windows.Forms.TabPage
$tabActivity.Text = 'Activity'
$null = $leftTabs.TabPages.Add($tabPrefs)
$null = $leftTabs.TabPages.Add($tabHotkeys)
$null = $leftTabs.TabPages.Add($tabActivity)
$null = $leftTabs.TabPages.Add($tabMaint)
$null = $leftTabs.TabPages.Add($script:tabAccount)
$leftLayout.Controls.Add($leftTabs,0,1)

$script:grpPreferences = New-Object System.Windows.Forms.GroupBox
$script:grpPreferences.Text = 'Preferences'
$script:grpPreferences.Dock = 'Top'
$script:grpPreferences.Anchor = 'Top,Left,Right'
$script:grpPreferences.Margin = New-Object System.Windows.Forms.Padding(0)
$script:grpPreferences.Padding = New-Object System.Windows.Forms.Padding(8,8,8,8)
$script:grpPreferences.AutoSize = $false
$script:grpPreferences.MinimumSize = New-Object System.Drawing.Size(0,508)
$script:grpPreferences.Height = 540
$tabPrefs.AutoScroll = $true
$tabPrefs.Padding = New-Object System.Windows.Forms.Padding(8,8,8,8)
$tabPrefs.Controls.Add($script:grpPreferences)

$prefLayout = New-Object System.Windows.Forms.TableLayoutPanel
$prefLayout.Dock = 'Fill'
$prefLayout.ColumnCount = 2
$prefLayout.RowCount = 14
$prefLayout.Padding = New-Object System.Windows.Forms.Padding(12,6,12,8)
$prefLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$prefLayout.AutoSize = $false
$null = $prefLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 236)))
$null = $prefLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
for ($i=0; $i -lt 13; $i++) { $null = $prefLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) }
$null = $prefLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:grpPreferences.Controls.Add($prefLayout)

$script:chkStartup = New-Object System.Windows.Forms.CheckBox
$script:chkStartup.Text = 'Start with Windows'
$script:chkStartup.Dock = 'Fill'
$script:chkStartup.Checked = $true
$script:chkStartup.Enabled = $true
$prefLayout.Controls.Add($script:chkStartup,0,0)

$script:chkStartMinimized = New-Object System.Windows.Forms.CheckBox
$script:chkStartMinimized.Text = 'Start minimized in background'
$script:chkStartMinimized.Dock = 'Fill'
$script:chkStartMinimized.Checked = $true
$script:chkStartMinimized.Enabled = $true
$prefLayout.Controls.Add($script:chkStartMinimized,0,1)

$script:chkCloseToBackground = New-Object System.Windows.Forms.CheckBox
$script:chkCloseToBackground.Text = 'Close hides to background'
$script:chkCloseToBackground.Dock = 'Fill'
$prefLayout.Controls.Add($script:chkCloseToBackground,0,2)

$script:chkShowTrayIcon = New-Object System.Windows.Forms.CheckBox
$script:chkShowTrayIcon.Text = 'Show tray icon'
$script:chkShowTrayIcon.Dock = 'Fill'
$prefLayout.Controls.Add($script:chkShowTrayIcon,0,3)

$script:chkAllowLan = New-Object System.Windows.Forms.CheckBox
$script:chkAllowLan.Text = 'Allow local network access'
$script:chkAllowLan.Dock = 'Fill'
$prefLayout.SetColumnSpan($script:chkAllowLan,2)
$prefLayout.Controls.Add($script:chkAllowLan,0,12)

$script:chkTogglePopups = New-Object System.Windows.Forms.CheckBox
$script:chkTogglePopups.Text = 'Show overlay when using toggle buttons or hotkeys'
$script:chkTogglePopups.AutoSize = $false
$script:chkTogglePopups.Dock = 'Fill'
$script:chkTogglePopups.CheckAlign = 'MiddleLeft'
$script:chkTogglePopups.TextAlign = 'MiddleLeft'
$script:chkTogglePopups.Margin = New-Object System.Windows.Forms.Padding(3,3,3,3)
$script:chkTogglePopups.Padding = New-Object System.Windows.Forms.Padding(0)
$script:chkTogglePopups.AutoEllipsis = $false
$prefLayout.SetColumnSpan($script:chkTogglePopups,2)
$prefLayout.Controls.Add($script:chkTogglePopups,0,4)

$script:chkShowCurrentDeviceInfoInTray = New-Object System.Windows.Forms.CheckBox
$script:chkShowCurrentDeviceInfoInTray.Text = 'Show current device info in tray'
$script:chkShowCurrentDeviceInfoInTray.Dock = 'Fill'
$script:chkShowCurrentDeviceInfoInTray.CheckAlign = 'MiddleLeft'
$script:chkShowCurrentDeviceInfoInTray.TextAlign = 'MiddleLeft'
$script:chkShowCurrentDeviceInfoInTray.Margin = New-Object System.Windows.Forms.Padding(3,3,3,3)
$prefLayout.SetColumnSpan($script:chkShowCurrentDeviceInfoInTray,2)
$prefLayout.Controls.Add($script:chkShowCurrentDeviceInfoInTray,0,5)

$script:chkToggleSounds = New-Object System.Windows.Forms.CheckBox
$script:chkToggleSounds.Text = 'Play Windows sound when a toggle turns on or off'
$script:chkToggleSounds.Dock = 'Fill'
$script:chkToggleSounds.CheckAlign = 'MiddleLeft'
$script:chkToggleSounds.TextAlign = 'MiddleLeft'
$script:chkToggleSounds.Margin = New-Object System.Windows.Forms.Padding(3,3,3,3)
$prefLayout.SetColumnSpan($script:chkToggleSounds,2)
$prefLayout.Controls.Add($script:chkToggleSounds,0,6)

$lblOpacity = New-Object System.Windows.Forms.Label
$lblOpacity.Text = 'Overlay opacity'
$lblOpacity.Dock = 'Fill'
$lblOpacity.TextAlign = 'MiddleLeft'
$lblOpacity.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$prefLayout.Controls.Add($lblOpacity,0,7)

$overlayOpacityPanel = New-Object System.Windows.Forms.TableLayoutPanel
$overlayOpacityPanel.Dock = 'Fill'
$overlayOpacityPanel.ColumnCount = 2
$overlayOpacityPanel.RowCount = 1
$overlayOpacityPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $overlayOpacityPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $overlayOpacityPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$script:trkOverlayOpacity = New-Object TailscaleControlNoWheelTrackBar
$script:trkOverlayOpacity.Minimum = 0
$script:trkOverlayOpacity.Maximum = 100
$script:trkOverlayOpacity.TickFrequency = 10
$script:trkOverlayOpacity.SmallChange = 1
$script:trkOverlayOpacity.LargeChange = 5
$script:trkOverlayOpacity.Dock = 'Fill'
$script:trkOverlayOpacity.AutoSize = $false
$script:trkOverlayOpacity.Height = 30
$script:lblOverlayOpacityValue = New-Object System.Windows.Forms.Label
$script:lblOverlayOpacityValue.Dock = 'Fill'
$script:lblOverlayOpacityValue.TextAlign = 'MiddleRight'
$overlayOpacityPanel.Controls.Add($script:trkOverlayOpacity,0,0)
$overlayOpacityPanel.Controls.Add($script:lblOverlayOpacityValue,1,0)
$prefLayout.Controls.Add($overlayOpacityPanel,1,7)

$lblOverlay = New-Object System.Windows.Forms.Label
$lblOverlay.Text = 'Overlay seconds'
$lblOverlay.Dock = 'Fill'
$lblOverlay.TextAlign = 'MiddleLeft'
$lblOverlay.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$prefLayout.Controls.Add($lblOverlay,0,8)

$overlaySecondsPanel = New-Object System.Windows.Forms.TableLayoutPanel
$overlaySecondsPanel.Dock = 'Fill'
$overlaySecondsPanel.ColumnCount = 2
$overlaySecondsPanel.RowCount = 1
$overlaySecondsPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $overlaySecondsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $overlaySecondsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$script:trkOverlay = New-Object TailscaleControlNoWheelTrackBar
$script:trkOverlay.Minimum = 5
$script:trkOverlay.Maximum = 50
$script:trkOverlay.TickFrequency = 5
$script:trkOverlay.SmallChange = 1
$script:trkOverlay.LargeChange = 5
$script:trkOverlay.Dock = 'Fill'
$script:trkOverlay.AutoSize = $false
$script:trkOverlay.Height = 30
$script:lblOverlayValue = New-Object System.Windows.Forms.Label
$script:lblOverlayValue.Dock = 'Fill'
$script:lblOverlayValue.TextAlign = 'MiddleRight'
$overlaySecondsPanel.Controls.Add($script:trkOverlay,0,0)
$overlaySecondsPanel.Controls.Add($script:lblOverlayValue,1,0)
$prefLayout.Controls.Add($overlaySecondsPanel,1,8)

$lblRefresh = New-Object System.Windows.Forms.Label
$lblRefresh.Text = 'Refresh seconds'
$lblRefresh.Dock = 'Fill'
$lblRefresh.TextAlign = 'MiddleLeft'
$lblRefresh.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$prefLayout.Controls.Add($lblRefresh,0,9)

$refreshPanel = New-Object System.Windows.Forms.TableLayoutPanel
$refreshPanel.Dock = 'Fill'
$refreshPanel.ColumnCount = 2
$refreshPanel.RowCount = 1
$refreshPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $refreshPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $refreshPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$script:trkRefresh = New-Object TailscaleControlNoWheelTrackBar
$script:trkRefresh.Minimum = 3
$script:trkRefresh.Maximum = 60
$script:trkRefresh.TickFrequency = 3
$script:trkRefresh.SmallChange = 1
$script:trkRefresh.LargeChange = 5
$script:trkRefresh.Dock = 'Fill'
$script:trkRefresh.AutoSize = $false
$script:trkRefresh.Height = 30
$script:lblRefreshValue = New-Object System.Windows.Forms.Label
$script:lblRefreshValue.Dock = 'Fill'
$script:lblRefreshValue.TextAlign = 'MiddleRight'
$refreshPanel.Controls.Add($script:trkRefresh,0,0)
$refreshPanel.Controls.Add($script:lblRefreshValue,1,0)
$prefLayout.Controls.Add($refreshPanel,1,9)

$lblToggleSoundVolume = New-Object System.Windows.Forms.Label
$lblToggleSoundVolume.Text = 'Toggle sound volume'
$lblToggleSoundVolume.Dock = 'Fill'
$lblToggleSoundVolume.TextAlign = 'MiddleLeft'
$lblToggleSoundVolume.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$prefLayout.Controls.Add($lblToggleSoundVolume,0,10)

$toggleSoundVolumePanel = New-Object System.Windows.Forms.TableLayoutPanel
$toggleSoundVolumePanel.Dock = 'Fill'
$toggleSoundVolumePanel.ColumnCount = 2
$toggleSoundVolumePanel.RowCount = 1
$toggleSoundVolumePanel.Margin = New-Object System.Windows.Forms.Padding(0)
$null = $toggleSoundVolumePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $toggleSoundVolumePanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$script:trkToggleSoundVolume = New-Object TailscaleControlNoWheelTrackBar
$script:trkToggleSoundVolume.Minimum = 0
$script:trkToggleSoundVolume.Maximum = 100
$script:trkToggleSoundVolume.TickFrequency = 10
$script:trkToggleSoundVolume.SmallChange = 1
$script:trkToggleSoundVolume.LargeChange = 5
$script:trkToggleSoundVolume.Dock = 'Fill'
$script:trkToggleSoundVolume.AutoSize = $false
$script:trkToggleSoundVolume.Height = 30
$script:lblToggleSoundVolumeValue = New-Object System.Windows.Forms.Label
$script:lblToggleSoundVolumeValue.Dock = 'Fill'
$script:lblToggleSoundVolumeValue.TextAlign = 'MiddleRight'
$toggleSoundVolumePanel.Controls.Add($script:trkToggleSoundVolume,0,0)
$toggleSoundVolumePanel.Controls.Add($script:lblToggleSoundVolumeValue,1,0)
$prefLayout.Controls.Add($toggleSoundVolumePanel,1,10)

$lblExitPref = New-Object System.Windows.Forms.Label
$lblExitPref.Text = 'Preferred exit node'
$lblExitPref.Dock = 'Fill'
$lblExitPref.TextAlign = 'MiddleLeft'
$lblExitPref.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$prefLayout.Controls.Add($lblExitPref,0,11)

$script:cmbExitNode = New-Object System.Windows.Forms.ComboBox
$script:cmbExitNode.DropDownStyle = 'DropDownList'
$script:cmbExitNode.Dock = 'Fill'
$prefLayout.Controls.Add($script:cmbExitNode,1,11)

$script:grpHotkeys = New-Object System.Windows.Forms.GroupBox
$script:grpHotkeys.Text = 'Hotkeys'
$script:grpHotkeys.Dock = 'Top'
$script:grpHotkeys.Height = 520
$script:grpHotkeys.Margin = New-Object System.Windows.Forms.Padding(0)
$tabHotkeys.AutoScroll = $true
$tabHotkeys.Controls.Add($script:grpHotkeys)

$hotkeyPanel = New-Object System.Windows.Forms.Panel
$hotkeyPanel.Dock = 'Fill'
$hotkeyPanel.AutoScroll = $true
$hotkeyPanel.Padding = New-Object System.Windows.Forms.Padding(12,12,12,12)
$script:grpHotkeys.Controls.Add($hotkeyPanel)

$script:lblFooter = New-Object System.Windows.Forms.Label
$script:lblFooter.Text = 'Compact global hotkeys. Changes save automatically.'
$script:lblFooter.Location = New-Object System.Drawing.Point(16,16)
$script:lblFooter.Size = New-Object System.Drawing.Size(500,24)
$hotkeyPanel.Controls.Add($script:lblFooter) | Out-Null

New-HotkeyRow -Parent $hotkeyPanel -Name 'ToggleConnect' -Top 52 -LabelText 'Toggle Connect'
New-HotkeyRow -Parent $hotkeyPanel -Name 'ToggleExitNode' -Top 92 -LabelText 'Toggle Exit Node'
New-HotkeyRow -Parent $hotkeyPanel -Name 'ToggleDns' -Top 132 -LabelText 'Toggle DNS'
New-HotkeyRow -Parent $hotkeyPanel -Name 'ToggleSubnets' -Top 172 -LabelText 'Toggle Subnets'
New-HotkeyRow -Parent $hotkeyPanel -Name 'ToggleIncoming' -Top 212 -LabelText 'Toggle Incoming'
New-HotkeyRow -Parent $hotkeyPanel -Name 'ShowSettings' -Top 252 -LabelText 'Toggle Settings'
New-QuickAccountSwitchSection -Parent $hotkeyPanel -Top 300

$script:grpMachines = New-Object System.Windows.Forms.GroupBox
$script:grpMachines.Text = 'Machines'
$script:grpMachines.Dock = 'Fill'
$script:grpMachines.Margin = New-Object System.Windows.Forms.Padding(0,0,0,7)
$rightLayout.Controls.Add($script:grpMachines,0,0)

$machinesRoot = New-Object System.Windows.Forms.TableLayoutPanel
$machinesRoot.Dock = 'Fill'
$machinesRoot.ColumnCount = 1
$machinesRoot.RowCount = 1
$machinesRoot.Padding = New-Object System.Windows.Forms.Padding(8,10,8,8)
$null = $machinesRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:grpMachines.Controls.Add($machinesRoot)

$machinesTabs = New-Object System.Windows.Forms.TabControl
$machinesTabs.Dock = 'Fill'
$machinesTabs.Padding = New-Object System.Drawing.Point(8,6)
$machinesTabs.DrawMode = 'Normal'
$machinesTabs.Appearance = 'Normal'
$machinesTabs.Multiline = $false
$machinesTabs.SizeMode = 'Fixed'
$machinesTabs.ItemSize = New-Object System.Drawing.Size(96,28)
$machinesRoot.Controls.Add($machinesTabs,0,0)

$tabMachinesList = New-Object System.Windows.Forms.TabPage
$tabMachinesList.Text = 'List'
$tabMachinesDetails = New-Object System.Windows.Forms.TabPage
$tabMachinesDetails.Text = 'Info'
$tabMachinesPing = New-Object System.Windows.Forms.TabPage
$tabMachinesPing.Text = 'Tools'
$tabMachinesDnsResolve = New-Object System.Windows.Forms.TabPage
$tabMachinesDnsResolve.Text = 'DNS Resolve'
$tabMachinesPublicIp = New-Object System.Windows.Forms.TabPage
$tabMachinesPublicIp.Text = 'Public IP'
$null = $machinesTabs.TabPages.Add($tabMachinesList)
$null = $machinesTabs.TabPages.Add($tabMachinesDetails)
$null = $machinesTabs.TabPages.Add($tabMachinesPing)
$null = $machinesTabs.TabPages.Add($tabMachinesDnsResolve)
$null = $machinesTabs.TabPages.Add($tabMachinesPublicIp)

$machinesLayout = New-Object System.Windows.Forms.TableLayoutPanel
$machinesLayout.Dock = 'Fill'
$machinesLayout.ColumnCount = 1
$machinesLayout.RowCount = 4
$machinesLayout.Padding = New-Object System.Windows.Forms.Padding(10,10,10,10)
$null = $machinesLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24)))
$null = $machinesLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$null = $machinesLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$null = $machinesLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabMachinesList.Controls.Add($machinesLayout)

$script:lblMachineHelp = New-Object System.Windows.Forms.Label
$script:lblMachineHelp.Text = 'Search by machine, owner, IP, MagicDNS, or last seen.'
$script:lblMachineHelp.Dock = 'Fill'
$machinesLayout.Controls.Add($script:lblMachineHelp,0,0)

$script:txtMachineFilter = New-Object System.Windows.Forms.TextBox
$script:txtMachineFilter.Dock = 'Fill'
$machinesLayout.Controls.Add($script:txtMachineFilter,0,1)

$script:gridMachines = New-Object System.Windows.Forms.DataGridView
$script:gridMachines.Dock = 'Fill'
$script:gridMachines.ReadOnly = $true
$script:gridMachines.RowHeadersVisible = $false
$script:gridMachines.AllowUserToAddRows = $false
$script:gridMachines.AllowUserToDeleteRows = $false
$script:gridMachines.AllowUserToResizeRows = $false
$script:gridMachines.SelectionMode = 'FullRowSelect'
$script:gridMachines.MultiSelect = $false
$script:gridMachines.AutoSizeColumnsMode = 'Fill'
$script:gridMachines.AutoGenerateColumns = $false
$script:gridMachines.BorderStyle = 'FixedSingle'
$script:gridMachines.ColumnHeadersHeight = 34
$script:gridMachines.RowTemplate.Height = 28
$script:gridMachines.ScrollBars = 'Vertical'
$script:gridMachines.ClipboardCopyMode = 'EnableWithoutHeaderText'
$script:gridMachines.Columns.Clear()

$machineColumns = @(
    @{ Name='Machine'; Header='Machine'; Width=72; MinWidth=72 },
    @{ Name='Owner'; Header='Owner'; Width=82; MinWidth=72 },
    @{ Name='IPv4'; Header='IPv4'; Width=88; MinWidth=72 },
    @{ Name='IPv6'; Header='IPv6'; Width=150; MinWidth=116 },
    @{ Name='OS'; Header='OS'; Width=56; MinWidth=56 },
    @{ Name='Connection'; Header='Conn'; Width=106; MinWidth=96 },
    @{ Name='DNSName'; Header='MagicDNS'; Width=162; MinWidth=108 },
    @{ Name='LastSeen'; Header='Last Seen'; Width=100; MinWidth=84 }
)
foreach ($col in $machineColumns) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = [string]$col.Name
    $c.HeaderText = [string]$col.Header
    $c.Width = [int]$col.Width
    $c.MinimumWidth = $(if ($col.ContainsKey('MinWidth')) { [int]$col.MinWidth } else { [Math]::Max(72, [int]([math]::Floor($c.Width * 0.8))) })
    $c.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $c.Resizable = [System.Windows.Forms.DataGridViewTriState]::True
    $c.SortMode = 'NotSortable'
    $null = $script:gridMachines.Columns.Add($c)
}
Set-MachineColumnLayout -UseConfig
$script:MachineColumnsInitialized = $true
$script:gridMachines.add_SizeChanged({
    if (-not $script:IsApplyingMachineColumnLayout) { Set-MachineColumnLayout -PreserveCurrent }
})
$script:gridMachines.add_ColumnWidthChanged({
    param($controlSender,$e)
    if ($script:IsApplyingMachineColumnLayout -or $null -eq $e -or $null -eq $e.Column) { return }
    Save-MachineColumnLayoutFromGrid
})

$script:lblMachineHint = New-Object System.Windows.Forms.Label
$script:lblMachineHint.Text = 'Select a device, then use Info for the summary and Commands for network or device actions.'
$script:lblMachineHint.Dock = 'Fill'
$gridContext = New-Object System.Windows.Forms.ContextMenuStrip
$gridCopyCell = $gridContext.Items.Add('Copy Cell')
$gridCopyRow = $gridContext.Items.Add('Copy Row')
$script:gridMachines.ContextMenuStrip = $gridContext

function Get-SelectedMachineRowText {
    if ($null -eq $script:gridMachines.CurrentRow) { return '' }
    $values = New-Object System.Collections.Generic.List[string]
    foreach ($cell in $script:gridMachines.CurrentRow.Cells) {
        $values.Add([string]$cell.Value)
    }
    return (($values | Where-Object { $_ -ne $null }) -join "`t")
}

$gridCopyCell.add_Click({
    try {
        if ($null -ne $script:gridMachines.CurrentCell) {
            [System.Windows.Forms.Clipboard]::SetText([string]$script:gridMachines.CurrentCell.Value)
        }
    } catch { Write-LogException -Context 'Open Tailscale admin panel' -ErrorRecord $_ }
})
$gridCopyRow.add_Click({
    try {
        $rowText = Get-SelectedMachineRowText
        if (-not [string]::IsNullOrWhiteSpace($rowText)) {
            [System.Windows.Forms.Clipboard]::SetText($rowText)
        }
    } catch { Write-LogException -Context 'Copy machine row to clipboard' -ErrorRecord $_ }
})
$script:gridMachines.add_CellFormatting({
    param($controlSender,$e)
    try {
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
        $row = $script:gridMachines.Rows[$e.RowIndex]
        $colName = [string]$script:gridMachines.Columns[$e.ColumnIndex].Name
        $machineObj = $row.Tag
        if ($null -eq $machineObj) { return }
        if ([string]$colName -eq 'Connection') {
            $conn = [string](Get-PropertyValue $machineObj @('Connection'))
            if ($conn -eq 'Local') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(234,241,250)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(67,96,147)
            } elseif ($conn -eq 'Direct') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,246,237)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(45,97,65)
            } elseif ($conn -eq 'Relay') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(252,245,220)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(126,96,35)
            } elseif ($conn -eq 'Offline' -or [string](Get-PropertyValue $machineObj @('Status')) -eq 'Offline') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,236,236)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(137,58,58)
            }
            if ($null -ne $e.CellStyle.BackColor) {
                $e.CellStyle.SelectionBackColor = $e.CellStyle.BackColor
                $e.CellStyle.SelectionForeColor = $e.CellStyle.ForeColor
            }
            return
        }
        if ([string]$colName -eq 'LastSeen') {
            $lastSeen = [string]$row.Cells['LastSeen'].Value
            if ($lastSeen -eq 'Connected') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,246,237)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(45,97,65)
            } elseif (-not [string]::IsNullOrWhiteSpace($lastSeen)) {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,236,236)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(137,58,58)
            }
            if ($null -ne $e.CellStyle.BackColor) {
                $e.CellStyle.SelectionBackColor = $e.CellStyle.BackColor
                $e.CellStyle.SelectionForeColor = $e.CellStyle.ForeColor
            }
        }
    } catch {}
})

$machinesLayout.Controls.Add($script:lblMachineHint,0,2)
$machinesLayout.Controls.Add($script:gridMachines,0,3)

$detailRoot = New-Object System.Windows.Forms.TableLayoutPanel
$detailRoot.Dock = 'Fill'
$detailRoot.ColumnCount = 1
$detailRoot.RowCount = 2
$detailRoot.Padding = New-Object System.Windows.Forms.Padding(10,10,10,10)
$null = $detailRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))
$null = $detailRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabMachinesDetails.Controls.Add($detailRoot)

$metricsRoot = New-Object System.Windows.Forms.TableLayoutPanel
$metricsRoot.Dock = 'Fill'
$metricsRoot.ColumnCount = 1
$metricsRoot.RowCount = 3
$metricsRoot.Padding = New-Object System.Windows.Forms.Padding(10,10,10,10)
$null = $metricsRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))
$null = $metricsRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
$null = $metricsRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$diagRoot = New-Object System.Windows.Forms.TableLayoutPanel
$diagRoot.Dock = 'Fill'
$diagRoot.ColumnCount = 1
$diagRoot.RowCount = 5
$diagRoot.Padding = New-Object System.Windows.Forms.Padding(8,0,8,4)
$null = $diagRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))
$null = $diagRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)))
$null = $diagRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))
$null = $diagRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))
$null = $diagRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabMachinesPing.Controls.Add($diagRoot)

$script:lblPingHelp = New-Object System.Windows.Forms.Label
$script:lblPingHelp.Text = ''
$script:lblPingHelp.Dock = 'Fill'
$script:lblPingHelp.Visible = $false
$diagRoot.Controls.Add($script:lblPingHelp,0,0)

$diagSelectionPanel = New-Object System.Windows.Forms.Panel
$diagSelectionPanel.Dock = 'Fill'
$diagSelectionPanel.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
$diagRoot.Controls.Add($diagSelectionPanel,0,1)

$diagSelectionTitle = New-Object System.Windows.Forms.Label
$diagSelectionTitle.Text = 'Selected device'
$diagSelectionTitle.AutoSize = $true
$diagSelectionTitle.Location = New-Object System.Drawing.Point(0,2)
$diagSelectionTitle.Font = New-Object System.Drawing.Font('Segoe UI', 9.25, [System.Drawing.FontStyle]::Bold)
$diagSelectionPanel.Controls.Add($diagSelectionTitle) | Out-Null

$script:lblDiagSelection = New-Object System.Windows.Forms.Label
$script:lblDiagSelection.Text = 'No device selected'
$script:lblDiagSelection.AutoSize = $false
$script:lblDiagSelection.Location = New-Object System.Drawing.Point(118,1)
$script:lblDiagSelection.Size = New-Object System.Drawing.Size(720,22)
$script:lblDiagSelection.TextAlign = 'MiddleLeft'
$diagSelectionPanel.Controls.Add($script:lblDiagSelection) | Out-Null

$grpDiagLocal = New-Object System.Windows.Forms.GroupBox
$grpDiagLocal.Text = 'Network tools'
$grpDiagLocal.Dock = 'Fill'
$grpDiagLocal.Padding = New-Object System.Windows.Forms.Padding(4,2,4,3)
$diagRoot.Controls.Add($grpDiagLocal,0,2)

$diagCommands = New-Object System.Windows.Forms.FlowLayoutPanel
$diagCommands.Dock = 'Top'
$diagCommands.AutoSize = $true
$diagCommands.FlowDirection = 'LeftToRight'
$diagCommands.WrapContents = $true
$diagCommands.AutoScroll = $false
$diagCommands.Margin = New-Object System.Windows.Forms.Padding(0)
$diagCommands.Padding = New-Object System.Windows.Forms.Padding(0)
$grpDiagLocal.Controls.Add($diagCommands)

$script:btnDiagStatus = New-ActionButton -Text 'Status' -Left 0 -Top 0 -Width 96
$script:btnDiagStatus.TextAlign = 'MiddleCenter'
$script:btnDiagNetcheck = New-ActionButton -Text 'Netcheck' -Left 0 -Top 0 -Width 96
$script:btnDiagDns = New-ActionButton -Text 'DNS' -Left 0 -Top 0 -Width 88
$script:btnDiagIPs = New-ActionButton -Text 'IPs' -Left 0 -Top 0 -Width 80
$script:btnDiagMetrics = New-ActionButton -Text 'Metrics' -Left 0 -Top 0 -Width 90
$script:btnAdminPanel = New-ActionButton -Text 'Admin Panel' -Left 0 -Top 0 -Width 116
$script:btnDiagClear = $null
foreach ($btn in @($script:btnDiagStatus,$script:btnDiagNetcheck,$script:btnDiagDns,$script:btnDiagIPs,$script:btnDiagMetrics,$script:btnAdminPanel)) {
    $btn.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $btn.Height = 32
    $diagCommands.Controls.Add($btn) | Out-Null
}

$grpDiagDevice = New-Object System.Windows.Forms.GroupBox
$grpDiagDevice.Text = 'Device actions'
$grpDiagDevice.Dock = 'Fill'
$grpDiagDevice.Padding = New-Object System.Windows.Forms.Padding(4,2,4,3)
$diagRoot.Controls.Add($grpDiagDevice,0,3)

$diagDeviceCommands = New-Object System.Windows.Forms.FlowLayoutPanel
$diagDeviceCommands.Dock = 'Top'
$diagDeviceCommands.AutoSize = $true
$diagDeviceCommands.FlowDirection = 'LeftToRight'
$diagDeviceCommands.WrapContents = $true
$diagDeviceCommands.AutoScroll = $false
$diagDeviceCommands.Margin = New-Object System.Windows.Forms.Padding(0)
$diagDeviceCommands.Padding = New-Object System.Windows.Forms.Padding(0)
$grpDiagDevice.Controls.Add($diagDeviceCommands)

$script:btnCmdPingDns = New-ActionButton -Text 'Ping DNS' -Left 0 -Top 0 -Width 96
$script:btnCmdPingIPv4 = New-ActionButton -Text 'Ping IPv4' -Left 0 -Top 0 -Width 96
$script:btnCmdPingIPv6 = New-ActionButton -Text 'Ping IPv6' -Left 0 -Top 0 -Width 96
$script:btnCmdWhois = New-ActionButton -Text 'Whois' -Left 0 -Top 0 -Width 104
$script:btnCmdPingAll = New-ActionButton -Text 'Ping all' -Left 0 -Top 0 -Width 92
foreach ($btn in @($script:btnCmdPingAll,$script:btnCmdPingDns,$script:btnCmdPingIPv4,$script:btnCmdPingIPv6,$script:btnCmdWhois)) {
    $btn.Margin = New-Object System.Windows.Forms.Padding(0,0,4,0)
    $btn.Height = 32
    $diagDeviceCommands.Controls.Add($btn) | Out-Null
}

$script:txtPingDetails = New-Object System.Windows.Forms.RichTextBox
$script:txtPingDetails.Dock = 'Fill'
$script:txtPingDetails.ReadOnly = $true
$script:txtPingDetails.BorderStyle = 'FixedSingle'
$script:txtPingDetails.ScrollBars = 'Both'
$script:txtPingDetails.Font = New-Object System.Drawing.Font('Consolas', 8.8)
$script:txtPingDetails.WordWrap = $false
$diagRoot.Controls.Add($script:txtPingDetails,0,4)

$dnsResolveRoot = New-Object System.Windows.Forms.TableLayoutPanel
$dnsResolveRoot.Dock = 'Fill'
$dnsResolveRoot.ColumnCount = 1
$dnsResolveRoot.RowCount = 4
$dnsResolveRoot.Padding = New-Object System.Windows.Forms.Padding(10,10,10,10)
$null = $dnsResolveRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$null = $dnsResolveRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
$null = $dnsResolveRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$null = $dnsResolveRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabMachinesDnsResolve.Controls.Add($dnsResolveRoot)

$script:lblDnsResolveHelp = New-Object System.Windows.Forms.Label
$script:lblDnsResolveHelp.Text = 'Resolve a domain using Current, System DNS, Tailscale DNS, or a custom DNS server.'
$script:lblDnsResolveHelp.Dock = 'Fill'
$script:lblDnsResolveHelp.TextAlign = 'MiddleLeft'
$dnsResolveRoot.Controls.Add($script:lblDnsResolveHelp,0,0)

$dnsResolveTopRow = New-Object System.Windows.Forms.TableLayoutPanel
$dnsResolveTopRow.Dock = 'Fill'
$dnsResolveTopRow.ColumnCount = 5
$dnsResolveTopRow.RowCount = 1
$dnsResolveTopRow.Margin = New-Object System.Windows.Forms.Padding(0)
$dnsResolveTopRow.Padding = New-Object System.Windows.Forms.Padding(0)
$null = $dnsResolveTopRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 62)))
$null = $dnsResolveTopRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 188)))
$null = $dnsResolveTopRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 64)))
$null = $dnsResolveTopRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 184)))
$null = $dnsResolveTopRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$dnsResolveRoot.Controls.Add($dnsResolveTopRow,0,1)

$dnsResolveDomainLabel = New-Object System.Windows.Forms.Label
$dnsResolveDomainLabel.Text = 'Domain:'
$dnsResolveDomainLabel.Dock = 'Fill'
$dnsResolveDomainLabel.TextAlign = 'MiddleLeft'
$dnsResolveTopRow.Controls.Add($dnsResolveDomainLabel,0,0)

$script:txtDnsResolveDomain = New-Object System.Windows.Forms.TextBox
$script:txtDnsResolveDomain.Dock = 'Fill'
$script:txtDnsResolveDomain.Margin = New-Object System.Windows.Forms.Padding(0,5,8,0)
$dnsResolveTopRow.Controls.Add($script:txtDnsResolveDomain,1,0)

$dnsResolveResolverLabel = New-Object System.Windows.Forms.Label
$dnsResolveResolverLabel.Text = 'Resolver:'
$dnsResolveResolverLabel.Dock = 'Fill'
$dnsResolveResolverLabel.TextAlign = 'MiddleRight'
$dnsResolveResolverLabel.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$dnsResolveTopRow.Controls.Add($dnsResolveResolverLabel,2,0)

$script:cmbDnsResolveResolver = New-Object System.Windows.Forms.ComboBox
$script:cmbDnsResolveResolver.DropDownStyle = 'DropDownList'
$script:cmbDnsResolveResolver.Dock = 'Fill'
$script:cmbDnsResolveResolver.Margin = New-Object System.Windows.Forms.Padding(0,4,8,0)
$null = $script:cmbDnsResolveResolver.Items.Add('Current')
$null = $script:cmbDnsResolveResolver.Items.Add('System DNS')
$null = $script:cmbDnsResolveResolver.Items.Add('Tailscale DNS')
$null = $script:cmbDnsResolveResolver.Items.Add('Other')
$script:cmbDnsResolveResolver.SelectedIndex = 0
$dnsResolveTopRow.Controls.Add($script:cmbDnsResolveResolver,3,0)

$script:lblDnsResolveServerPreview = New-Object System.Windows.Forms.Label
$script:lblDnsResolveServerPreview.Text = 'Using: -'
$script:lblDnsResolveServerPreview.Dock = 'Fill'
$script:lblDnsResolveServerPreview.TextAlign = 'MiddleLeft'
$script:lblDnsResolveServerPreview.AutoEllipsis = $true
$script:lblDnsResolveServerPreview.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$dnsResolveTopRow.Controls.Add($script:lblDnsResolveServerPreview,4,0)

$dnsResolveSecondRow = New-Object System.Windows.Forms.TableLayoutPanel
$dnsResolveSecondRow.Dock = 'Fill'
$dnsResolveSecondRow.ColumnCount = 8
$dnsResolveSecondRow.RowCount = 1
$dnsResolveSecondRow.Margin = New-Object System.Windows.Forms.Padding(0)
$dnsResolveSecondRow.Padding = New-Object System.Windows.Forms.Padding(0)
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 78)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 58)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 66)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 78)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 102)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
$null = $dnsResolveSecondRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$dnsResolveRoot.Controls.Add($dnsResolveSecondRow,0,2)

$script:lblDnsResolveOther = New-Object System.Windows.Forms.Label
$script:lblDnsResolveOther.Text = 'Other DNS:'
$script:lblDnsResolveOther.Dock = 'Fill'
$script:lblDnsResolveOther.TextAlign = 'MiddleLeft'
$dnsResolveSecondRow.Controls.Add($script:lblDnsResolveOther,0,0)

$script:txtDnsResolveOtherServer = New-Object System.Windows.Forms.TextBox
$script:txtDnsResolveOtherServer.Dock = 'Fill'
$script:txtDnsResolveOtherServer.Text = '1.1.1.1'
$script:txtDnsResolveOtherServer.Margin = New-Object System.Windows.Forms.Padding(0,5,10,0)
$dnsResolveSecondRow.Controls.Add($script:txtDnsResolveOtherServer,1,0)

$dnsResolveCacheLabel = New-Object System.Windows.Forms.Label
$dnsResolveCacheLabel.Text = 'Cache:'
$dnsResolveCacheLabel.Dock = 'Fill'
$dnsResolveCacheLabel.TextAlign = 'MiddleLeft'
$dnsResolveSecondRow.Controls.Add($dnsResolveCacheLabel,2,0)

$script:radDnsResolveUseCache = New-Object System.Windows.Forms.RadioButton
$script:radDnsResolveUseCache.Text = 'Allow'
$script:radDnsResolveUseCache.Dock = 'Fill'
$script:radDnsResolveUseCache.TextAlign = 'MiddleLeft'
$script:radDnsResolveUseCache.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
$dnsResolveSecondRow.Controls.Add($script:radDnsResolveUseCache,3,0)

$script:radDnsResolveNoCache = New-Object System.Windows.Forms.RadioButton
$script:radDnsResolveNoCache.Text = 'Bypass'
$script:radDnsResolveNoCache.Dock = 'Fill'
$script:radDnsResolveNoCache.Checked = $true
$script:radDnsResolveNoCache.TextAlign = 'MiddleLeft'
$script:radDnsResolveNoCache.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
$dnsResolveSecondRow.Controls.Add($script:radDnsResolveNoCache,4,0)

$script:btnDnsResolveRun = New-ActionButton -Text 'Resolve' -Left 0 -Top 0 -Width 100
$script:btnDnsResolveRun.Dock = 'Fill'
$script:btnDnsResolveRun.Margin = New-Object System.Windows.Forms.Padding(0,3,0,3)
$dnsResolveSecondRow.Controls.Add($script:btnDnsResolveRun,5,0)

$dnsResolveSecondHint = New-Object System.Windows.Forms.Label
$dnsResolveSecondHint.Text = ''
$dnsResolveSecondHint.Dock = 'Fill'
$dnsResolveSecondHint.TextAlign = 'MiddleLeft'
$dnsResolveSecondHint.AutoEllipsis = $true
$dnsResolveSecondHint.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
$dnsResolveSecondRow.Controls.Add($dnsResolveSecondHint,7,0)

$script:txtDnsResolveOutput = New-Object System.Windows.Forms.RichTextBox
$script:txtDnsResolveOutput.Dock = 'Fill'
$script:txtDnsResolveOutput.ReadOnly = $true
$script:txtDnsResolveOutput.BorderStyle = 'FixedSingle'
$script:txtDnsResolveOutput.ScrollBars = 'Both'
$script:txtDnsResolveOutput.Font = New-Object System.Drawing.Font('Consolas', 8.8)
$script:txtDnsResolveOutput.WordWrap = $false
$dnsResolveRoot.Controls.Add($script:txtDnsResolveOutput,0,3)

$publicIpRoot = New-Object System.Windows.Forms.TableLayoutPanel
$publicIpRoot.Dock = 'Fill'
$publicIpRoot.ColumnCount = 1
$publicIpRoot.RowCount = 3
$publicIpRoot.Padding = New-Object System.Windows.Forms.Padding(10,10,10,10)
$null = $publicIpRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$null = $publicIpRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
$null = $publicIpRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabMachinesPublicIp.Controls.Add($publicIpRoot)

$script:lblPublicIpHelp = New-Object System.Windows.Forms.Label
$script:lblPublicIpHelp.Text = 'Check the public IP used by the current route. If an exit node is active, this should be the exit node network.'
$script:lblPublicIpHelp.Dock = 'Fill'
$script:lblPublicIpHelp.TextAlign = 'MiddleLeft'
$publicIpRoot.Controls.Add($script:lblPublicIpHelp,0,0)

$publicIpBar = New-Object System.Windows.Forms.FlowLayoutPanel
$publicIpBar.Dock = 'Fill'
$publicIpBar.FlowDirection = 'LeftToRight'
$publicIpBar.WrapContents = $false
$publicIpBar.AutoScroll = $false
$publicIpBar.Padding = New-Object System.Windows.Forms.Padding(0)
$publicIpBar.Margin = New-Object System.Windows.Forms.Padding(0)
$publicIpRoot.Controls.Add($publicIpBar,0,1)

$script:radPublicIpFast = New-Object System.Windows.Forms.RadioButton
$script:radPublicIpFast.Text = 'Fast'
$script:radPublicIpFast.Checked = $true
$script:radPublicIpFast.AutoSize = $true
$script:radPublicIpFast.Margin = New-Object System.Windows.Forms.Padding(8,9,14,0)
$publicIpBar.Controls.Add($script:radPublicIpFast) | Out-Null

$script:radPublicIpDetailed = New-Object System.Windows.Forms.RadioButton
$script:radPublicIpDetailed.Text = 'Detailed'
$script:radPublicIpDetailed.AutoSize = $true
$script:radPublicIpDetailed.Margin = New-Object System.Windows.Forms.Padding(0,9,18,0)
$publicIpBar.Controls.Add($script:radPublicIpDetailed) | Out-Null

$script:btnPublicIpRun = New-ActionButton -Text 'Test public IP' -Left 0 -Top 0 -Width 140
$script:btnPublicIpRun.Height = 32
$script:btnPublicIpRun.Margin = New-Object System.Windows.Forms.Padding(0,2,8,0)
$publicIpBar.Controls.Add($script:btnPublicIpRun) | Out-Null

$script:txtPublicIpOutput = New-Object System.Windows.Forms.RichTextBox
$script:txtPublicIpOutput.Dock = 'Fill'
$script:txtPublicIpOutput.ReadOnly = $true
$script:txtPublicIpOutput.BorderStyle = 'FixedSingle'
$script:txtPublicIpOutput.ScrollBars = 'Both'
$script:txtPublicIpOutput.Font = New-Object System.Drawing.Font('Consolas', 8.8)
$script:txtPublicIpOutput.WordWrap = $false
$publicIpRoot.Controls.Add($script:txtPublicIpOutput,0,2)

$script:lblPingNote = New-Object System.Windows.Forms.Label
$script:lblPingNote.Text = 'Network output appears here.'
$script:lblPingNote.Visible = $false

$script:gridPing = New-Object System.Windows.Forms.DataGridView
$script:gridPing.ReadOnly = $true
$script:gridPing.RowHeadersVisible = $false
$script:gridPing.AllowUserToAddRows = $false
$script:gridPing.AllowUserToDeleteRows = $false
$script:gridPing.AllowUserToResizeRows = $false
$script:gridPing.SelectionMode = 'FullRowSelect'
$script:gridPing.MultiSelect = $false
$script:gridPing.AutoSizeColumnsMode = 'None'
$script:gridPing.AutoGenerateColumns = $false
$script:gridPing.BorderStyle = 'FixedSingle'
$script:gridPing.ColumnHeadersHeight = 32
$script:gridPing.RowTemplate.Height = 26
$script:gridPing.ScrollBars = 'Both'
$script:gridPing.ClipboardCopyMode = 'EnableWithoutHeaderText'
$script:gridPing.Columns.Clear()

$pingColumns = @(
    @{ Name='PingDevice'; Header='Device'; Width=132 },
    @{ Name='PingOwner'; Header='Owner'; Width=84 },
    @{ Name='PingOnline'; Header='Online'; Width=62 },
    @{ Name='PingLastSeen'; Header='Last Seen'; Width=96 },
    @{ Name='PingDnsTarget'; Header='Target'; Width=176 },
    @{ Name='PingIPv4Target'; Header='IPv4'; Width=106 },
    @{ Name='PingIPv6Target'; Header='IPv6'; Width=176 },
    @{ Name='PingPath'; Header='Conn'; Width=64 },
    @{ Name='PingDerp'; Header='DERP'; Width=84 },
    @{ Name='PingDnsMs'; Header='DNS ms'; Width=68 },
    @{ Name='PingIPv4Ms'; Header='IPv4 ms'; Width=70 },
    @{ Name='PingIPv6Ms'; Header='IPv6 ms'; Width=70 }
)
foreach ($col in $pingColumns) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = [string]$col.Name
    $c.HeaderText = [string]$col.Header
    $c.Width = [int]$col.Width
    $c.MinimumWidth = [Math]::Max(58, [int]([math]::Floor($c.Width * 0.75)))
    $c.SortMode = 'NotSortable'
    $null = $script:gridPing.Columns.Add($c)
}
$script:gridPing.Columns['PingOnline'].DefaultCellStyle.Alignment = 'MiddleCenter'
$script:gridPing.Columns['PingPath'].DefaultCellStyle.Alignment = 'MiddleCenter'
$script:gridPing.Columns['PingDerp'].DefaultCellStyle.Alignment = 'MiddleCenter'
$script:gridPing.Columns['PingDnsMs'].DefaultCellStyle.Alignment = 'MiddleRight'
$script:gridPing.Columns['PingIPv4Ms'].DefaultCellStyle.Alignment = 'MiddleRight'
$script:gridPing.Columns['PingIPv6Ms'].DefaultCellStyle.Alignment = 'MiddleRight'

$detailHelp = New-Object System.Windows.Forms.Label
$detailHelp.Text = ''
$detailHelp.Dock = 'Fill'
$detailHelp.Visible = $false
$detailRoot.Controls.Add($detailHelp,0,0)

$script:txtMachineDetails = New-Object System.Windows.Forms.RichTextBox
$script:txtMachineDetails.Dock = 'Fill'
$script:txtMachineDetails.ReadOnly = $true
$script:txtMachineDetails.DetectUrls = $false
$script:txtMachineDetails.BorderStyle = 'FixedSingle'
$script:txtMachineDetails.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$script:txtMachineDetails.WordWrap = $true
$detailRoot.Controls.Add($script:txtMachineDetails,0,1)

$detailMetricsBar = New-Object System.Windows.Forms.FlowLayoutPanel
$detailMetricsBar.Dock = 'Fill'
$detailMetricsBar.FlowDirection = 'LeftToRight'
$detailMetricsBar.WrapContents = $false
$detailMetricsBar.AutoScroll = $false
$metricsRoot.Controls.Add($detailMetricsBar,0,1)

$script:lblMetricsInfo = New-Object System.Windows.Forms.Label
$script:lblMetricsInfo.Text = 'Load and clear a structured summary of local Tailscale metrics.'
$script:lblMetricsInfo.AutoSize = $true
$script:lblMetricsInfo.Margin = New-Object System.Windows.Forms.Padding(0,10,14,0)
$detailMetricsBar.Controls.Add($script:lblMetricsInfo) | Out-Null

$script:btnDetailMetrics = New-ActionButton -Text 'Load metrics' -Left 0 -Top 0 -Width 156
$script:btnDetailMetrics.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$detailMetricsBar.Controls.Add($script:btnDetailMetrics) | Out-Null

$script:btnDetailClearMetrics = New-ActionButton -Text 'Clear metrics' -Left 0 -Top 0 -Width 156
$script:btnDetailClearMetrics.Margin = New-Object System.Windows.Forms.Padding(0)
$detailMetricsBar.Controls.Add($script:btnDetailClearMetrics) | Out-Null

$script:txtMetricsSummary = New-Object System.Windows.Forms.RichTextBox
$script:txtMetricsSummary.Dock = 'Fill'
$script:txtMetricsSummary.ReadOnly = $true
$script:txtMetricsSummary.BorderStyle = 'FixedSingle'
$script:txtMetricsSummary.ScrollBars = 'Both'
$script:txtMetricsSummary.Font = New-Object System.Drawing.Font('Consolas', 8.8)
$script:txtMetricsSummary.WordWrap = $false
$script:txtMetricsSummary.MaxLength = [int]$script:MetricsTextMaxChars
$metricsRoot.Controls.Add($script:txtMetricsSummary,0,2)

$script:gridPing.add_CellFormatting({
    param($controlSender,$e)
    try {
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
        $row = $script:gridPing.Rows[$e.RowIndex]
        $colName = [string]$script:gridPing.Columns[$e.ColumnIndex].Name
        $tag = $row.Tag
        if ($null -eq $tag) { return }
        if ($colName -eq 'PingPath') {
            $path = [string]$row.Cells['PingPath'].Value
            if ($path -eq 'Local') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(234,241,250)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(67,96,147)
            }
            elseif ($path -eq 'Direct') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,246,237)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(45,97,65)
            }
            elseif ($path -eq 'Relay') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(252,245,220)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(126,96,35)
            }
            elseif ($path -eq 'Offline' -or $path -eq 'Failed') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,236,236)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(137,58,58)
            }
            if ($null -ne $e.CellStyle.BackColor) {
                $e.CellStyle.SelectionBackColor = $e.CellStyle.BackColor
                $e.CellStyle.SelectionForeColor = $e.CellStyle.ForeColor
            }
        }
        elseif ($colName -eq 'PingOnline') {
            $online = [string]$row.Cells['PingOnline'].Value
            if ($online -eq 'Yes') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,246,237)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(45,97,65)
            }
            elseif ($online -eq 'No') {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,236,236)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(137,58,58)
            }
            if ($null -ne $e.CellStyle.BackColor) {
                $e.CellStyle.SelectionBackColor = $e.CellStyle.BackColor
                $e.CellStyle.SelectionForeColor = $e.CellStyle.ForeColor
            }
        }
    } catch {}
})

$script:grpAccount = New-Object System.Windows.Forms.GroupBox
$script:grpAccount.Text = 'Account'
$script:grpAccount.Dock = 'Fill'
$script:grpAccount.Margin = New-Object System.Windows.Forms.Padding(0)
$script:grpAccount.BackColor = [System.Drawing.Color]::White
$script:tabAccount.AutoScroll = $true
$script:tabAccount.BackColor = [System.Drawing.Color]::White
$script:tabAccount.Controls.Add($script:grpAccount)

$accountLayout = New-Object System.Windows.Forms.TableLayoutPanel
$accountLayout.Dock = 'Fill'
$accountLayout.ColumnCount = 1
$accountLayout.RowCount = 4
$accountLayout.Padding = New-Object System.Windows.Forms.Padding(10,8,10,8)
$accountLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$accountLayout.BackColor = [System.Drawing.Color]::White
$null = $accountLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174)))
$null = $accountLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$null = $accountLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $accountLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$script:grpAccount.Controls.Add($accountLayout)

$accountDetails = New-Object System.Windows.Forms.TableLayoutPanel
$accountDetails.Dock = 'Fill'
$accountDetails.ColumnCount = 4
$accountDetails.RowCount = 7
$accountDetails.Margin = New-Object System.Windows.Forms.Padding(0)
$accountDetails.Padding = New-Object System.Windows.Forms.Padding(12,8,12,8)
$accountDetails.BackColor = [System.Drawing.Color]::White
$null = $accountDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 116)))
$null = $accountDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$null = $accountDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 116)))
$null = $accountDetails.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($i = 0; $i -lt 7; $i++) { $null = $accountDetails.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) }
$accountLayout.Controls.Add($accountDetails,0,0)

function Add-AccountDetailPair {
    param([int]$Row,[int]$LabelCol,[string]$LabelText,[string]$ValueName)
    $lbl = New-Object TailscaleControlSummaryLabel
    $lbl.Text = $LabelText
    $lbl.Font = $script:SummaryLabelFont
    Set-SummarySingleLineLabel -Label $lbl -Ellipsis $false
    $accountDetails.Controls.Add($lbl,$LabelCol,$Row) | Out-Null

    $val = New-Object TailscaleControlSummaryLabel
    $val.Text = '-'
    $val.Font = $script:SummaryValueFont
    Set-SummarySingleLineLabel -Label $val -Ellipsis $true
    $accountDetails.Controls.Add($val,($LabelCol + 1),$Row) | Out-Null
    Set-Variable -Scope Script -Name $ValueName -Value $val
}

Add-AccountDetailPair -Row 0 -LabelCol 0 -LabelText 'Identifier:' -ValueName 'lblAccountIdentifier'
Add-AccountDetailPair -Row 1 -LabelCol 0 -LabelText 'Email:' -ValueName 'lblAccountEmail'
Add-AccountDetailPair -Row 2 -LabelCol 0 -LabelText 'Active User:' -ValueName 'lblAccountActiveUser'
Add-AccountDetailPair -Row 3 -LabelCol 0 -LabelText 'Tailnet:' -ValueName 'lblAccountTailnet'
Add-AccountDetailPair -Row 4 -LabelCol 0 -LabelText 'Device:' -ValueName 'lblAccountDevice'
Add-AccountDetailPair -Row 5 -LabelCol 0 -LabelText 'MagicDNS:' -ValueName 'lblAccountDnsName'
Add-AccountDetailPair -Row 0 -LabelCol 2 -LabelText 'Visible Devices:' -ValueName 'lblAccountVisibleDevices'
Add-AccountDetailPair -Row 1 -LabelCol 2 -LabelText 'Visible Users:' -ValueName 'lblAccountVisibleUsers'

$accountButtonBar = New-Object System.Windows.Forms.TableLayoutPanel
$accountButtonBar.Dock = 'Fill'
$accountButtonBar.ColumnCount = 2
$accountButtonBar.RowCount = 1
$accountButtonBar.Margin = New-Object System.Windows.Forms.Padding(0)
$accountButtonBar.Padding = New-Object System.Windows.Forms.Padding(0,2,0,2)
$accountButtonBar.BackColor = [System.Drawing.Color]::White
$null = $accountButtonBar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$null = $accountButtonBar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$null = $accountButtonBar.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$accountLayout.Controls.Add($accountButtonBar,0,1)

$script:lblAccountLoggedAccounts = New-Object System.Windows.Forms.Label
$script:lblAccountLoggedAccounts.Text = 'Logged Accounts: -'
$script:lblAccountLoggedAccounts.Tag = 'InlineLoggedAccounts'
$script:lblAccountLoggedAccounts.Dock = 'Fill'
$script:lblAccountLoggedAccounts.AutoSize = $false
$script:lblAccountLoggedAccounts.TextAlign = 'MiddleLeft'
$script:lblAccountLoggedAccounts.Font = $script:SummaryValueFont
$script:lblAccountLoggedAccounts.ForeColor = [System.Drawing.Color]::FromArgb(92,102,116)
$script:lblAccountLoggedAccounts.Margin = New-Object System.Windows.Forms.Padding(0)
$accountButtonBar.Controls.Add($script:lblAccountLoggedAccounts,0,0) | Out-Null

$accountButtons = New-FlowButtonRow
$accountButtons.Dock = 'Fill'
$accountButtons.Anchor = 'Right'
$accountButtons.FlowDirection = 'RightToLeft'
$accountButtons.Padding = New-Object System.Windows.Forms.Padding(0)
$accountButtons.Margin = New-Object System.Windows.Forms.Padding(0)
$accountButtons.BackColor = [System.Drawing.Color]::White
$accountButtonBar.Controls.Add($accountButtons,1,0) | Out-Null

$script:btnAccountLogout = New-ActionButton -Text 'Logout' -Left 0 -Top 0 -Width 92
$script:btnAccountLogout.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
$accountButtons.Controls.Add($script:btnAccountLogout) | Out-Null

$script:btnAccountAdd = New-ActionButton -Text 'Add Another Account' -Left 0 -Top 0 -Width 166
$script:btnAccountAdd.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
$accountButtons.Controls.Add($script:btnAccountAdd) | Out-Null

$script:btnAccountSwitch = New-ActionButton -Text 'Switch Selected' -Left 0 -Top 0 -Width 142
$script:btnAccountSwitch.Margin = New-Object System.Windows.Forms.Padding(0)
$accountButtons.Controls.Add($script:btnAccountSwitch) | Out-Null

$script:gridAccounts = New-Object System.Windows.Forms.DataGridView
$script:gridAccounts.Dock = 'Fill'
$script:gridAccounts.ReadOnly = $true
$script:gridAccounts.RowHeadersVisible = $false
$script:gridAccounts.AllowUserToAddRows = $false
$script:gridAccounts.AllowUserToDeleteRows = $false
$script:gridAccounts.AllowUserToResizeRows = $false
$script:gridAccounts.SelectionMode = 'FullRowSelect'
$script:gridAccounts.MultiSelect = $false
$script:gridAccounts.AutoSizeColumnsMode = 'Fill'
$script:gridAccounts.AutoGenerateColumns = $false
$script:gridAccounts.BorderStyle = 'FixedSingle'
$script:gridAccounts.BackgroundColor = [System.Drawing.Color]::White
$script:gridAccounts.GridColor = [System.Drawing.Color]::Gainsboro
$script:gridAccounts.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$script:gridAccounts.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,43,54)
$script:gridAccounts.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(232,241,255)
$script:gridAccounts.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,43,54)
$script:gridAccounts.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$script:gridAccounts.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(92,102,116)
$script:gridAccounts.EnableHeadersVisualStyles = $false
$script:gridAccounts.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,250,252)
$script:gridAccounts.ColumnHeadersHeight = 32
$script:gridAccounts.RowTemplate.Height = 28
$script:gridAccounts.ScrollBars = 'Both'
$script:gridAccounts.Columns.Clear()
$accountColumns = @(
    @{ Name='AccountStatus'; Header='Status'; Width=120; Fill=18 },
    @{ Name='AccountIdentifier'; Header='Identifier'; Width=330; Fill=42 },
    @{ Name='AccountUser'; Header='User'; Width=300; Fill=40 }
)
foreach ($col in $accountColumns) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = [string]$col.Name
    $c.HeaderText = [string]$col.Header
    $c.Width = [int]$col.Width
    $c.MinimumWidth = [Math]::Max(56,[int]([int]$col.Width * 0.7))
    $c.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $c.FillWeight = [int]$col.Fill
    $c.Resizable = [System.Windows.Forms.DataGridViewTriState]::True
    $c.SortMode = 'NotSortable'
    $null = $script:gridAccounts.Columns.Add($c)
}
$accountLayout.Controls.Add($script:gridAccounts,0,2)

$script:lblAccountListStatus = New-Object System.Windows.Forms.Label
$script:lblAccountListStatus.Text = 'Double-click a disconnected account to switch.'
$script:lblAccountListStatus.Dock = 'Fill'
$script:lblAccountListStatus.TextAlign = 'MiddleLeft'
$script:lblAccountListStatus.Font = $script:SummaryLabelFont
$accountLayout.Controls.Add($script:lblAccountListStatus,0,3)

$script:grpMaintenance = New-Object System.Windows.Forms.GroupBox
$script:grpMaintenance.Text = 'Maintenance'
$script:grpMaintenance.Dock = 'Fill'
$script:grpMaintenance.Margin = New-Object System.Windows.Forms.Padding(0)
$tabMaint.AutoScroll = $true
$tabMaint.Controls.Add($script:grpMaintenance)

$maintLayout = New-Object System.Windows.Forms.TableLayoutPanel
$maintLayout.Dock = 'Fill'
$maintLayout.ColumnCount = 1
$maintLayout.RowCount = 1
$maintLayout.Padding = New-Object System.Windows.Forms.Padding(10,6,10,6)
$null = $maintLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:grpMaintenance.Controls.Add($maintLayout)

$maintTabs = New-Object System.Windows.Forms.TabControl
$maintTabs.Dock = 'Fill'
$maintTabs.Margin = New-Object System.Windows.Forms.Padding(0)
$maintTabs.Padding = New-Object System.Drawing.Point(8,6)
$maintTabs.DrawMode = 'Normal'
$maintTabs.Appearance = 'Normal'
$maintTabs.Multiline = $false
$maintTabs.SizeMode = 'Fixed'
$maintTabs.ItemSize = New-Object System.Drawing.Size(120,28)
$maintTabs.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$maintLayout.Controls.Add($maintTabs,0,0)

$tabUpdateInner = New-Object System.Windows.Forms.TabPage
$tabUpdateInner.Text = 'Tailscale'
$tabMtuInner = New-Object System.Windows.Forms.TabPage
$tabMtuInner.Text = 'Tailscale MTU'
$tabControlInner = New-Object System.Windows.Forms.TabPage
$tabControlInner.Text = 'Tailscale Control'
$null = $maintTabs.TabPages.Add($tabUpdateInner)
$null = $maintTabs.TabPages.Add($tabControlInner)
$null = $maintTabs.TabPages.Add($tabMtuInner)

$tabUpdateInner.AutoScroll = $false
$updateScroll = New-Object System.Windows.Forms.Panel
$updateScroll.Dock = 'Fill'
$updateScroll.AutoScroll = $true
try { $updateScroll.HorizontalScroll.Enabled = $false; $updateScroll.HorizontalScroll.Visible = $false } catch { }
$updateScroll.Padding = New-Object System.Windows.Forms.Padding(0)
$tabUpdateInner.Controls.Add($updateScroll)

$updateLayout = New-Object System.Windows.Forms.TableLayoutPanel
$updateLayout.Dock = 'Top'
$updateLayout.AutoSize = $true
$updateLayout.ColumnCount = 2
$updateLayout.RowCount = 8
$updateLayout.Padding = New-Object System.Windows.Forms.Padding(12,12,12,12)
$updateLayout.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$null = $updateLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))
$null = $updateLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
for ($i = 0; $i -lt 5; $i++) { $null = $updateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) }
$null = $updateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
$null = $updateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
$null = $updateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$updateScroll.Controls.Add($updateLayout)
function New-MaintenanceHeaderIconBox {
    param([string]$FileName)
    $box = New-Object System.Windows.Forms.PictureBox
    $box.Dock = 'Right'
    $box.Size = New-Object System.Drawing.Size(26,26)
    $box.Margin = New-Object System.Windows.Forms.Padding(8,1,0,1)
    $box.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
    try {
        $img = Get-CachedIconBitmap -FileName $FileName -Size 24
        if ($null -ne $img) { $box.Image = $img }
    }
    catch { }
    return $box
}

function Add-MaintenanceCornerIcon {
    param($Parent,[string]$FileName)
    try {
        if ($null -eq $Parent -or [string]::IsNullOrWhiteSpace([string]$FileName)) { return $null }
        $box = New-Object System.Windows.Forms.PictureBox
        $box.Size = New-Object System.Drawing.Size(76,76)
        $box.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
        $box.BackColor = [System.Drawing.Color]::Transparent
        $box.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $box.Margin = New-Object System.Windows.Forms.Padding(0)
        try {
            $img = Get-CachedIconBitmap -FileName $FileName -Size 64
            if ($null -ne $img) { $box.Image = $img }
        } catch { }
        $place = {
            param($sender,$eventArgs)
            try {
                if ($null -eq $box -or $box.IsDisposed -or $null -eq $Parent) { return }
                $x = [Math]::Max(12, [int]($Parent.ClientSize.Width - $box.Width - 18))
                $box.Location = New-Object System.Drawing.Point($x, 12)
                $box.BringToFront()
            } catch { }
        }.GetNewClosure()
        $Parent.Controls.Add($box) | Out-Null
        & $place $null $null
        $Parent.add_Resize($place)
        return $box
    }
    catch { return $null }
}

function Add-MaintValuePair {
    param($Layout,[int]$Row,[string]$LabelText,[string]$ValueName,[string]$IconFileName = '')
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Dock = 'Fill'
    $lbl.TextAlign = 'MiddleLeft'
    $lbl.Font = $script:SummaryLabelFont
    $Layout.Controls.Add($lbl,0,$Row) | Out-Null

    $val = New-Object System.Windows.Forms.Label
    $val.Text = '-'
    $val.Dock = 'Fill'
    $val.TextAlign = 'MiddleLeft'
    $val.AutoEllipsis = $true

    if (-not [string]::IsNullOrWhiteSpace([string]$IconFileName)) {
        $hostPanel = New-Object System.Windows.Forms.TableLayoutPanel
        $hostPanel.Dock = 'Fill'
        $hostPanel.Margin = New-Object System.Windows.Forms.Padding(0)
        $hostPanel.Padding = New-Object System.Windows.Forms.Padding(0)
        $hostPanel.ColumnCount = 2
        $hostPanel.RowCount = 1
        $null = $hostPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
        $null = $hostPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,34)))
        $hostPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
        $hostPanel.Controls.Add($val,0,0) | Out-Null
        $hostPanel.Controls.Add((New-MaintenanceHeaderIconBox -FileName $IconFileName),1,0) | Out-Null
        $Layout.Controls.Add($hostPanel,1,$Row) | Out-Null
    }
    else {
        $Layout.Controls.Add($val,1,$Row) | Out-Null
    }

    Set-Variable -Scope Script -Name $ValueName -Value $val
}
[void](Add-MaintenanceCornerIcon -Parent $updateScroll -FileName 'tailscale.ico')
Add-MaintValuePair -Layout $updateLayout -Row 0 -LabelText 'Current version:' -ValueName 'lblMaintLocalVersion'
Add-MaintValuePair -Layout $updateLayout -Row 1 -LabelText 'Latest version:' -ValueName 'lblMaintLatestVersion'
Add-MaintValuePair -Layout $updateLayout -Row 2 -LabelText 'Last check:' -ValueName 'lblMaintLastCheck'
Add-MaintValuePair -Layout $updateLayout -Row 3 -LabelText 'Update mode:' -ValueName 'lblMaintAutoUpdate'
Add-MaintValuePair -Layout $updateLayout -Row 4 -LabelText 'Last update result:' -ValueName 'lblMaintUpdateStatus'
$updateButtons = New-FlowButtonRow
$updateButtons.Margin = New-Object System.Windows.Forms.Padding(0,6,0,6)
$updateButtons.MinimumSize = New-Object System.Drawing.Size(0, 34)
$updateLayout.Controls.Add($updateButtons,0,5)
$updateLayout.SetColumnSpan($updateButtons,2)

$script:btnInstallClientAutoUpdateTask = New-ActionButton -Text 'Install Auto Update Task' -Left 0 -Top 0 -Width 190
$script:btnInstallClientAutoUpdateTask.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$updateButtons.Controls.Add($script:btnInstallClientAutoUpdateTask) | Out-Null

$script:btnCheckClientUpdate = New-ActionButton -Text 'Check Update' -Left 0 -Top 0 -Width 128
$script:btnCheckClientUpdate.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$updateButtons.Controls.Add($script:btnCheckClientUpdate) | Out-Null

$script:btnRunClientUpdate = New-ActionButton -Text 'Update' -Left 0 -Top 0 -Width 96
$script:btnRunClientUpdate.Enabled = $false
$script:btnRunClientUpdate.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$updateButtons.Controls.Add($script:btnRunClientUpdate) | Out-Null

$clientAutoRow = New-FlowButtonRow
$clientAutoRow.Dock = 'Fill'
$clientAutoRow.MinimumSize = New-Object System.Drawing.Size(0, 34)
$clientAutoRow.Margin = New-Object System.Windows.Forms.Padding(0,3,0,3)

$script:chkCheckUpdateEvery = New-Object System.Windows.Forms.CheckBox
$script:chkCheckUpdateEvery.Text = 'Auto check/update every'
$script:chkCheckUpdateEvery.AutoSize = $true
$script:chkCheckUpdateEvery.Margin = New-Object System.Windows.Forms.Padding(0,7,8,0)
$clientAutoRow.Controls.Add($script:chkCheckUpdateEvery) | Out-Null

$script:numCheckUpdateHours = New-Object System.Windows.Forms.NumericUpDown
$script:numCheckUpdateHours.Minimum = 1
$script:numCheckUpdateHours.Maximum = 168
$script:numCheckUpdateHours.Value = 24
$script:numCheckUpdateHours.Width = 58
$script:numCheckUpdateHours.Margin = New-Object System.Windows.Forms.Padding(0,4,6,0)
$clientAutoRow.Controls.Add($script:numCheckUpdateHours) | Out-Null

$script:lblCheckUpdateHours = New-Object System.Windows.Forms.Label
$script:lblCheckUpdateHours.Text = 'hours'
$script:lblCheckUpdateHours.AutoSize = $true
$script:lblCheckUpdateHours.Margin = New-Object System.Windows.Forms.Padding(0,8,0,0)
$clientAutoRow.Controls.Add($script:lblCheckUpdateHours) | Out-Null

$updateLayout.Controls.Add($clientAutoRow,0,6) | Out-Null
$updateLayout.SetColumnSpan($clientAutoRow,2) | Out-Null

$tabMtuInner.AutoScroll = $false
$mtuScroll = New-Object System.Windows.Forms.Panel
$mtuScroll.Dock = 'Fill'
$mtuScroll.AutoScroll = $true
try { $mtuScroll.HorizontalScroll.Enabled = $false; $mtuScroll.HorizontalScroll.Visible = $false } catch { }
$mtuScroll.Padding = New-Object System.Windows.Forms.Padding(0)
$tabMtuInner.Controls.Add($mtuScroll)

$mtuLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mtuLayout.Dock = 'Top'
$mtuLayout.AutoSize = $true
$mtuLayout.ColumnCount = 2
$mtuLayout.RowCount = 11
$mtuLayout.Padding = New-Object System.Windows.Forms.Padding(12,12,12,12)
$mtuLayout.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$null = $mtuLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))
$null = $mtuLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
for ($i = 0; $i -lt 10; $i++) { $null = $mtuLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) }
$null = $mtuLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$mtuScroll.Controls.Add($mtuLayout)
[void](Add-MaintenanceCornerIcon -Parent $mtuScroll -FileName 'tailscale-mtu.ico')

Add-MaintValuePair -Layout $mtuLayout -Row 0 -LabelText 'Status:' -ValueName 'lblMaintMtuStatus'
Add-MaintValuePair -Layout $mtuLayout -Row 1 -LabelText 'Current version:' -ValueName 'lblMaintMtuVersion'
Add-MaintValuePair -Layout $mtuLayout -Row 2 -LabelText 'Service:' -ValueName 'lblMaintMtuService'
Add-MaintValuePair -Layout $mtuLayout -Row 3 -LabelText 'Desired IPv4:' -ValueName 'lblMaintMtuDesiredIPv4'
Add-MaintValuePair -Layout $mtuLayout -Row 4 -LabelText 'Desired IPv6:' -ValueName 'lblMaintMtuDesiredIPv6'
Add-MaintValuePair -Layout $mtuLayout -Row 5 -LabelText 'Check every:' -ValueName 'lblMaintMtuCheckInterval'
Add-MaintValuePair -Layout $mtuLayout -Row 6 -LabelText 'Last result:' -ValueName 'lblMaintMtuLastResult'
Add-MaintValuePair -Layout $mtuLayout -Row 7 -LabelText 'Last error:' -ValueName 'lblMaintMtuLastError'
Add-MaintValuePair -Layout $mtuLayout -Row 8 -LabelText 'Repository:' -ValueName 'lblMaintMtuRepo'
Add-MaintValuePair -Layout $mtuLayout -Row 9 -LabelText 'By:' -ValueName 'lblMaintMtuAuthor'

$script:lblMaintMtuRepo.Text = 'https://github.com/luizbizzio/tailscale-mtu'
$script:lblMaintMtuAuthor.Text = 'Luiz Bizzio'

$mtuButtons = New-FlowButtonRow
$mtuLayout.Controls.Add($mtuButtons,0,10)
$mtuLayout.SetColumnSpan($mtuButtons,2)

$script:btnInstallMtu = New-ActionButton -Text 'Install Tailscale MTU' -Left 0 -Top 0 -Width 160
$script:btnOpenMtu = New-ActionButton -Text 'Open Tailscale MTU' -Left 0 -Top 0 -Width 160
$script:btnCheckMtuRepo = New-ActionButton -Text 'Check Repo' -Left 0 -Top 0 -Width 116
$script:btnInstallMtu.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnOpenMtu.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnCheckMtuRepo.Margin = New-Object System.Windows.Forms.Padding(0)
$script:btnCheckMtuRepo.TabStop = $false
$mtuButtons.Controls.Add($script:btnInstallMtu) | Out-Null
$mtuButtons.Controls.Add($script:btnOpenMtu) | Out-Null
$mtuButtons.Controls.Add($script:btnCheckMtuRepo) | Out-Null

$tabControlInner.AutoScroll = $false
$controlScroll = New-Object System.Windows.Forms.Panel
$controlScroll.Dock = 'Fill'
$controlScroll.AutoScroll = $true
try { $controlScroll.HorizontalScroll.Enabled = $false; $controlScroll.HorizontalScroll.Visible = $false } catch { }
$controlScroll.Padding = New-Object System.Windows.Forms.Padding(0)
$tabControlInner.Controls.Add($controlScroll)

$controlLayout = New-Object System.Windows.Forms.TableLayoutPanel
$controlLayout.Dock = 'Top'
$controlLayout.AutoSize = $true
$controlLayout.ColumnCount = 2
$controlLayout.RowCount = 11
$controlLayout.Padding = New-Object System.Windows.Forms.Padding(12,12,12,12)
$controlLayout.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$null = $controlLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))
$null = $controlLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
for ($i = 0; $i -lt 8; $i++) { $null = $controlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) }
$null = $controlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$null = $controlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$null = $controlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
$controlScroll.Controls.Add($controlLayout)
[void](Add-MaintenanceCornerIcon -Parent $controlScroll -FileName 'tailscale-control.ico')

Add-MaintValuePair -Layout $controlLayout -Row 0 -LabelText 'Current version:' -ValueName 'lblControlVersion'
Add-MaintValuePair -Layout $controlLayout -Row 1 -LabelText 'Latest version:' -ValueName 'lblControlLatestVersion'
Add-MaintValuePair -Layout $controlLayout -Row 2 -LabelText 'Last check:' -ValueName 'lblControlLastCheck'
Add-MaintValuePair -Layout $controlLayout -Row 3 -LabelText 'Auto update:' -ValueName 'lblControlAutoUpdate'
Add-MaintValuePair -Layout $controlLayout -Row 4 -LabelText 'Last update result:' -ValueName 'lblControlUpdateStatus'
Add-MaintValuePair -Layout $controlLayout -Row 5 -LabelText 'Open path:' -ValueName 'lblControlPath'
Add-MaintValuePair -Layout $controlLayout -Row 6 -LabelText 'Repository:' -ValueName 'lblControlRepo'
Add-MaintValuePair -Layout $controlLayout -Row 7 -LabelText 'By:' -ValueName 'lblControlAuthor'

$script:lblControlVersion.Text = $script:AppVersion
$script:lblControlLatestVersion.Text = '-'
$script:lblControlLastCheck.Text = '-'
$script:lblControlAutoUpdate.Text = 'Off'
$script:lblControlUpdateStatus.Text = '-'
$script:lblControlPath.Text = $script:AppRoot
$script:lblControlRepo.Text = 'https://github.com/luizbizzio/tailscale-control'
$script:lblControlAuthor.Text = 'Luiz Bizzio'

$controlButtons = New-FlowButtonRow
$controlButtons.WrapContents = $true
$controlButtons.MinimumSize = New-Object System.Drawing.Size(0, 42)
$script:btnUpdate.Width = 110
$script:btnUpdate.Enabled = $false
$script:btnCheckControlUpdate = New-ActionButton -Text 'Check Update' -Left 0 -Top 0 -Width 118
$script:btnCheckControlRepo = New-ActionButton -Text 'Check Repo' -Left 0 -Top 0 -Width 110
$script:btnOpenControlPath = New-ActionButton -Text 'Open Path' -Left 0 -Top 0 -Width 104
$script:btnUninstall.Width = 110
$script:btnCheckControlUpdate.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnUpdate.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnCheckControlRepo.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnOpenControlPath.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
$script:btnCheckControlRepo.TabStop = $false
$script:btnOpenControlPath.TabStop = $false
$script:btnUninstall.Margin = New-Object System.Windows.Forms.Padding(0)
$controlButtons.Controls.Add($script:btnCheckControlUpdate) | Out-Null
$controlButtons.Controls.Add($script:btnUpdate) | Out-Null
$controlButtons.Controls.Add($script:btnCheckControlRepo) | Out-Null
$controlButtons.Controls.Add($script:btnOpenControlPath) | Out-Null
$controlButtons.Controls.Add($script:btnUninstall) | Out-Null
$controlLayout.Controls.Add($controlButtons,0,8) | Out-Null
$controlLayout.SetColumnSpan($controlButtons,2) | Out-Null

$controlAutoRow = New-FlowButtonRow
$controlAutoRow.MinimumSize = New-Object System.Drawing.Size(0, 34)

$script:chkControlCheckUpdateEvery = New-Object System.Windows.Forms.CheckBox
$script:chkControlCheckUpdateEvery.Text = 'Auto check/update every'
$script:chkControlCheckUpdateEvery.AutoSize = $true
$script:chkControlCheckUpdateEvery.Margin = New-Object System.Windows.Forms.Padding(0,6,8,0)
$controlAutoRow.Controls.Add($script:chkControlCheckUpdateEvery) | Out-Null

$script:numControlCheckUpdateHours = New-Object System.Windows.Forms.NumericUpDown
$script:numControlCheckUpdateHours.Minimum = 1
$script:numControlCheckUpdateHours.Maximum = 168
$script:numControlCheckUpdateHours.Value = 24
$script:numControlCheckUpdateHours.Width = 58
$script:numControlCheckUpdateHours.Margin = New-Object System.Windows.Forms.Padding(0,3,6,0)
$controlAutoRow.Controls.Add($script:numControlCheckUpdateHours) | Out-Null

$script:lblControlCheckUpdateHours = New-Object System.Windows.Forms.Label
$script:lblControlCheckUpdateHours.Text = 'hours'
$script:lblControlCheckUpdateHours.AutoSize = $true
$script:lblControlCheckUpdateHours.Margin = New-Object System.Windows.Forms.Padding(0,7,0,0)
$controlAutoRow.Controls.Add($script:lblControlCheckUpdateHours) | Out-Null

$controlLayout.Controls.Add($controlAutoRow,0,9) | Out-Null
$controlLayout.SetColumnSpan($controlAutoRow,2) | Out-Null

$controlExportRow = New-FlowButtonRow
$controlExportRow.Dock = 'None'
$controlExportRow.Anchor = 'Left'
$controlExportRow.WrapContents = $false
$controlExportRow.AutoSize = $true
$controlExportRow.MinimumSize = New-Object System.Drawing.Size(0, 40)

$script:btnExportDiagnostics = New-ActionButton -Text 'Export JSON' -Left 0 -Top 0 -Width 112
$script:btnExportDiagnostics.Margin = New-Object System.Windows.Forms.Padding(0,4,10,0)
$script:btnExportDiagnostics.Anchor = 'Left'
$script:btnExportDiagnostics.TabStop = $false
$controlExportRow.Controls.Add($script:btnExportDiagnostics) | Out-Null

$script:chkExportRedactSensitive = New-Object System.Windows.Forms.CheckBox
$script:chkExportRedactSensitive.Text = 'Redact sensitive info'
$script:chkExportRedactSensitive.Checked = $true
$script:chkExportRedactSensitive.AutoSize = $true
$script:chkExportRedactSensitive.Margin = New-Object System.Windows.Forms.Padding(0,11,0,0)
$script:chkExportRedactSensitive.Anchor = 'Left'
$script:chkExportRedactSensitive.TabStop = $false
$controlExportRow.Controls.Add($script:chkExportRedactSensitive) | Out-Null
$controlLayout.Controls.Add($controlExportRow,0,10) | Out-Null
$controlLayout.SetColumnSpan($controlExportRow,2) | Out-Null

$script:grpActivity = New-Object System.Windows.Forms.GroupBox

$script:grpActivity.Text = 'Activity'
$script:grpActivity.Dock = 'Fill'
$script:grpActivity.Margin = New-Object System.Windows.Forms.Padding(0)
$tabActivity.AutoScroll = $true
$tabActivity.Controls.Add($script:grpActivity)

$activityLayout = New-Object System.Windows.Forms.TableLayoutPanel
$activityLayout.Dock = 'Fill'
$activityLayout.ColumnCount = 1
$activityLayout.RowCount = 2
$activityLayout.Padding = New-Object System.Windows.Forms.Padding(12,12,12,12)
$null = $activityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$null = $activityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:grpActivity.Controls.Add($activityLayout)

$activityHeader = New-Object System.Windows.Forms.TableLayoutPanel
$activityHeader.Dock = 'Fill'
$activityHeader.ColumnCount = 4
$activityHeader.RowCount = 1
$activityHeader.Margin = New-Object System.Windows.Forms.Padding(0)
$activityHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$activityHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,220))) | Out-Null
$activityHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,96))) | Out-Null
$activityHeader.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,82))) | Out-Null
$activityLayout.Controls.Add($activityHeader,0,0)

$lblActivity = New-Object System.Windows.Forms.Label
$lblActivity.Text = 'Recent activity and command output.'
$lblActivity.Dock = 'Fill'
$lblActivity.TextAlign = 'MiddleLeft'
$activityHeader.Controls.Add($lblActivity,0,0)

$script:chkLogRefreshActivity = New-Object System.Windows.Forms.CheckBox
$script:chkLogRefreshActivity.Text = 'Show refresh events'
$script:chkLogRefreshActivity.Dock = 'Fill'
$script:chkLogRefreshActivity.CheckAlign = 'MiddleLeft'
$script:chkLogRefreshActivity.TextAlign = 'MiddleLeft'
$script:chkLogRefreshActivity.Margin = New-Object System.Windows.Forms.Padding(4,7,4,2)
$activityHeader.Controls.Add($script:chkLogRefreshActivity,1,0)

$script:btnActivityClearOutput = New-Object System.Windows.Forms.Button
$script:btnActivityClearOutput.Text = 'Clear Output'
$script:btnActivityClearOutput.Dock = 'Fill'
$script:btnActivityClearOutput.Margin = New-Object System.Windows.Forms.Padding(6,2,4,2)
$activityHeader.Controls.Add($script:btnActivityClearOutput,2,0)

$script:btnActivityClearLog = New-Object System.Windows.Forms.Button
$script:btnActivityClearLog.Text = 'Clear Log'
$script:btnActivityClearLog.Dock = 'Fill'
$script:btnActivityClearLog.Margin = New-Object System.Windows.Forms.Padding(4,2,0,2)
$activityHeader.Controls.Add($script:btnActivityClearLog,3,0)

$script:txtLog = New-Object System.Windows.Forms.RichTextBox
$script:txtLog.Dock = 'Fill'
$script:txtLog.Multiline = $true
$script:txtLog.ScrollBars = 'Vertical'
$script:txtLog.ReadOnly = $true
$script:txtLog.WordWrap = $false
$script:txtLog.DetectUrls = $false
$script:txtLog.MaxLength = [int]$script:ActivityTextMaxChars
$script:txtLog.BorderStyle = 'FixedSingle'
$script:txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$activityLayout.Controls.Add($script:txtLog,0,1)

$script:btnActivityClearOutput.Add_Click({ Clear-ActivityOutputView })
$script:btnActivityClearLog.Add_Click({ Clear-ActivityLogFile })
Set-AppToolTips
try {
    foreach ($box in @($script:txtPingDetails,$script:txtDnsResolveOutput,$script:txtPublicIpOutput,$script:txtMachineDetails,$script:txtMetricsSummary)) {
        if ($null -ne $box) { $box.MaxLength = [int]$script:UiTextMaxChars }
    }
} catch { }

$script:statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:statusStrip.SizingGrip = $false
$script:toolStatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:toolStatusLabel.Text = 'Ready.'
$script:statusStrip.Items.Add($script:toolStatusLabel) | Out-Null
$script:toolStatusSpacer = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:toolStatusSpacer.Spring = $true
$script:toolStatusSpacer.Text = ''
$script:statusStrip.Items.Add($script:toolStatusSpacer) | Out-Null
$script:toolAuthorLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:toolAuthorLabel.Text = ('Copyright ' + [char]0x00A9 + ' 2026 Luiz Bizzio')
$script:toolAuthorLabel.TextAlign = 'MiddleRight'
$script:statusStrip.Items.Add($script:toolAuthorLabel) | Out-Null
$script:statusStrip.Dock = 'Fill'
$rootLayout.Controls.Add($script:statusStrip,0,3)

$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:NotifyIcon.Icon = Get-AppNotifyIcon
$script:NotifyIcon.Text = 'Tailscale Control'
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$contextMenu.MinimumSize = New-Object System.Drawing.Size(340,0)
$contextMenu.ImageScalingSize = New-Object System.Drawing.Size(16,16)
$contextMenu.ShowImageMargin = $true
try { $script:TrayMenuRenderer = New-Object TailscaleControlTrayTextRendererV3; $contextMenu.Renderer = $script:TrayMenuRenderer } catch { }
$script:TrayMenuShow = $contextMenu.Items.Add('Tailscale Control')
$script:TrayMenuShow.Tag = [string]$script:AppVersion
$script:TrayMenuShow.ToolTipText = 'Open Tailscale Control'
Register-TrayVersionSuffix -Item $script:TrayMenuShow
Set-TrayTitleIcon
$script:TrayMenuMtu = $contextMenu.Items.Add('Tailscale MTU')
$script:TrayMenuMtu.Enabled = $false
$script:TrayMenuMtu.Tag = ''
$script:TrayMenuMtu.ToolTipText = 'Open Tailscale MTU'
Register-TrayVersionSuffix -Item $script:TrayMenuMtu
Set-TrayMtuIcon
$script:TrayMenuAdminPanel = $contextMenu.Items.Add('Tailscale Admin Panel')
Set-TrayMenuItemFixedSize -Item $script:TrayMenuAdminPanel -BaseWidth 300 -BaseHeight 28
Register-TrayExternalLinkGlyph -Item $script:TrayMenuAdminPanel
Set-TrayTailscaleIcon
$script:TrayMenuAdminSwitchSeparator = New-Object System.Windows.Forms.ToolStripSeparator
[void]$contextMenu.Items.Add($script:TrayMenuAdminSwitchSeparator)
$script:TrayMenuSelectAccount = New-Object System.Windows.Forms.ToolStripMenuItem
$script:TrayMenuSelectAccount.Text = 'Switch Tailnet'
Set-TrayMenuItemFixedSize -Item $script:TrayMenuSelectAccount -BaseWidth 300 -BaseHeight 28
Register-TraySubmenuArrowGlyph -Item $script:TrayMenuSelectAccount
[void]$contextMenu.Items.Add($script:TrayMenuSelectAccount)
[void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$script:TrayMenuNetworkDevices = New-Object System.Windows.Forms.ToolStripMenuItem
$script:TrayMenuNetworkDevices.Text = 'Network Devices'
Set-TrayMenuItemFixedSize -Item $script:TrayMenuNetworkDevices -BaseWidth 300 -BaseHeight 28
Register-TraySubmenuArrowGlyph -Item $script:TrayMenuNetworkDevices
try { $script:TrayMenuNetworkDevices.DropDownDirection = [System.Windows.Forms.ToolStripDropDownDirection]::Right } catch { }
[void]$script:TrayMenuNetworkDevices.DropDownItems.Add('Loading...')
Set-TrayDropDownRenderer -Item $script:TrayMenuNetworkDevices
[void]$contextMenu.Items.Add($script:TrayMenuNetworkDevices)
$script:TrayMenuInfoTopSeparator = New-Object System.Windows.Forms.ToolStripSeparator
[void]$contextMenu.Items.Add($script:TrayMenuInfoTopSeparator)
$script:TrayMenuInfoUser = $contextMenu.Items.Add('User: -')
$script:TrayMenuInfoAccountEmail = $contextMenu.Items.Add('Account Email: -')
$script:TrayMenuInfoTailnet = $contextMenu.Items.Add('Tailnet: -')
$script:TrayMenuInfoStatus = $contextMenu.Items.Add('Status: -')
$script:TrayMenuInfoDevice = $contextMenu.Items.Add('Device: -')
$script:TrayMenuInfoMagicDns = $contextMenu.Items.Add('MagicDNS: -')
$script:TrayMenuInfoIPv4 = $contextMenu.Items.Add('IPv4: -')
$script:TrayMenuInfoIPv6 = $contextMenu.Items.Add('IPv6: -')
$script:TrayMenuInfoDns = $contextMenu.Items.Add('DNS: -')
$script:TrayMenuInfoTailscaleVersion = $contextMenu.Items.Add('Tailscale Version: -')
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoUser -Label 'User'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoAccountEmail -Label 'Account Email'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoTailnet -Label 'Tailnet'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoStatus -Label 'Status'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoMagicDns -Label 'MagicDNS'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoDevice -Label 'Device'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoIPv4 -Label 'IPv4'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoIPv6 -Label 'IPv6'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoDns -Label 'DNS'
Register-TrayCopyMenuItem -Item $script:TrayMenuInfoTailscaleVersion -Label 'Tailscale Version'
Set-TrayCurrentDeviceInfoVisibility -Visible:([bool](Get-ObjectPropertyOrDefault $config 'show_current_device_info_in_tray' $false))
$contextMenu.add_MouseUp({
    param($trayMenuSender,$trayMouseEvent)
    try {
        if ($trayMouseEvent.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
        $hitItem = $trayMenuSender.GetItemAt($trayMouseEvent.Location)
        if ($null -eq $hitItem) { return }
        if ($hitItem -eq $script:TrayMenuInfoUser) { Copy-TrayMenuValue -Item $hitItem -Label 'User'; return }
        if ($hitItem -eq $script:TrayMenuInfoAccountEmail) { Copy-TrayMenuValue -Item $hitItem -Label 'Account Email'; return }
        if ($hitItem -eq $script:TrayMenuInfoTailnet) { Copy-TrayMenuValue -Item $hitItem -Label 'Tailnet'; return }
        if ($hitItem -eq $script:TrayMenuInfoStatus) { Copy-TrayMenuValue -Item $hitItem -Label 'Status'; return }
        if ($hitItem -eq $script:TrayMenuInfoMagicDns) { Copy-TrayMenuValue -Item $hitItem -Label 'MagicDNS'; return }
        if ($hitItem -eq $script:TrayMenuInfoDevice) { Copy-TrayMenuValue -Item $hitItem -Label 'Device'; return }
        if ($hitItem -eq $script:TrayMenuInfoIPv4) { Copy-TrayMenuValue -Item $hitItem -Label 'IPv4'; return }
        if ($hitItem -eq $script:TrayMenuInfoIPv6) { Copy-TrayMenuValue -Item $hitItem -Label 'IPv6'; return }
        if ($hitItem -eq $script:TrayMenuInfoDns) { Copy-TrayMenuValue -Item $hitItem -Label 'DNS'; return }
        if ($hitItem -eq $script:TrayMenuInfoTailscaleVersion) { Copy-TrayMenuValue -Item $hitItem -Label 'Tailscale Version'; return }
    }
    catch { Write-LogException -Context 'Tray copy from context menu' -ErrorRecord $_ }
})
$script:TrayMenuInfoBottomSeparator = New-Object System.Windows.Forms.ToolStripSeparator
[void]$contextMenu.Items.Add($script:TrayMenuInfoBottomSeparator)
$script:TrayMenuToggleConnect = $contextMenu.Items.Add('Connect [Unknown]')
$script:TrayMenuToggleExitNode = $contextMenu.Items.Add('Exit Node [Unknown]')
$script:TrayMenuToggleSubnets = $contextMenu.Items.Add('Subnets [Unknown]')
$script:TrayMenuToggleDns = $contextMenu.Items.Add('Accept DNS [Unknown]')
$script:TrayMenuToggleIncoming = $contextMenu.Items.Add('Incoming [Unknown]')
foreach ($trayItem in @($script:TrayMenuToggleConnect,$script:TrayMenuToggleExitNode,$script:TrayMenuToggleSubnets,$script:TrayMenuToggleDns,$script:TrayMenuToggleIncoming)) {
    if ($null -ne $trayItem) {
        $trayItem.CheckOnClick = $false
        $trayItem.ShowShortcutKeys = $false
        Set-TrayMenuItemFixedSize -Item $trayItem -BaseWidth 300 -BaseHeight 28
    }
}
[void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$script:TrayMenuChooseExitNode = New-Object System.Windows.Forms.ToolStripMenuItem
$script:TrayMenuChooseExitNode.Text = 'Preferred Exit Node'
Set-TrayMenuItemFixedSize -Item $script:TrayMenuChooseExitNode -BaseWidth 300 -BaseHeight 28
Register-TraySubmenuArrowGlyph -Item $script:TrayMenuChooseExitNode
[void]$contextMenu.Items.Add($script:TrayMenuChooseExitNode)
[void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$script:TrayMenuExit = $contextMenu.Items.Add('Exit')
Set-TrayMenuItemFixedSize -Item $script:TrayMenuExit -BaseWidth 300 -BaseHeight 28
$contextMenu.add_Opening({
    try {
        Update-TrayMenuFixedLayout -Menu $contextMenu
        if ($null -ne $script:TrayMenuAdminPanel) { $script:TrayMenuAdminPanel.Text = 'Tailscale Admin Panel' }
        Set-TrayTitleIcon
        Set-TrayTailscaleIcon
        Set-TrayMtuIcon
        Update-TraySelectAccountMenuState
        Update-TrayMenuState
    } catch { Write-LogException -Context 'Update tray menu state' -ErrorRecord $_ }
})
$script:TrayContextMenu = $contextMenu
$contextMenu.AutoClose = $true
$script:NotifyIcon.ContextMenuStrip = $contextMenu
$script:NotifyIcon.add_MouseUp({
    param($notifySender,$notifyMouseEvent)
    try {
        if ($notifyMouseEvent.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Show-MainForm
            return
        }
    }
    catch { Write-LogException -Context 'Notify icon mouse up' -ErrorRecord $_ }
})

$script:HotkeyWindow = New-Object TailscaleControlHotkeys
$script:HotkeyWindow.add_HotKeyPressed({
    param($controlSender,$e)
    foreach ($entry in $script:HotkeyIds.GetEnumerator()) {
        if ([int]$entry.Value -eq [int]$e.Id) {
            Invoke-HotkeyAction -Name $entry.Key
            break
        }
    }
})

$script:TrayMenuShow.add_Click({ Show-MainForm })
$script:TrayMenuAdminPanel.add_Click({
    Invoke-LoggedUrlOpen -Url 'https://login.tailscale.com/admin' -Title 'Tailscale Admin Panel' -CommandText 'Open Tailscale Admin Panel' -SuccessMessage 'Opened https://login.tailscale.com/admin' -FailureTitle 'Open Tailscale Admin Panel failed' -FailureOverlayTitle 'Open admin panel failed' -SkipActivity
})
$script:TrayMenuMtu.add_Click({
    try {
        $mtuInfo = Get-TailscaleMtuAppInfo
        if ($null -ne $mtuInfo -and [bool]$mtuInfo.Installed) { Open-TailscaleMtu }
    }
    catch { }
})
$script:TrayMenuToggleConnect.add_Click({ Invoke-TrayToggleAction 'Toggle Connect' { Switch-Connect } })
$script:TrayMenuToggleExitNode.add_Click({ Invoke-TrayToggleAction 'Toggle Exit Node' { Switch-ExitNode } })
$script:TrayMenuToggleSubnets.add_Click({ Invoke-TrayToggleAction 'Toggle Subnets' { Switch-Subnets } })
$script:TrayMenuToggleDns.add_Click({ Invoke-TrayToggleAction 'Toggle DNS' { Switch-Dns } })
$script:TrayMenuToggleIncoming.add_Click({ Invoke-TrayToggleAction 'Toggle Incoming' { Switch-Incoming } })
$script:TrayMenuExit.add_Click({
    $script:Exiting = $true
    try { $script:MainForm.Close() } catch {}
})
$script:NotifyIcon.add_DoubleClick({ Show-MainForm })

$script:btnRefresh.add_Click({ Start-StatusRefreshAsync -Button $script:btnRefresh })
$script:btnAccountSwitch.add_Click({ Invoke-Action 'Switch Account' { Switch-SelectedTailscaleAccount -Button $script:btnAccountSwitch } })
$script:btnAccountAdd.add_Click({ Invoke-Action 'Add Another Account' { Add-TailscaleAccount -Button $script:btnAccountAdd } })
$script:btnAccountLogout.add_Click({ Invoke-Action 'Logout Account' { Invoke-TailscaleLogoutCurrentAccount -Button $script:btnAccountLogout } })
$script:gridAccounts.add_CellDoubleClick({ Invoke-Action 'Switch Account' { Switch-SelectedTailscaleAccount } })
$leftTabs.add_SelectedIndexChanged({ try { if ($leftTabs.SelectedTab -eq $script:tabAccount) { Update-AccountView -Snapshot $script:Snapshot -RefreshAccounts $true } } catch { Write-LogException -Context 'Refresh account tab selection' -ErrorRecord $_ } })
$script:btnToggleConnect.add_Click({ Invoke-Action 'Toggle Connect' { Switch-Connect -Button $script:btnToggleConnect } })
$script:btnToggleExit.add_Click({ Invoke-Action 'Toggle Exit Node' { Switch-ExitNode -Button $script:btnToggleExit } })
$script:btnToggleDns.add_Click({ Invoke-Action 'Toggle DNS' { Switch-Dns -Button $script:btnToggleDns } })
$script:btnToggleSubnets.add_Click({ Invoke-Action 'Toggle Subnets' { Switch-Subnets -Button $script:btnToggleSubnets } })
$script:btnToggleIncoming.add_Click({ Invoke-Action 'Toggle Incoming' { Switch-Incoming -Button $script:btnToggleIncoming } })
$script:btnUpdate.add_Click({ Start-TailscaleControlMaintenanceTask -Operation 'Update' })
$script:btnUninstall.add_Click({
    try { Invoke-UninstallApp }
    catch { Show-Overlay -Title 'Uninstall failed' -Message $_.Exception.Message -ErrorStyle }
})
$script:btnInstallClientAutoUpdateTask.add_Click({ Invoke-TailscaleClientAutoUpdateTaskButton })
$script:btnCheckClientUpdate.add_Click({ Start-TailscaleClientMaintenanceTask -Operation 'Check' })
$script:btnAdminPanel.add_Click({
    Invoke-LoggedUrlOpen -Url 'https://login.tailscale.com/admin' -Title 'Tailscale Admin Panel' -CommandText 'Open Tailscale Admin Panel' -SuccessMessage 'Opened https://login.tailscale.com/admin' -FailureTitle 'Tailscale Admin Panel failed' -FailureOverlayTitle 'Open admin panel failed' -SkipActivity
})

$script:btnRunClientUpdate.add_Click({ Start-TailscaleClientMaintenanceTask -Operation 'Update' })
$script:btnInstallMtu.add_Click({
    try { Install-TailscaleMtu -Button $script:btnInstallMtu }
    catch { Show-Overlay -Title 'Tailscale MTU install failed' -Message $_.Exception.Message -ErrorStyle }
})
$script:btnOpenMtu.add_Click({
    try {
        Open-TailscaleMtu
        Write-ActivityCommandBlock -Title 'Open Tailscale MTU' -CommandText 'Open Tailscale MTU' -ExitCode 0 -Output 'Opened the installed Tailscale MTU app.'
    }
    catch {
        Write-ActivityFailureBlock -Title 'Open Tailscale MTU failed' -CommandText 'Open Tailscale MTU' -Message $_.Exception.Message
        Show-Overlay -Title 'Open Tailscale MTU failed' -Message $_.Exception.Message -ErrorStyle
    }
})
$script:btnCheckMtuRepo.add_Click({
    Invoke-LoggedUrlOpen -Url 'https://github.com/luizbizzio/tailscale-mtu' -Title 'Tailscale MTU Repo' -CommandText 'Open Tailscale MTU repository' -SuccessMessage 'Opened https://github.com/luizbizzio/tailscale-mtu' -FailureTitle 'Open Tailscale MTU repo failed' -FailureOverlayTitle 'Open repo failed' -SkipActivity
})
Register-TailscaleDiagnosticsButton -Button $script:btnDiagStatus -Kind 'Status' -Title 'Status' -Arguments @('status') -BusyText 'Loading...'
Register-TailscaleDiagnosticsButton -Button $script:btnDiagNetcheck -Kind 'Netcheck' -Title 'Netcheck' -Arguments @('netcheck') -BusyText 'Checking...'
Register-TailscaleDiagnosticsButton -Button $script:btnDiagDns -Kind 'DNS' -Title 'DNS' -Arguments @('dns','status','--all') -BusyText 'Checking...'
Register-TailscaleDiagnosticsButton -Button $script:btnDiagIPs -Kind 'IPs' -Title 'IPs' -Arguments @('ip') -BusyText 'Loading...'
Register-TailscaleDiagnosticsButton -Button $script:btnDiagMetrics -Kind 'Metrics' -Title 'Metrics' -Arguments @('metrics') -BusyText 'Loading...'
if ($null -ne $script:btnDiagClear) {
    $script:btnDiagClear.add_Click({
        try { $script:DiagnosticsContentMode = 'Selection'; Set-DiagnosticsOutput -Text '' -Mode 'Selection' }
        catch { Write-LogException -Context 'Run ping command block' -ErrorRecord $_ }
    })
}
Register-PingButton -Button $script:btnCmdPingAll -Kind 'All'
Register-PingButton -Button $script:btnCmdPingDns -Kind 'DNS'
Register-PingButton -Button $script:btnCmdPingIPv4 -Kind 'IPv4'
Register-PingButton -Button $script:btnCmdPingIPv6 -Kind 'IPv6'
$script:btnCmdWhois.add_Click({
    if ($script:IsDiagnosticsCommandTaskRunning -or $script:IsPingDiagnosticsTaskRunning) { return }
    try {
        Set-ControlFocusSafe -Control $script:btnCmdWhois
        $m = Get-SelectedMachine
        if ($null -eq $m) { return }
        $target = [string](Get-PropertyValue $m @('IPv4'))
        if ([string]::IsNullOrWhiteSpace($target)) { $target = [string](Get-PropertyValue $m @('IPv6')) }
        if ([string]::IsNullOrWhiteSpace($target)) { Set-DiagnosticsOutput -Text 'No IPv4 or IPv6 target is available for whois.'; return }
        Start-TailscaleDiagnosticsCommandAsync -Kind 'Whois' -Title 'Whois' -Arguments @('whois',$target) -Button $script:btnCmdWhois -BusyText 'Checking...'
    } catch {
        Set-DiagnosticsOutput -Text (Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Whois failed')
    }
})
$script:btnDetailMetrics.add_Click({
    try {
        $metricsText = Show-MetricsDiagnostics -ReturnText
        if ($null -eq $metricsText) { $metricsText = '' }
        if ($metricsText -isnot [string]) { $metricsText = [string]($metricsText | Out-String) }
        if ($null -ne $script:txtMetricsSummary) { $script:txtMetricsSummary.Text = Limit-UiText -Text $metricsText -MaxChars ([int]$script:MetricsTextMaxChars) -MaxLines ([int]$script:UiTextMaxLines) }
    }
    catch {
        $errText = Get-ExceptionDiagnosticText -ErrorRecord $_ -Prefix 'Metrics failed'
        try { Write-Log $errText } catch { }
        if ($null -ne $script:txtMetricsSummary) { $script:txtMetricsSummary.Text = Limit-UiText -Text $errText -MaxChars ([int]$script:MetricsTextMaxChars) -MaxLines ([int]$script:UiTextMaxLines) }
    }
})
$script:btnDetailClearMetrics.add_Click({
    try { if ($null -ne $script:txtMetricsSummary) { $script:txtMetricsSummary.Text = '' } }
    catch { Write-LogException -Context 'Clear metrics output' -ErrorRecord $_ }
})
$script:gridPing.add_SelectionChanged({
    try { Update-PingSelection -Machine $null } catch { Write-LogException -Context 'Clear ping selection after output clear' -ErrorRecord $_ }
})
$script:btnCheckControlRepo.add_Click({
    Invoke-LoggedUrlOpen -Url 'https://github.com/luizbizzio/tailscale-control' -Title 'Tailscale Control Repo' -CommandText 'Open Tailscale Control repository' -SuccessMessage 'Opened https://github.com/luizbizzio/tailscale-control' -FailureTitle 'Open Tailscale Control repo failed' -FailureOverlayTitle 'Open repo failed'
})
$script:btnCheckControlUpdate.add_Click({ Start-TailscaleControlMaintenanceTask -Operation 'Check' })
$script:btnOpenControlPath.add_Click({ Open-TailscaleControlPath })
$script:btnExportDiagnostics.add_Click({ Invoke-Action 'Export Diagnostic JSON' { $redactExport = $true; if ($null -ne $script:chkExportRedactSensitive) { $redactExport = [bool]$script:chkExportRedactSensitive.Checked }; Export-TailscaleDiagnosticJson -RedactSensitive $redactExport | Out-Null } })
$script:btnDnsResolveRun.add_Click({ Start-DnsResolveTestAsync })
$script:cmbDnsResolveResolver.add_SelectedIndexChanged({ Update-DnsResolveOtherState })
$script:txtDnsResolveOtherServer.add_TextChanged({ Update-DnsResolveServerPreview })
$script:btnPublicIpRun.add_Click({ Start-PublicIpTestAsync })
$machinesTabs.add_SelectedIndexChanged({
    try {
        if ($machinesTabs.SelectedTab -eq $tabMachinesDnsResolve) { Update-DnsResolveServerPreview }
    }
    catch { Write-LogException -Context 'DNS resolve tab selected preview refresh' -ErrorRecord $_ }
})
$machinesTabs.add_MouseDown({
    param($tabControlSender,$mouseEvent)
    try {
        for ($tabIndex = 0; $tabIndex -lt $machinesTabs.TabCount; $tabIndex++) {
            if ($machinesTabs.GetTabRect($tabIndex).Contains($mouseEvent.Location)) {
                if ($machinesTabs.TabPages[$tabIndex] -eq $tabMachinesDnsResolve) { Update-DnsResolveServerPreview }
                break
            }
        }
    }
    catch { Write-LogException -Context 'DNS resolve tab clicked preview refresh' -ErrorRecord $_ }
})
try { Update-DnsResolveDefaultDomain -Force; Update-DnsResolveServerPreview; Set-DnsResolveOutput -Text 'DNS RESOLVE TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Enter a domain, choose resolver, then click Resolve.'; Set-PublicIpOutput -Text 'PUBLIC IP TEST' + [Environment]::NewLine + [Environment]::NewLine + 'Choose Fast or Detailed, then click Test public IP to check the current outbound public IP.'; Update-DnsResolveOtherState } catch { }
$script:txtMachineFilter.add_TextChanged({ if ($null -ne $script:Snapshot) { Update-MachinesView -Snapshot $script:Snapshot } })
$script:gridMachines.add_SelectionChanged({
    try {
        if ($null -ne $script:gridMachines.CurrentRow -and $null -ne $script:gridMachines.CurrentRow.Tag) {
            $machine = $script:gridMachines.CurrentRow.Tag
            Update-MachineDetailsView -Machine $machine
            Update-PingSelection -Machine $machine
            Update-DiagnosticsSelectionSummary -Machine $machine
            Update-DnsResolveDefaultDomain
            Update-SelectedDeviceActionButtons
        }
    }
    catch { Write-LogException -Context 'Grid machine selection changed' -ErrorRecord $_ }
})
$script:gridMachines.add_CellDoubleClick({
    param($controlSender,$e)
    try {
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
            $value = [string]$script:gridMachines.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                [System.Windows.Forms.Clipboard]::SetText($value)
            }
        }
    } catch { Write-LogException -Context 'Copy machine cell by double click' -ErrorRecord $_ }
})
$script:gridMachines.add_CellMouseDown({
    param($controlSender,$e)
    try {
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
            $script:gridMachines.ClearSelection()
            $script:gridMachines.CurrentCell = $script:gridMachines.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
            $script:gridMachines.Rows[$e.RowIndex].Selected = $true
        }
    } catch { Write-LogException -Context 'Select machine row on right click' -ErrorRecord $_ }
})
foreach ($ctrl in @($script:chkStartup,$script:chkStartMinimized,$script:chkCloseToBackground,$script:chkShowTrayIcon,$script:chkCheckUpdateEvery,$script:chkControlCheckUpdateEvery,$script:chkTogglePopups,$script:chkToggleSounds,$script:chkLogRefreshActivity,$script:chkShowCurrentDeviceInfoInTray)) {
    if ($null -ne $ctrl) { $ctrl.add_CheckedChanged({ Save-UiSettings -Silent; try { Update-TailscaleClientTaskSetupUi } catch { } }) }
}
if ($null -ne $script:chkAllowLan) {
    $script:chkAllowLan.add_CheckedChanged({
        if ($script:IsLoadingConfig) { return }
        try { Set-AllowLanAccessPreferenceFromTray -Enabled ([bool]$script:chkAllowLan.Checked) }
        catch { Write-LogException -Context 'Apply Local Network Access preference' -ErrorRecord $_ }
    })
}
foreach ($ctrl in @($script:numCheckUpdateHours,$script:numControlCheckUpdateHours)) {
    if ($null -eq $ctrl) { continue }
    $ctrl.add_ValueChanged({ Save-UiSettings -Silent })
}
foreach ($ctrl in @($script:trkOverlay,$script:trkOverlayOpacity,$script:trkRefresh,$script:trkToggleSoundVolume)) {
    if ($null -eq $ctrl) { continue }
    $ctrl.add_ValueChanged({ Update-PreferenceSliderLabels })
    $ctrl.add_MouseUp({ Save-UiSettings -Silent })
    $ctrl.add_KeyUp({ Save-UiSettings -Silent })
}
if ($null -ne $script:cmbExitNode) { $script:cmbExitNode.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } }) }
foreach ($name in $script:HotkeyNames) {
    if ((Get-QuickAccountSwitchIndex -Name $name) -gt 0) { continue }
    if ($null -eq $script:HotkeyControls -or -not $script:HotkeyControls.ContainsKey($name)) { continue }
    $controls = $script:HotkeyControls[$name]
    if ($null -eq $controls) { continue }
    if ($null -ne $controls.Enabled) { $controls.Enabled.add_CheckedChanged({ Save-UiSettings -Silent }) }
    if ($null -ne $controls.Modifiers) { $controls.Modifiers.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } }) }
    if ($null -ne $controls.Key) { $controls.Key.add_SelectedIndexChanged({ if (-not $script:IsLoadingConfig) { Save-UiSettings -Silent } }) }
}

function Save-UiSettings {
    param([switch]$Silent)
    if ($script:IsLoadingConfig -or $script:IsSavingSettings) { return }
    $script:IsSavingSettings = $true
    try {
        $cfg = Get-Config
        $cfg.start_with_windows = [bool]$script:chkStartup.Checked
        $cfg.start_minimized = [bool]$script:chkStartMinimized.Checked
        $cfg.close_to_background = [bool]$script:chkCloseToBackground.Checked
        $cfg.show_tray_icon = [bool]$script:chkShowTrayIcon.Checked
        $cfg.allow_lan_on_exit = [bool]$script:chkAllowLan.Checked
        $cfg.overlay_seconds = [math]::Round(([double]$script:trkOverlay.Value / 10.0), 1)
        $cfg.overlay_opacity = [int]$script:trkOverlayOpacity.Value
        $cfg.refresh_seconds = [int]$script:trkRefresh.Value
        if ($null -ne $script:chkCheckUpdateEvery) { $cfg.check_update_every_enabled = ([bool]$script:chkCheckUpdateEvery.Checked -and (Test-TailscaleClientElevatedTasksReady)) }
        if ($null -ne $script:numCheckUpdateHours) { $cfg.check_update_every_hours = [int]$script:numCheckUpdateHours.Value }
        if ($null -ne $script:chkTogglePopups) { $cfg.show_toggle_popups = [bool]$script:chkTogglePopups.Checked }
        if ($null -ne $script:chkShowCurrentDeviceInfoInTray) { $cfg.show_current_device_info_in_tray = [bool]$script:chkShowCurrentDeviceInfoInTray.Checked }
        if ($null -ne $script:chkLogRefreshActivity) { $cfg.log_refresh_activity = [bool]$script:chkLogRefreshActivity.Checked }
        if ($null -ne $script:chkToggleSounds) { $cfg.play_toggle_sounds = [bool]$script:chkToggleSounds.Checked }
        if ($null -ne $script:trkToggleSoundVolume) { $cfg.toggle_sound_volume = [int]$script:trkToggleSoundVolume.Value }
        if ($null -ne $script:cmbExitNode.SelectedItem) {
            $cfg.preferred_exit_label = ConvertTo-DnsName ([string]$script:cmbExitNode.SelectedItem)
            $cfg.preferred_exit_node = $cfg.preferred_exit_label
        }
        else {
            $cfg.preferred_exit_label = ''
            $cfg.preferred_exit_node = ''
        }
        $quickCountForSave = Get-QuickAccountSwitchCountFromHotkeys -Hotkeys $cfg.hotkeys
        try { $quickCountForSave = [Math]::Max([int]$quickCountForSave, [int]@($script:QuickAccountSwitchAccounts).Count) } catch { }
        Ensure-ConfigHotkeyEntries -Config $cfg -Count $quickCountForSave
        foreach ($name in $script:HotkeyNames) {
            if (-not $script:HotkeyControls.ContainsKey($name)) { continue }
            $entry = $cfg.hotkeys.PSObject.Properties[$name].Value
            $controls = $script:HotkeyControls[$name]
            if ($null -eq $entry -or $null -eq $controls) { continue }
            $quickIndex = Get-QuickAccountSwitchIndex -Name $name
            if ($quickIndex -gt 0 -and -not [bool]$script:QuickAccountSwitchAvailable) {
                $entry.enabled = $false
            }
            else {
                $entry.enabled = [bool]$controls.Enabled.Checked
            }
            $entry.modifiers = [string]$controls.Modifiers.SelectedItem
            $entry.key = if ($null -ne $controls.Key.SelectedItem) { [string]$controls.Key.SelectedItem } else { [string]$controls.Key.Text }
            if ($quickIndex -gt 0 -and $null -ne $controls.Account) {
                $selectedIdentifier = ''
                $selectedDisplay = ''
                if ($null -ne $controls.Account.SelectedItem) {
                    $selectedIdentifier = [string](Get-ObjectPropertyOrDefault $controls.Account.SelectedItem 'SwitchIdentifier' '')
                    if ([string]::IsNullOrWhiteSpace($selectedIdentifier)) { $selectedIdentifier = [string](Get-ObjectPropertyOrDefault $controls.Account.SelectedItem 'Identifier' '') }
                    if ([string]::IsNullOrWhiteSpace($selectedIdentifier)) { $selectedIdentifier = [string](Get-ObjectPropertyOrDefault $controls.Account.SelectedItem 'Tailnet' '') }
                    $selectedDisplay = [string](Get-ObjectPropertyOrDefault $controls.Account.SelectedItem 'Display' '')
                }
                Add-Member -InputObject $entry -MemberType NoteProperty -Name 'account_identifier' -Value $selectedIdentifier -Force
                Add-Member -InputObject $entry -MemberType NoteProperty -Name 'account_display' -Value $selectedDisplay -Force
            }
        }
        Save-Config -Config $cfg
        Set-StartupSetting -Config $cfg
        $errors = @((Register-Hotkeys -Config $cfg) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        Update-NotifyVisibility
        Set-AppTheme
        Update-HotkeyToolTips
        if ($null -ne $script:UiTimer) { $script:UiTimer.Interval = [int]$cfg.refresh_seconds * 1000 }
        Update-Status
        if (-not $Silent) {
            if (@($errors).Count -gt 0) {
                Show-Overlay -Title 'Saved with warnings' -Message (($errors -join ' ') -replace '\s+',' ') -ErrorStyle
            }
        }
    }
    catch {
        Show-Overlay -Title 'Save failed' -Message $_.Exception.Message -ErrorStyle
    }
    finally {
        $script:IsSavingSettings = $false
        $script:AutoUpdateOverride = $null
    }
}

$script:MainForm.add_FormClosing({
    param($controlSender,$e)
    if (-not $script:Exiting) {
        $cfg = Get-Config
        if ([bool]$cfg.close_to_background) {
            $e.Cancel = $true
            Hide-MainFormToBackground
            return
        }
    }
    try { Clear-HotkeyExecutionLock } catch { Write-LogException -Context 'Release hotkey execution lock during shutdown' -ErrorRecord $_ }
    try { Unregister-Hotkeys } catch { Write-LogException -Context 'Unregister hotkeys during shutdown' -ErrorRecord $_ }
    try { if ($null -ne $script:UiTimer) { $script:UiTimer.Stop() } } catch { Write-LogException -Context 'Stop UI timer during shutdown' -ErrorRecord $_ }
    try { if ($null -ne $script:HotkeyPollTimer) { $script:HotkeyPollTimer.Stop(); $script:HotkeyPollTimer.Dispose(); $script:HotkeyPollTimer = $null } } catch { Write-LogException -Context 'Stop hotkey fallback timer during shutdown' -ErrorRecord $_ }
    try { if ($null -ne $script:NotifyIcon) { $script:NotifyIcon.Visible = $false; $script:NotifyIcon.Dispose() } } catch { Write-LogException -Context 'Dispose notify icon during uninstall' -ErrorRecord $_ }
    try { if ($null -ne $script:TrayContextMenu) { $script:TrayContextMenu.Dispose(); $script:TrayContextMenu = $null } } catch { Write-LogException -Context 'Dispose tray menu during shutdown' -ErrorRecord $_ }
    try { if ($null -ne $script:HotkeyWindow) { $script:HotkeyWindow.Dispose() } } catch { Write-LogException -Context 'Dispose hotkey window during shutdown' -ErrorRecord $_ }
    try { if ($null -ne $mutex) { $mutex.ReleaseMutex(); $mutex.Dispose() } } catch { Write-LogException -Context 'Release mutex during shutdown' -ErrorRecord $_ }
})

$script:UiTimer = New-Object System.Windows.Forms.Timer
$script:UiTimer.Interval = [int]$config.refresh_seconds * 1000
$script:UiTimer.add_Tick({
    if (
        -not $script:IsBusy -and
        -not $script:IsAsyncActionRunning -and
        -not $script:IsClientMaintenanceTaskRunning -and
        -not $script:IsControlMaintenanceTaskRunning
    ) {
        Invoke-ScheduledClientUpdate
        Invoke-ScheduledControlUpdate
        Update-Status
    }
})
Start-InstanceActivationMonitor
try { [void](Sync-QuickAccountSwitchFromLoggedAccounts -Config $config -Force) } catch { Write-LogException -Context 'Startup quick account switch sync' -ErrorRecord $_ }
Ensure-ConfigHotkeyEntries -Config $config -Count ([Math]::Max((Get-QuickAccountSwitchCountFromHotkeys -Hotkeys $config.hotkeys), [int]$script:QuickAccountSwitchNames.Count))
$script:ConfigCache = $config
$script:IsLoadingConfig = $true
foreach ($name in $script:HotkeyNames) {
    if (-not $script:HotkeyControls.ContainsKey($name)) { continue }
    if ($null -eq $config.hotkeys.PSObject.Properties[$name]) { continue }
    $entry = $config.hotkeys.PSObject.Properties[$name].Value
    $controls = $script:HotkeyControls[$name]
    $controls.Enabled.Checked = [bool]$entry.enabled
    $controls.Modifiers.SelectedItem = [string]$entry.modifiers
    if ($controls.Key.Items.Contains([string]$entry.key)) { $controls.Key.SelectedItem = [string]$entry.key } else { $controls.Key.Text = [string]$entry.key }
    if ($null -ne $controls.Capture) {
        $modsText = [string]$entry.modifiers
        $keyText = [string]$entry.key
        if ([string]::IsNullOrWhiteSpace($keyText)) { $controls.Capture.Text = '' }
        elseif ([string]::IsNullOrWhiteSpace($modsText) -or $modsText -eq 'None') { $controls.Capture.Text = $keyText }
        else { $controls.Capture.Text = $modsText + '+' + $keyText }
    }
}
Update-HotkeyToolTips
$script:chkStartup.Checked = [bool]$config.start_with_windows
$script:chkStartMinimized.Checked = [bool]$config.start_minimized
$script:chkCloseToBackground.Checked = [bool]$config.close_to_background
$script:chkShowTrayIcon.Checked = [bool]$config.show_tray_icon
$script:chkAllowLan.Checked = [bool]$config.allow_lan_on_exit
$script:trkOverlay.Value = [int]([math]::Round([double]$config.overlay_seconds * 10.0))
$script:trkOverlayOpacity.Value = [int]$config.overlay_opacity
$script:trkRefresh.Value = [int]$config.refresh_seconds
if ($null -ne $script:chkCheckUpdateEvery) { $script:chkCheckUpdateEvery.Checked = [bool]$config.check_update_every_enabled }
if ($null -ne $script:numCheckUpdateHours) { $script:numCheckUpdateHours.Value = [decimal]$config.check_update_every_hours }
Update-TailscaleClientTaskSetupUi
Start-TailscaleClientTaskReadinessRefresh
if ($null -ne $script:chkControlCheckUpdateEvery) { $script:chkControlCheckUpdateEvery.Checked = [bool]$config.control_check_update_every_enabled }
if ($null -ne $script:numControlCheckUpdateHours) { $script:numControlCheckUpdateHours.Value = [decimal]$config.control_check_update_every_hours }
if ($null -ne $script:chkTogglePopups) { $script:chkTogglePopups.Checked = [bool]$config.show_toggle_popups }
if ($null -ne $script:chkShowCurrentDeviceInfoInTray) { $script:chkShowCurrentDeviceInfoInTray.Checked = [bool](Get-ObjectPropertyOrDefault $config 'show_current_device_info_in_tray' $false) }
if ($null -ne $script:chkLogRefreshActivity) { $script:chkLogRefreshActivity.Checked = [bool](Get-ObjectPropertyOrDefault $config 'log_refresh_activity' $false) }
if ($null -ne $script:chkToggleSounds) { $script:chkToggleSounds.Checked = [bool](Get-ObjectPropertyOrDefault $config 'play_toggle_sounds' $false) }
if ($null -ne $script:trkToggleSoundVolume) { $script:trkToggleSoundVolume.Value = Convert-ToSafeInt (Get-ObjectPropertyOrDefault $config 'toggle_sound_volume' 20) 100 }
Update-PreferenceSliderLabels
$script:IsLoadingConfig = $false
$script:IsSavingSettings = $false
$script:AutoUpdateOverride = $null

if ($Background) {
    try { $script:MainForm.Opacity = 0.0 } catch { }
    try { $script:MainForm.ShowInTaskbar = $false } catch { }
}
Set-AppTheme
Set-StartupSetting -Config $config
Update-NotifyVisibility
$startupHotkeyErrors = @((Register-Hotkeys -Config $config) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
foreach ($startupHotkeyError in $startupHotkeyErrors) { try { Write-Log ('Hotkey warning: ' + [string]$startupHotkeyError) } catch { } }
Start-ShowSettingsHotkeyFallback

$script:MainForm.add_Load({
    try { Set-MainFormAppIcon } catch { }
    if ($script:StartupHidePending) {
        try { $script:MainForm.Opacity = 0.0 } catch { }
        try { $script:MainForm.ShowInTaskbar = $false } catch { }
        try { $script:MainForm.WindowState = 'Minimized' } catch { }
        try { $script:MainForm.Location = New-Object System.Drawing.Point(-32000,-32000) } catch { }
        try { Start-HideMainFormToBackground } catch { }
    }
    else {
        try { Set-MainFormCenteredOnPrimaryScreen } catch { }
    }
})

$script:MainForm.add_Shown({
    try { Set-MainFormAppIcon } catch { }
    try {
        if (-not $script:StartupHidePending) {
            Set-BodySplitPreferredLayout
        }
        Update-RightLayoutSizing
    }
    catch {
        Write-LogException -Context 'Split sizing adjustment' -ErrorRecord $_
    }
    try { Set-MachineColumnLayout -UseConfig } catch { Write-Log ('Machine column layout apply failed: ' + $_.Exception.Message) }
    try { Update-DnsResolveDefaultDomain -Force; Update-DnsResolveOtherState } catch { }
    Update-Status
    $script:UiTimer.Start()
    if (-not $Background) {
        $script:StartupHidePending = $false
        try { Set-MainFormCenteredOnPrimaryScreen } catch { }
        $script:MainFormPresentedOnce = $true
        try { $script:MainForm.Opacity = 1.0 } catch { }
        try { $script:MainForm.ShowInTaskbar = $true } catch { }
    }
    if ($script:StartupHidePending) {
        Write-Log 'Window hidden because background launch mode is enabled.'
        Start-HideMainFormToBackground
    }
})

if (-not $Background) { try { Set-MainFormCenteredOnPrimaryScreen } catch { } }
Write-Log 'Entering Windows message loop.'
[System.Windows.Forms.Application]::Run($script:MainForm)
}
catch {
    $msg = $_.Exception.Message
    $pos = ''
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { $pos = $_.InvocationInfo.PositionMessage }
    Write-Log ('UI initialization failed: ' + $msg)
    if (-not [string]::IsNullOrWhiteSpace($pos)) { Write-Log ('UI initialization position: ' + ($pos -replace "`r?`n", ' | ')) }
    Write-Host ('UI initialization failed: ' + $msg) -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($pos)) { Write-Host $pos -ForegroundColor Yellow }
    throw
}
