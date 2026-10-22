<#
.SYNOPSIS
    Groups Intune-managed devices into free disk space bands and exports a CSV per band.

.DESCRIPTION
    Reads free and total storage for Intune-managed devices of the chosen operating system
    (Windows by default) and sorts them into bands of free space:

      0-1% | 1-5% | 5-10% | 10-15% | 15-20% | Over 20%

    Scope: when run without -GroupId or -AllDevices, the script asks whether to analyse all devices
    in the tenant or only the devices in one or more Entra ID groups (by group ID or name,
    nested groups included).

    For each band the console shows the device count, broken down by Windows Autopilot group tag.
    You are then asked, band by band, whether to save that band's device list to CSV
    (or use -ExportAll to save every band without prompts).

    Optional: -ExcludeGroupName removes every device in an Entra ID group (nested groups included)
    from all results. If that group can't be fully read, the script stops rather than include them.

    Storage values come from the device's Intune hardware inventory, so they are only as fresh as
    its last inventory. Devices that have not reported storage are skipped.

    The script is READ-ONLY against your tenant. It never changes anything in Intune or Entra ID.

.PARAMETER GroupId
    One or more Entra ID group IDs (or display names) to analyse. Skips the scope prompt.

.PARAMETER AllDevices
    Analyse all devices in the tenant. Skips the scope prompt.

.PARAMETER OperatingSystem
    Operating system to report on, as Intune names it. Default: Windows.

.PARAMETER ExcludeGroupName
    Optional: display name of an Entra ID group whose devices are left out of all results.

.PARAMETER MaxSyncAgeDays
    Optional: only include devices that checked in with Intune within this many days. Default: 0 (all).

.PARAMETER OutputFolder
    Where band CSVs are saved. Default: the script's folder.

.PARAMETER ExportAll
    Save every non-empty band to CSV without prompting.

.PARAMETER TenantId
    Optional: the tenant to sign in to.

.EXAMPLE
    .\Get-DeviceStorageAnalysis.ps1
    Prompts for the scope (all devices, or one or more groups).

.EXAMPLE
    .\Get-DeviceStorageAnalysis.ps1 -GroupId '<group-id-1>','<group-id-2>' -ExportAll

.EXAMPLE
    .\Get-DeviceStorageAnalysis.ps1 -AllDevices -ExcludeGroupName 'Devices - Do Not Touch' -MaxSyncAgeDays 30

.NOTES
    Author   : Vlad Grigoroiu
    Version  : 1.3.0
    Date     : 2026-10-22
    License  : MIT
    Requires : PowerShell 5.1 or 7+, Microsoft Graph PowerShell SDK modules:
               Microsoft.Graph.Authentication, Microsoft.Graph.DeviceManagement,
               Microsoft.Graph.DeviceManagement.Enrollment,
               Microsoft.Graph.Groups (only when groups are used)
    Scopes   : DeviceManagementManagedDevices.Read.All, DeviceManagementServiceConfig.Read.All
               + GroupMember.Read.All, Device.Read.All when groups are used
    Disclaimer: Provided as is, without warranty of any kind. Test before use in production.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$GroupId,
    [switch]$AllDevices,
    [string]$OperatingSystem = 'Windows',
    [string]$ExcludeGroupName,
    [ValidateRange(0, 3650)][int]$MaxSyncAgeDays = 0,
    [string]$OutputFolder,
    [switch]$ExportAll,
    [string]$TenantId
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.3.0'
$EmptyGuid   = '00000000-0000-0000-0000-000000000000'
$GuidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

#region Helpers
function Read-GroupRefs([switch]$AllowEmpty) {
    $prompt = if ($AllowEmpty) { 'Re-enter group ID(s), or press Enter to continue with the groups found' }
              else { 'Enter group ID(s) or name(s), separated by commas' }
    do {
        $raw = Read-Host $prompt
        $refs = @("$raw" -split ',' | ForEach-Object { $_.Trim().Trim('"', "'") } | Where-Object { $_ })
    } until ($refs.Count -or $AllowEmpty)
    , $refs
}

function Resolve-Group([string]$Ref) {
    # Accepts a group object ID or an exact display name. Returns $null if not found or ambiguous.
    if ($Ref -match $GuidPattern) {
        try { return Get-MgGroup -GroupId $Ref -Property Id, DisplayName -ErrorAction Stop } catch { return $null }
    }
    $found = @(Get-MgGroup -Filter "displayName eq '$($Ref.Replace("'", "''"))'" -Property Id, DisplayName)
    if ($found.Count -gt 1) { Write-Warning "$($found.Count) groups are named '$Ref' - use the group ID instead." }
    if ($found.Count -eq 1) { return $found[0] }
    $null
}

function Get-GroupDeviceIds([string]$Id) {
    # Device members incl. nested groups. Device type is filtered client-side: a server-side OData cast
    # needs ConsistencyLevel=eventual, which uses an index that can lag behind recent changes.
    $members = @(Get-MgGroupTransitiveMember -GroupId $Id -All |
        Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.device' })
    $ids = @($members | ForEach-Object { $_.AdditionalProperties['deviceId'] } | Where-Object { $_ })
    [pscustomobject]@{ MemberCount = $members.Count; DeviceIds = $ids }
}
#endregion

#region Scope choice (before sign-in, so only the permissions needed are requested)
if ($AllDevices -and $GroupId) { throw 'Use either -AllDevices or -GroupId, not both.' }
$interactiveScope = -not $AllDevices -and -not $GroupId
if ($interactiveScope) {
    Write-Host "`nWhich devices do you want to analyse?" -ForegroundColor Cyan
    Write-Host '  [1] All devices in the tenant'
    Write-Host '  [2] Devices in one or more Entra ID groups'
    do { $choice = "$(Read-Host 'Choose 1 or 2')".Trim() } until ($choice -in '1', '2')
    if ($choice -eq '2') { $GroupId = Read-GroupRefs }
}
$useGroups = [bool]$GroupId
#endregion

#region Modules and sign-in
$modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DeviceManagement', 'Microsoft.Graph.DeviceManagement.Enrollment')
if ($useGroups -or $ExcludeGroupName) { $modules += 'Microsoft.Graph.Groups' }
$missingModules = @($modules | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
if ($missingModules.Count) {
    throw "Missing module(s): $($missingModules -join ', '). Install with: Install-Module $($missingModules -join ',') -Scope CurrentUser"
}
$modules | ForEach-Object { Import-Module $_ }

$scopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementServiceConfig.Read.All')
if ($useGroups -or $ExcludeGroupName) { $scopes += 'GroupMember.Read.All', 'Device.Read.All' }

$ctx = Get-MgContext
$needConnect = -not $ctx
if ($ctx -and $ctx.AuthType -eq 'Delegated') {
    $missing = @($scopes | Where-Object { $ctx.Scopes -notcontains $_ -and $ctx.Scopes -notcontains ($_ -replace '\.Read\.', '.ReadWrite.') })
    if ($missing.Count) { $needConnect = $true }
}
if ($TenantId -and $ctx -and $ctx.TenantId -ne $TenantId) { $needConnect = $true }
if ($needConnect) {
    $cp = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $cp.TenantId = $TenantId }
    Connect-MgGraph @cp
}

if (-not $OutputFolder) { $OutputFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path } }
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
Write-Host "Get-DeviceStorageAnalysis v$ScriptVersion - tenant $((Get-MgContext).TenantId)" -ForegroundColor Cyan
#endregion

#region Scope groups
$scopeIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$deviceGroups = @{}   # Entra deviceId -> names of the selected groups it belongs to
$scopeGroups = @()
if ($useGroups) {
    $pending = @($GroupId)
    while ($true) {
        $notFound = @()
        foreach ($ref in $pending) {
            $g = Resolve-Group $ref
            if (-not $g) { $notFound += $ref }
            elseif (@($scopeGroups | Where-Object { $_.Id -eq $g.Id }).Count -eq 0) { $scopeGroups += $g }
        }
        if (-not $notFound.Count) { break }
        Write-Warning "Group(s) not found: $($notFound -join ', ')"
        if (-not $interactiveScope) { throw "Group(s) not found: $($notFound -join ', ')" }
        $pending = Read-GroupRefs -AllowEmpty:($scopeGroups.Count -gt 0)
        if (-not $pending.Count) { break }
    }

    foreach ($g in $scopeGroups) {
        $r = Get-GroupDeviceIds $g.Id
        Write-Host ("Group '{0}': {1} devices" -f $g.DisplayName, $r.DeviceIds.Count) -ForegroundColor Cyan
        if ($r.MemberCount -ne $r.DeviceIds.Count) {
            Write-Warning "$($r.MemberCount - $r.DeviceIds.Count) devices in '$($g.DisplayName)' returned no deviceId (missing Device.Read.All consent?) and can't be included."
        }
        foreach ($id in $r.DeviceIds) {
            [void]$scopeIds.Add([string]$id)
            if (-not $deviceGroups.ContainsKey($id)) { $deviceGroups[$id] = New-Object System.Collections.Generic.List[string] }
            if (-not $deviceGroups[$id].Contains($g.DisplayName)) { $deviceGroups[$id].Add($g.DisplayName) }
        }
    }
    if ($scopeIds.Count -eq 0) { throw 'The selected group(s) contain no devices. Nothing to analyse.' }
}
#endregion

#region Optional exclusion group (fail-safe: stops if the group can't be fully read)
$excludeIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($ExcludeGroupName) {
    $excludeGroup = @(Get-MgGroup -Filter "displayName eq '$($ExcludeGroupName.Replace("'", "''"))'")
    if ($excludeGroup.Count -ne 1) {
        throw "Expected exactly 1 group named '$ExcludeGroupName', found $($excludeGroup.Count). Stopping so its devices are not included."
    }
    $r = Get-GroupDeviceIds $excludeGroup[0].Id
    if ($r.MemberCount -eq 0) {
        throw "Group '$ExcludeGroupName' has no device members. Stopping so no devices are wrongly included."
    }
    if ($r.DeviceIds.Count -ne $r.MemberCount) {
        throw "$($r.MemberCount - $r.DeviceIds.Count) group devices returned no deviceId (missing Device.Read.All consent?). Stopping so they are not included."
    }
    $r.DeviceIds | ForEach-Object { [void]$excludeIds.Add([string]$_) }
    Write-Host "Devices in exclusion group '$ExcludeGroupName': $($excludeIds.Count)" -ForegroundColor Magenta
}
#endregion

#region Data
# Autopilot identities: serial number -> group tag (trimmed, so stray spaces don't split the breakdown)
$apLookup = @{}
Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All |
    ForEach-Object { if ($_.SerialNumber) { $apLookup["$($_.SerialNumber)".Trim()] = "$($_.GroupTag)".Trim() } }

$safeOs = $OperatingSystem.Replace("'", "''")
$intuneDevices = @(Get-MgDeviceManagementManagedDevice -All `
    -Filter "operatingSystem eq '$safeOs'" `
    -Property DeviceName, SerialNumber, UserPrincipalName, LastSyncDateTime, AzureADDeviceId,
              TotalStorageSpaceInBytes, FreeStorageSpaceInBytes)
Write-Host "Intune $OperatingSystem devices: $($intuneDevices.Count)" -ForegroundColor Cyan

if ($useGroups) {
    $intuneDevices = @($intuneDevices | Where-Object { $scopeIds.Contains([string]$_.AzureADDeviceId) })
    Write-Host "  In the selected group(s): $($intuneDevices.Count)" -ForegroundColor Cyan
}

$noEntraId = @($intuneDevices | Where-Object { -not $_.AzureADDeviceId -or $_.AzureADDeviceId -eq $EmptyGuid })
if ($ExcludeGroupName) {
    $excluded = @($intuneDevices | Where-Object { $excludeIds.Contains([string]$_.AzureADDeviceId) })
    Write-Host "  Excluded by group: $($excluded.Count)" -ForegroundColor Magenta
    if ($noEntraId.Count) {
        $names = if ($noEntraId.Count -le 20) { ': ' + (($noEntraId.DeviceName) -join ', ') } else { '' }
        Write-Warning "$($noEntraId.Count) devices have no Entra device ID and cannot be checked against the exclusion group$names"
    }
}

$noStorage = @($intuneDevices | Where-Object { -not ($_.TotalStorageSpaceInBytes -gt 0) }).Count
if ($noStorage) { Write-Host "  Skipped (no storage reported): $noStorage" -ForegroundColor DarkGray }

$cutoff = if ($MaxSyncAgeDays -gt 0) { (Get-Date).ToUniversalTime().AddDays(-$MaxSyncAgeDays) } else { $null }
$devices = @($intuneDevices | Where-Object {
    $_.TotalStorageSpaceInBytes -gt 0 -and
    -not $excludeIds.Contains([string]$_.AzureADDeviceId) -and
    (-not $cutoff -or ($_.LastSyncDateTime -and ([datetime]$_.LastSyncDateTime).ToUniversalTime() -ge $cutoff)) })
if ($cutoff) { Write-Host "  Only devices synced in the last $MaxSyncAgeDays days" -ForegroundColor DarkGray }

$pct = foreach ($d in $devices) {
    $serial = "$($d.SerialNumber)".Trim()
    $tag = if ($serial -and $apLookup.ContainsKey($serial)) { $apLookup[$serial] } else { '' }
    $row = [ordered]@{
        DeviceName   = $d.DeviceName
        SerialNumber = $d.SerialNumber
        GroupTag     = if ($tag) { $tag } else { 'No tag' }
        UPN          = $d.UserPrincipalName
        LastSync     = $d.LastSyncDateTime
        FreeGB       = [math]::Round($d.FreeStorageSpaceInBytes / 1GB, 1)
        TotalGB      = [math]::Round($d.TotalStorageSpaceInBytes / 1GB, 1)
        FreePct      = [math]::Round(($d.FreeStorageSpaceInBytes / $d.TotalStorageSpaceInBytes) * 100, 1)
    }
    if ($useGroups) { $row.ScopeGroups = ($deviceGroups[[string]$d.AzureADDeviceId]) -join '; ' }
    [pscustomobject]$row
}
$pct = @($pct)
#endregion

#region Bands
$bands = @(
    @{ Label = '0-1%'    ; File = 'Band_0-1pct'   ; Filter = { $_.FreePct -le 1 } }
    @{ Label = '1-5%'    ; File = 'Band_1-5pct'   ; Filter = { $_.FreePct -gt 1  -and $_.FreePct -le 5 } }
    @{ Label = '5-10%'   ; File = 'Band_5-10pct'  ; Filter = { $_.FreePct -gt 5  -and $_.FreePct -le 10 } }
    @{ Label = '10-15%'  ; File = 'Band_10-15pct' ; Filter = { $_.FreePct -gt 10 -and $_.FreePct -le 15 } }
    @{ Label = '15-20%'  ; File = 'Band_15-20pct' ; Filter = { $_.FreePct -gt 15 -and $_.FreePct -le 20 } }
    @{ Label = 'Over 20%'; File = 'Band_Over20pct'; Filter = { $_.FreePct -gt 20 } }
)

$scopeText = if ($useGroups) { "group(s): $(($scopeGroups.DisplayName) -join ', ')" } else { 'all devices' }
Write-Host "`nScope: $scopeText" -ForegroundColor Cyan
Write-Host "Devices analysed: $($pct.Count)`n" -ForegroundColor Cyan
$results = foreach ($b in $bands) {
    $b.Devices = @($pct | Where-Object $b.Filter | Sort-Object FreePct)
    Write-Host ("{0,-9} : {1} devices" -f $b.Label, $b.Devices.Count) -ForegroundColor Yellow
    $b.Devices | Group-Object GroupTag | Sort-Object Count -Descending |
        ForEach-Object { Write-Host ("    {0,-40} {1}" -f $_.Name, $_.Count) }
    $b
}

Write-Host ''
foreach ($b in $results) {
    if ($b.Devices.Count -eq 0) { continue }
    $save = $ExportAll
    if (-not $save) { $save = (Read-Host "Save list for $($b.Label) ($($b.Devices.Count) devices)? [Y/N]") -match '^y' }
    if ($save) {
        $path = Join-Path $OutputFolder "$($b.File)_$(Get-Date -Format yyyyMMdd).csv"
        $b.Devices | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        Write-Host "  Saved: $path" -ForegroundColor Green
    }
}
#endregion
