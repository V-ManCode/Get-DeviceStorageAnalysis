# Get-DeviceStorageAnalysis

See at a glance how many Intune-managed devices are running out of disk space, across the whole tenant or only in the Entra ID groups you choose. Devices are grouped into free-space bands, each band is broken down by Windows Autopilot group tag, and any band can be saved to CSV. Read-only against your tenant.

## Output

Console summary:

```
0-1%      : 12 devices
    Sales                                    7
    No tag                                   5
1-5%      : 48 devices
...
```

Bands: **0-1%, 1-5%, 5-10%, 10-15%, 15-20%, Over 20%**

Each band CSV (`Band_<band>_<yyyyMMdd>.csv`) has: DeviceName, SerialNumber, GroupTag, UPN, LastSync, FreeGB, TotalGB, FreePct, plus ScopeGroups (which selected groups the device is in) when you scan by group.

## Scope

Run it with no scope parameters and it asks:

```
Which devices do you want to analyse?
  [1] All devices in the tenant
  [2] Devices in one or more Entra ID groups
```

Option 2 takes one or more group IDs (or exact names), separated by commas. Nested groups are included, and a device in several selected groups is counted once.

## Parameters

| Parameter | Default | What it does |
|---|---|---|
| `-GroupId` | - | One or more group IDs (or names) to scan. Skips the scope prompt |
| `-AllDevices` | off | Scan all devices in the tenant. Skips the scope prompt |
| `-OperatingSystem` | `Windows` | Platform to report on |
| `-ExcludeGroupName` | - | Leave out every device in this Entra ID group (nested groups included). Stops if the group can't be fully read |
| `-MaxSyncAgeDays` | `0` (all) | Only devices that checked in within this many days |
| `-OutputFolder` | script folder | Where CSVs are saved |
| `-ExportAll` | off | Save every band without Y/N prompts |
| `-TenantId` | - | Tenant to sign in to |

## Requirements

- PowerShell 5.1 or 7+
- `Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.DeviceManagement, Microsoft.Graph.DeviceManagement.Enrollment -Scope CurrentUser`
- `Install-Module Microsoft.Graph.Groups -Scope CurrentUser` (only when scanning or excluding groups)
- Graph permissions (delegated or application):
  - `DeviceManagementManagedDevices.Read.All`
  - `DeviceManagementServiceConfig.Read.All` (Autopilot group tags)
  - `GroupMember.Read.All`, `Device.Read.All` (only when scanning or excluding groups)

## Notes

- Free/total storage comes from Intune hardware inventory, so it is only as fresh as the device's last inventory. Use `-MaxSyncAgeDays` to drop devices that haven't checked in recently.
- Devices that haven't reported storage are skipped and counted in the console.

## Usage

```powershell
# Interactive: choose the scope, see the summary, choose which bands to save
.\Get-DeviceStorageAnalysis.ps1

# Two groups, save every band
.\Get-DeviceStorageAnalysis.ps1 -GroupId '<group-id-1>','<group-id-2>' -ExportAll

# Whole tenant, exclude a group, only active devices
.\Get-DeviceStorageAnalysis.ps1 -AllDevices -ExcludeGroupName 'Devices - Do Not Touch' -MaxSyncAgeDays 30
```

## Licence

MIT. Provided as is, without warranty. Test before using in production.

Author: Vlad Grigoroiu
