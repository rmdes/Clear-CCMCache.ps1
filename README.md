# Clear-CCMCache

PowerShell script that removes old, unused content from the SCCM/CCM client cache (`CCMCache`) and reconciles orphaned entries between disk and the CCM client store.

## Requirements

- Windows PowerShell **5.1** or PowerShell **7+**
- Administrative privileges
- Microsoft Endpoint Configuration Manager (SCCM/CCM) client installed

## Usage

```powershell
.\Clear-CCMCache.ps1 [-Days <int>] [-WhatIf] [-Confirm] [-Verbose]
```

| Parameter  | Description |
| ---------- | --- |
| `-Days`    | Days an item must be unreferenced before removal. Default `30`, range `1`–`3650`. |
| `-WhatIf`  | Show what would be removed without making any changes. |
| `-Confirm` | Prompt before each removal. |
| `-Verbose` | Per-item progress on the verbose stream. |

### Examples

```powershell
# Default: clean items unreferenced for more than 30 days
.\Clear-CCMCache.ps1

# Preview a 14-day cleanup, no changes made
.\Clear-CCMCache.ps1 -Days 14 -WhatIf

# Run a 14-day cleanup with verbose output
.\Clear-CCMCache.ps1 -Days 14 -Verbose

# Full help
Get-Help .\Clear-CCMCache.ps1 -Detailed
```

## What it does

1. Resolves the cache path from `ROOT\ccm\SoftMgmtAgent\CacheConfig`.
2. Removes cache entries (folder + CIM record) whose `LastReferenced` is older than `-Days`.
3. Reconciles orphans:
   - Folders on disk with no matching CIM record → deleted.
   - CIM records pointing at folders that no longer exist → deleted.

Every destructive step honors `-WhatIf` and `-Confirm`, and is bounded to paths under the resolved cache root as a safety guard.

## Notes

- The default cache path is `C:\Windows\ccmcache`, but the script always reads the actual location from CIM.
- Uses `Get-CimInstance` / `Remove-CimInstance` (not the deprecated `Get-WmiObject`), so the script runs unchanged on PowerShell 7.
- Provided as-is. Not supported by Microsoft.

## Contributing

Bug reports and pull requests welcome.

## Acknowledgements

Special thanks to [@theruck242](https://github.com/theruck242) for identifying a critical issue and providing a robust solution to enhance the script's safety and reliability.
