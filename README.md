# Clear-CCMCache

PowerShell script that removes old, unused content from the SCCM/CCM client cache (`CCMCache`) using the supported `UIResource.UIResourceMgr` COM interface, and reconciles legacy orphans between disk and the CIM index.

## Requirements

- Windows PowerShell **5.1** or PowerShell **7+**
- Administrative privileges
- Microsoft Endpoint Configuration Manager (SCCM/CCM) client installed and **CcmExec service running**

## Usage

```powershell
.\Clear-CCMCache.ps1 [-Days <int>] [-IncludePersisted] [-IncludeInUse] [-WhatIf] [-Confirm] [-Verbose]
```

| Parameter           | Description |
| ------------------- | --- |
| `-Days`             | Days an item must be unreferenced before removal. Default `30`, range `1`–`3650`. |
| `-IncludePersisted` | Also remove entries flagged `PersistInCache=$true`. Off by default — deleting persisted content typically triggers redownload on the next policy evaluation. |
| `-IncludeInUse`     | Also attempt to remove entries with `ReferenceCount > 0`. Off by default — in-use content is being read by a running install or task sequence. The COM interface may still refuse the delete. |
| `-WhatIf`           | Show what would be removed without making any changes. |
| `-Confirm`          | Prompt before each removal. |
| `-Verbose`          | Per-item progress and a summary line on the verbose stream. |

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

1. Preflights: confirms `CcmExec` is running and resolves the cache path from `ROOT\ccm\SoftMgmtAgent\CacheConfig`.
2. **Main pass** — for every cache element, joins COM (`UIResource.UIResourceMgr.GetCacheInfo()`) with CIM (`CacheInfoEx`):
   - Skips entries unreferenced for ≤ `-Days`.
   - Skips persisted entries unless `-IncludePersisted`.
   - Skips in-use entries unless `-IncludeInUse`.
   - Deletes survivors via `DeleteCacheElement()` — the supported, locking-aware path.
3. **Orphan reconciliation** (best effort, for legacy state from older scripts or interrupted ops):
   - Folders on disk with no CIM record → deleted.
   - CIM records pointing at missing folders → deleted (COM first, manual fallback).

Every destructive step honors `-WhatIf` and `-Confirm`, and disk deletions are bounded to paths under the resolved cache root as a safety guard.

## Notes

- The default cache path is `C:\Windows\ccmcache`, but the script always reads the actual location from CIM.
- The deletion spine uses the supported `UIResource.UIResourceMgr` COM interface — the same API the CCM client uses internally — so disk and the cache index stay consistent.
- Uses CIM cmdlets only for read/filter (no `Get-WmiObject`), so the script runs unchanged on PowerShell 5.1 and 7+.
- Provided as-is. Not supported by Microsoft.

## Contributing

Bug reports and pull requests welcome.

## Acknowledgements

Special thanks to [@theruck242](https://github.com/theruck242) for identifying a critical issue and providing a robust solution to enhance the script's safety and reliability.
