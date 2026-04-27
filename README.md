# Clear-CCMCache

PowerShell script that removes old, unused content from the SCCM/CCM client cache (`CCMCache`) using the supported `UIResource.UIResourceMgr` COM interface, and reconciles legacy orphans between disk and the CIM index.

## Requirements

- Windows PowerShell **5.1** or PowerShell **7+**
- Administrative privileges
- Microsoft Endpoint Configuration Manager (SCCM/CCM) client installed and **CcmExec service running**

## Usage

```powershell
.\Clear-CCMCache.ps1 [-Days <int>] [-MaxSizeMB <int>] `
                     [-IncludePersisted] [-IncludeInUse] `
                     [-PassThru] [-LogPath <string>] [-NoLog] `
                     [-WhatIf] [-Confirm] [-Verbose]
```

| Parameter           | Description |
| ------------------- | --- |
| `-Days`             | Days an item must be unreferenced before removal. Default `30`, range `1`-`3650`. |
| `-MaxSizeMB`        | Target maximum cache size in MB. If after the days-based pass the cache is still over this size, removes oldest "recent" entries until under target. `0` (default) disables. Persisted/in-use filters still apply. |
| `-IncludePersisted` | Also remove entries flagged `PersistInCache=$true`. Off by default — deleting persisted content typically triggers redownload on the next policy evaluation. |
| `-IncludeInUse`     | Also attempt to remove entries with `ReferenceCount > 0`. Off by default — in-use content is being read by a running install or task sequence. The COM interface may still refuse the delete. |
| `-PassThru`         | Emit per-item records (`Timestamp, Status, Path, SizeKB, Source, Reason, Error`) on the success stream. Off by default to keep host output quiet. |
| `-LogPath`          | Override the CMTrace log file location. Default: read from CCM client registry, fall back to `%SystemRoot%\CCM\Logs\ClearCache.log`, then to `%TEMP%\ClearCache.log`. |
| `-NoLog`            | Disable file logging entirely. |
| `-WhatIf`           | Show what would be removed without making any changes. Reports projected reclaimable size. |
| `-Confirm`          | Prompt before each removal. |
| `-Verbose`          | Per-item progress and a summary line on the verbose stream. |

### Examples

```powershell
# Default: clean items unreferenced for more than 30 days
.\Clear-CCMCache.ps1

# Preview a 14-day cleanup, no changes made; shows projected reclaimable size
.\Clear-CCMCache.ps1 -Days 14 -WhatIf

# Run a 14-day cleanup with verbose output
.\Clear-CCMCache.ps1 -Days 14 -Verbose

# Shrink the cache to 8 GB, removing oldest recent entries if days-based isn't enough
.\Clear-CCMCache.ps1 -MaxSizeMB 8192 -WhatIf -Verbose

# Clean and export a per-item audit record
.\Clear-CCMCache.ps1 -PassThru | Export-Csv .\cleanup.csv -NoTypeInformation

# Full help
Get-Help .\Clear-CCMCache.ps1 -Detailed
```

### Logging

Every action (Removed / WouldRemove / Skipped / Failed) is appended to a CMTrace-formatted log alongside the rest of the CCM client logs (default `%SystemRoot%\CCM\Logs\ClearCache.log`). Open it with `CMTrace.exe` to see runs interleaved with `CcmExec.log`, `CAS.log`, etc. The log auto-rotates to `ClearCache.lo_` when it crosses 10 MB.

Each run also logs the **cache utilization** at start: actual on-disk size, configured maximum (read from `CacheConfig.Size`), percent used, and headroom — the headline numbers an admin needs to decide whether a cleanup is even necessary.

The verbose summary line at the end of every run looks like:

```
Cache utilization: 850.42 MB on disk / 25.00 GB max (3.3% used, 24.17 GB headroom)
...
Run finished: removed=33 reclaimed=4.21 GB skipped=7 failed=0
```

In `-WhatIf` mode it also reports `would-remove=N projected=N` so you can preview the savings before committing.

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

## Exit codes

| Code | Meaning |
| --- | --- |
| `0` | Clean run. No per-item failures; if `-MaxSizeMB` was set, the cache is now under target. |
| `1` | Preflight failed (no CCM client, CcmExec not running, cache path missing, COM bind failed). |
| `2` | Run completed but one or more removals failed (`failed > 0` in the summary). |
| `3` | `-MaxSizeMB` was set and the cache is still over target after the run. |

Use these for SCCM / Intune / scheduled-task conditional logic. See [`Examples/`](Examples/) for deployment templates.

## Deployment

See the [`Examples/`](Examples/) folder for ready-to-use templates:

- **`Intune-Detect.ps1` + `Intune-Remediate.ps1`** — Proactive Remediation script pair.
- **`Install-ScheduledTask.ps1`** — registers a weekly Windows scheduled task running as SYSTEM.

## Notes

- The default cache path is `C:\Windows\ccmcache`, but the script always reads the actual location from CIM.
- The deletion spine uses the supported `UIResource.UIResourceMgr` COM interface — the same API the CCM client uses internally — so disk and the cache index stay consistent.
- Uses CIM cmdlets only for read/filter (no `Get-WmiObject`), so the script runs unchanged on PowerShell 5.1 and 7+.
- Provided as-is. Not supported by Microsoft.

## Contributing

Bug reports and pull requests welcome.

## Acknowledgements

Special thanks to [@theruck242](https://github.com/theruck242) for identifying a critical issue and providing a robust solution to enhance the script's safety and reliability.
