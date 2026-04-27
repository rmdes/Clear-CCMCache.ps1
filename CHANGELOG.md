# Changelog

This project follows [Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-04-27

A ground-up rewrite of the original `VirtualOx/Clear-CCMCache.ps1` script. **Breaking** for anyone calling the v1 surface; behavior is otherwise compatible with a default invocation (`.\Clear-CCMCache.ps1`).

### Breaking changes

- **Removed `-Detect`** in favor of standard `-WhatIf`.
- **Removed `-Help`** — use `Get-Help .\Clear-CCMCache.ps1 -Detailed`.
- **Requires CcmExec service running.** The script now binds to the supported CCM client COM interface and refuses to run if the client isn't healthy.

### Added

- **Supported deletion path.** `UIResource.UIResourceMgr.GetCacheInfo().DeleteCacheElement()` replaces the v1 manual disk + WMI surgery. Disk and the cache index stay consistent; locked items are respected; ccmexec's internal state stays in sync.
- **Safety filters with overrides.**
  - `PersistInCache=$true` entries are skipped by default (use `-IncludePersisted` to reclaim).
  - `ReferenceCount > 0` entries are skipped by default (use `-IncludeInUse` to override).
- **`-MaxSizeMB` target-size mode.** After the days-based pass, if the cache is still over target, removes "recent" entries oldest-first until under target or no more eligible candidates remain.
- **CMTrace logging.** Default `%SystemRoot%\CCM\Logs\ClearCache.log`. Auto-rotates at 10 MB. Configurable with `-LogPath`, disabled with `-NoLog`.
- **Structured records.** Every action emits a `[PSCustomObject]@{ Timestamp, Status, Path, SizeKB, Source, Reason, Error }`. `-PassThru` exposes the records on the success stream for piping.
- **Cache utilization reporting.** Disk size vs configured `CacheConfig.Size`, percent used, headroom — at run start and (when real removals occur) post-cleanup.
- **Honest reclaim accounting.** After any real removal, re-walks the cache root to report actual disk reclaim, separately from CIM/COM `ContentSize` accounting (which underestimates when content is extracted/staged inside cache folders).
- **Distinct exit codes.** `0` clean / `1` preflight failed / `2` per-item failures / `3` `-MaxSizeMB` target not met. Enables proper SCCM, Intune, scheduled-task integration.
- **`Examples/` folder.** Intune Proactive Remediation pair (detect + remediate), scheduled-task installer, deployment recipes.
- **Locale-invariant formatting.** Sizes and percentages render the same regardless of operator locale (`10.96 GB`, never `10,96 GB`).
- **`Test-PathUnder` safety guard.** Every `Remove-Item` is bounded to paths under the resolved cache root.
- **`ValidateRange` on `-Days`** prevents negative or absurd thresholds.
- **PS 5.1 + PS 7+ support.** Switched WMI cmdlets (`Get-WmiObject`/`Remove-WmiObject`, removed in PS 7) to CIM cmdlets.
- **Orphan reconciliation** between disk and CIM index, with case-insensitive path matching (CIM and `Get-ChildItem` return different casings on the same machine).
- **CLAUDE.md** with architecture notes and load-bearing constraints for future contributors.

### Fixed

- **`$_.Location` string interpolation** — v1 logged `System.Management...Location` instead of the actual path (4 occurrences).
- **COM-only ghost cleanup.** Cache elements left over from prior unsupported cleanups (folder deleted, CIM record manually removed, but ccmexec still tracking) are now detected and reaped via the supported COM path.
- **GUID normalization for the COM↔CIM join.** COM `CacheElementId` is `{UPPER-WITH-BRACES}`; CIM `CacheId` is `lower-no-braces`. Same identifier, different format. Both are normalized through `[Guid]` before matching.
- **Dual-counting in summary** when a `Skipped` entry is later promoted to `Removed` by the `-MaxSizeMB` bonus pass.
- **`-WhatIf` leaking into log writes.** `Add-Content`, `Move-Item`, and `New-Item` for log file management now run with `-WhatIf:$false` — logging is observability, not a destructive action.

### Notes

- This release is the result of a multi-pass collaborative review with Claude Code, validated end-to-end on a real CCM-managed Windows endpoint.
- The orphan-pass disk enumeration deliberately omits `-Force` to avoid reaping CCM-managed marker files (`skpswi.dat`) and any future hidden scratch directories.

## [1.x] - pre-2026

See `git log abb7526` and earlier for the upstream v1 history (WMI-based, days-only threshold, manual disk + index removal).

[2.0.0]: https://github.com/rmdes/Clear-CCMCache.ps1/releases/tag/v2.0.0
