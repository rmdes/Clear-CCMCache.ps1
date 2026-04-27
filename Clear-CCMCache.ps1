# ABOUTME: Removes old/unused content from the SCCM/CCM client cache via the supported COM interface.
# ABOUTME: Honors PersistInCache and ReferenceCount; runs on Windows PowerShell 5.1 and PowerShell 7+.

<#
.SYNOPSIS
Clears old and unused content from the CCMCache folder.

.DESCRIPTION
Removes cache entries whose LastReferenced timestamp is older than -Days using the
supported UIResource.UIResourceMgr COM interface — the same API the CCM client uses
internally — so disk and the cache index stay consistent and locked items are
respected.

By default, the script skips:
  * Persisted entries (PersistInCache = $true) — deleting them triggers redownload on
    the next policy evaluation, which at scale becomes a bandwidth event.
  * In-use entries (ReferenceCount > 0) — they are being read by a running install
    or task sequence and deleting them mid-flight will break the deployment.

Override these defaults with -IncludePersisted and -IncludeInUse.

After the main pass, a best-effort orphan reconciliation runs to clean up legacy
inconsistencies between disk folders and the CIM index (typically left behind by
older, unsupported cleanup scripts or interrupted operations).

Use -WhatIf to preview, or -Confirm to be prompted per item.

.PARAMETER Days
Number of days an item must be unreferenced before it is considered stale. Default: 30.

.PARAMETER IncludePersisted
Also remove entries flagged PersistInCache = $true. Off by default.

.PARAMETER IncludeInUse
Also attempt to remove entries with ReferenceCount > 0. Off by default. The COM
interface may still refuse the delete if the content is actively locked.

.EXAMPLE
.\Clear-CCMCache.ps1
Cleans entries unreferenced for more than 30 days, skipping persisted/in-use.

.EXAMPLE
.\Clear-CCMCache.ps1 -Days 14 -WhatIf
Previews a 14-day cleanup without making changes.

.EXAMPLE
.\Clear-CCMCache.ps1 -Days 7 -IncludePersisted -Verbose
Aggressive cleanup including persisted content (will likely trigger redownloads).

.NOTES
Requires administrative privileges and a running CCM client service (CcmExec).
Run `Get-Help .\Clear-CCMCache.ps1 -Detailed` for full help.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateRange(1, 3650)]
    [int]$Days = 30,

    [switch]$IncludePersisted,
    [switch]$IncludeInUse
)

$ErrorActionPreference = 'Stop'
$CcmNamespace = 'ROOT\ccm\SoftMgmtAgent'

function Test-PathUnder {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Parent
    )
    try {
        $p = [System.IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
        $r = [System.IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
        return $p.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function Stop-WithError {
    param([Parameter(Mandatory)][string]$Message)
    Write-Error -ErrorAction Continue -Message $Message
    exit 1
}

# COM CacheElementId comes back as '{UPPER-WITH-BRACES}'; CIM CacheId as 'lower-no-braces'.
# Normalize both sides through [Guid] so the join key is identical.
function ConvertTo-NormalizedGuid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $g = [Guid]::Empty
    if ([Guid]::TryParse($Value, [ref]$g)) { return $g.ToString('D') }
    return $null
}

# --- preflight ---------------------------------------------------------------

$svc = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
if (-not $svc) {
    Stop-WithError "CCM client service (CcmExec) is not installed on this machine."
}
if ($svc.Status -ne 'Running') {
    Stop-WithError "CCM client service (CcmExec) is '$($svc.Status)'. Start it before running — the supported COM interface needs a live client."
}

try {
    $CacheConfig = Get-CimInstance -Namespace $CcmNamespace -ClassName CacheConfig -Filter "ConfigKey='Cache'"
} catch {
    Stop-WithError "Failed to query CacheConfig in $CcmNamespace. Is the CCM client healthy? $_"
}

$CachePath = $CacheConfig.Location
if ([string]::IsNullOrWhiteSpace($CachePath)) {
    Stop-WithError "CCM cache path is empty."
}
if (-not (Test-Path -LiteralPath $CachePath)) {
    Stop-WithError "CCM cache path '$CachePath' does not exist on disk."
}
Write-Verbose "CCM cache path: $CachePath"

# --- bind to the supported COM interface -------------------------------------

try {
    $UIResMgr = New-Object -ComObject UIResource.UIResourceMgr
    $CacheCom = $UIResMgr.GetCacheInfo()
} catch {
    Stop-WithError "Failed to bind to UIResource.UIResourceMgr. Is the CCM client healthy? $_"
}

# --- gather state from COM (deletion handle) and CIM (filter properties) -----
# COM CacheElement exposes ID, Location, LastReferenceTime, ReferenceCount but
# NOT PersistInCache — that flag lives only on the CIM CacheInfoEx class.

try {
    $ComElements = @($CacheCom.GetCacheElements())
} catch {
    Stop-WithError "Failed to enumerate cache elements via COM. $_"
}

try {
    $CimEntries = @(Get-CimInstance -Namespace $CcmNamespace -ClassName CacheInfoEx)
} catch {
    Stop-WithError "Failed to query CacheInfoEx for filter metadata. $_"
}

$CimById = @{}
foreach ($e in $CimEntries) {
    $key = ConvertTo-NormalizedGuid $e.CacheId
    if ($key) { $CimById[$key] = $e }
}
Write-Verbose "Indexed $($CimById.Count) CIM cache entries; enumerated $($ComElements.Count) COM cache elements."

# --- main pass ---------------------------------------------------------------

$Now = Get-Date
$Removed = 0
$SkippedRecent = 0
$SkippedPersisted = 0
$SkippedInUse = 0
$ComOnly = 0
$Failed = 0

foreach ($com in $ComElements) {
    $id       = $com.CacheElementId
    $location = $com.Location
    $key      = ConvertTo-NormalizedGuid $id
    $cim      = if ($key) { $CimById[$key] } else { $null }

    # Use CIM filter properties when available; fall back to COM's own when the CIM
    # record is missing. A missing CIM twin means the element is a "COM-only ghost":
    # ccmexec still tracks it but the index doesn't, typically from prior unsupported
    # cleanups. DeleteCacheElement is the only way to flush ccmexec's internal state.
    if ($cim) {
        $lastRef  = $cim.LastReferenced
        $refCount = $cim.ReferenceCount
        $persist  = [bool]$cim.PersistInCache
        $source   = 'cim'
    } else {
        $lastRef  = $com.LastReferenceTime
        $refCount = $com.ReferenceCount
        $persist  = $false
        $source   = 'com-only'
        $ComOnly++
    }

    $idleDays = ($Now - $lastRef).Days
    if ($idleDays -le $Days) {
        $SkippedRecent++
        continue
    }

    if ($persist -and -not $IncludePersisted) {
        Write-Verbose "Skipping persisted: $location (use -IncludePersisted to override)."
        $SkippedPersisted++
        continue
    }

    if ($refCount -gt 0 -and -not $IncludeInUse) {
        Write-Verbose "Skipping in-use: $location (ref=$refCount; use -IncludeInUse to override)."
        $SkippedInUse++
        continue
    }

    $target = "$location (idle $idleDays d, ref=$refCount, persist=$persist, src=$source)"
    if (-not $PSCmdlet.ShouldProcess($target, 'Delete cache element via CCM COM interface')) { continue }

    try {
        $CacheCom.DeleteCacheElement($id)
        Write-Verbose "Deleted: $location"
        $Removed++
    } catch {
        Write-Warning "DeleteCacheElement refused/failed for '$location' (id=$id, src=$source): $_"
        $Failed++
    }
}

Write-Verbose ("Main pass: removed={0} recent={1} persisted={2} in-use={3} com-only-ghosts={4} failed={5}" -f `
    $Removed, $SkippedRecent, $SkippedPersisted, $SkippedInUse, $ComOnly, $Failed)

# --- orphan reconciliation (legacy/interrupted-state cleanup) ----------------

try {
    $RemainingCim = @(Get-CimInstance -Namespace $CcmNamespace -ClassName CacheInfoEx)
} catch {
    Write-Warning "Skipping orphan reconciliation — failed to re-query CacheInfoEx: $_"
    return
}

try {
    $DiskFolders = @(Get-ChildItem -LiteralPath $CachePath -Directory -ErrorAction Stop |
        Select-Object -ExpandProperty FullName)
} catch {
    Write-Warning "Skipping orphan reconciliation — failed to enumerate folders under '$CachePath': $_"
    return
}

$cmp     = [System.StringComparer]::OrdinalIgnoreCase
$CimSet  = [System.Collections.Generic.HashSet[string]]::new([string[]]@($RemainingCim | ForEach-Object { $_.Location }), $cmp)
$DiskSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$DiskFolders, $cmp)

# Disk folders with no CIM record — safe to delete: nothing in the index claims them.
foreach ($folder in $DiskFolders) {
    if ($CimSet.Contains($folder)) { continue }
    if (-not $PSCmdlet.ShouldProcess($folder, 'Remove orphaned folder (no CIM record)')) { continue }
    if (-not (Test-PathUnder -Path $folder -Parent $CachePath)) {
        Write-Warning "Skipping '$folder' — not under cache root '$CachePath'."
        continue
    }
    try {
        Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
        Write-Verbose "Removed orphaned folder: $folder"
    } catch {
        Write-Warning "Failed to remove orphaned folder '$folder': $_"
    }
}

# CIM records with no folder — try the supported COM path first; fall back to manual.
foreach ($entry in $RemainingCim) {
    $location = $entry.Location
    if ($DiskSet.Contains($location)) { continue }
    if (-not $PSCmdlet.ShouldProcess($location, 'Remove orphaned CIM record (folder missing)')) { continue }
    try {
        $CacheCom.DeleteCacheElement($entry.CacheId)
        Write-Verbose "Removed orphaned cache element via COM: $location"
    } catch {
        try {
            Remove-CimInstance -InputObject $entry -ErrorAction Stop
            Write-Verbose "Removed orphaned CIM record (COM refused, manual fallback): $location"
        } catch {
            Write-Warning "Failed to remove orphaned CIM entry for '$location': $_"
        }
    }
}
