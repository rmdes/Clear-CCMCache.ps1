# ABOUTME: Removes old/unused content and reconciles orphaned entries in the SCCM/CCM client cache.
# ABOUTME: Uses CIM cmdlets, so it runs unchanged on Windows PowerShell 5.1 and PowerShell 7+.

<#
.SYNOPSIS
Clears old and unused content from the CCMCache folder.

.DESCRIPTION
Removes cache entries whose LastReferenced timestamp is older than -Days (folder on disk
plus the matching CIM record), then reconciles orphans between disk and the CCM client
store: disk folders with no CIM record are removed, and CIM records pointing at missing
folders are removed.

Use -WhatIf to preview without making changes, or -Confirm to be prompted per item.

.PARAMETER Days
Number of days an item must be unreferenced before it is considered stale. Default: 30.

.EXAMPLE
.\Clear-CCMCache.ps1
Cleans entries unreferenced for more than 30 days (default).

.EXAMPLE
.\Clear-CCMCache.ps1 -Days 14 -WhatIf
Previews what a 14-day cleanup would remove, without changing anything.

.EXAMPLE
.\Clear-CCMCache.ps1 -Days 14 -Verbose
Runs a 14-day cleanup with per-item progress on the verbose stream.

.NOTES
Requires administrative privileges and a working CCM client.
Run `Get-Help .\Clear-CCMCache.ps1 -Detailed` for full help.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateRange(1, 3650)]
    [int]$Days = 30
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

# Resolve the cache root from the CCM client
try {
    $CacheConfig = Get-CimInstance -Namespace $CcmNamespace -ClassName CacheConfig -Filter "ConfigKey='Cache'"
} catch {
    Write-Error "Failed to query CacheConfig in $CcmNamespace. Is the CCM client installed? $_"
    exit 1
}

$CachePath = $CacheConfig.Location
if ([string]::IsNullOrWhiteSpace($CachePath)) {
    Write-Error "CCM cache path is empty."
    exit 1
}
if (-not (Test-Path -LiteralPath $CachePath)) {
    Write-Error "CCM cache path '$CachePath' does not exist on disk."
    exit 1
}
Write-Verbose "CCM cache path: $CachePath"

# Pull the cache index once
try {
    $CacheEntries = @(Get-CimInstance -Namespace $CcmNamespace -ClassName CacheInfoEx)
} catch {
    Write-Error "Failed to query CacheInfoEx. $_"
    exit 1
}

$Now = Get-Date
$Stale = $CacheEntries | Where-Object { ($Now - $_.LastReferenced).Days -gt $Days }

if ($Stale) {
    Write-Verbose "Found $($Stale.Count) stale cache entries (older than $Days days)."
    foreach ($entry in $Stale) {
        $location = $entry.Location
        $target = "$location (last referenced $($entry.LastReferenced))"
        if (-not $PSCmdlet.ShouldProcess($target, 'Remove stale cache entry')) { continue }

        if (-not (Test-PathUnder -Path $location -Parent $CachePath)) {
            Write-Warning "Skipping '$location' — not under cache root '$CachePath'."
            continue
        }

        if (Test-Path -LiteralPath $location) {
            try {
                Remove-Item -LiteralPath $location -Recurse -Force -ErrorAction Stop
                Write-Verbose "Deleted folder: $location"
            } catch {
                Write-Warning "Failed to delete folder '$location': $_"
            }
        }

        try {
            Remove-CimInstance -InputObject $entry -ErrorAction Stop
            Write-Verbose "Removed CIM entry: $location"
        } catch {
            Write-Warning "Failed to remove CIM entry for '$location': $_"
        }
    }
} else {
    Write-Verbose "No stale cache entries."
}

# Reconcile orphans against current state
try {
    $RemainingCim = @(Get-CimInstance -Namespace $CcmNamespace -ClassName CacheInfoEx)
} catch {
    Write-Error "Failed to re-query CacheInfoEx for orphan reconciliation. $_"
    exit 1
}

try {
    $DiskFolders = @(Get-ChildItem -LiteralPath $CachePath -Directory -ErrorAction Stop |
        Select-Object -ExpandProperty FullName)
} catch {
    Write-Error "Failed to enumerate folders under '$CachePath'. $_"
    exit 1
}

$cmp = [System.StringComparer]::OrdinalIgnoreCase
$CimSet  = [System.Collections.Generic.HashSet[string]]::new([string[]]@($RemainingCim | ForEach-Object { $_.Location }), $cmp)
$DiskSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$DiskFolders, $cmp)

# Disk folders with no CIM record
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

# CIM records pointing at folders that no longer exist
foreach ($entry in $RemainingCim) {
    $location = $entry.Location
    if ($DiskSet.Contains($location)) { continue }
    if (-not $PSCmdlet.ShouldProcess($location, 'Remove orphaned CIM record (folder missing)')) { continue }
    try {
        Remove-CimInstance -InputObject $entry -ErrorAction Stop
        Write-Verbose "Removed orphaned CIM entry: $location"
    } catch {
        Write-Warning "Failed to remove orphaned CIM entry for '$location': $_"
    }
}
