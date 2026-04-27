# ABOUTME: Removes old/unused content from the SCCM/CCM client cache via the supported COM interface.
# ABOUTME: Honors PersistInCache and ReferenceCount; runs on Windows PowerShell 5.1 and PowerShell 7+.

<#
.SYNOPSIS
Clears old and unused content from the CCMCache folder.

.DESCRIPTION
Removes cache entries whose LastReferenced timestamp is older than -Days using the
supported UIResource.UIResourceMgr COM interface - the same API the CCM client uses
internally - so disk and the cache index stay consistent and locked items are
respected.

By default, the script skips:
  * Persisted entries (PersistInCache = $true) - deleting them triggers redownload on
    the next policy evaluation, which at scale becomes a bandwidth event.
  * In-use entries (ReferenceCount > 0) - they are being read by a running install
    or task sequence and deleting them mid-flight will break the deployment.

Override these defaults with -IncludePersisted and -IncludeInUse.

After the main pass, a best-effort orphan reconciliation runs to clean up legacy
inconsistencies between disk folders and the CIM index.

Every action is logged to a CMTrace-compatible file (default
%SystemRoot%\CCM\Logs\ClearCache.log) so runs interleave with CcmExec.log etc.
in CMTrace.exe. With -PassThru, per-item records are also emitted on the success
stream for piping to Export-Csv or further processing.

Use -WhatIf to preview, or -Confirm to be prompted per item.

.PARAMETER Days
Number of days an item must be unreferenced before it is considered stale. Default: 30.

.PARAMETER MaxSizeMB
Target maximum cache size in MB. After the days-based pass, if the cache is still
projected to exceed this size, remove "recent" entries oldest-first until under target
or no more eligible candidates remain. 0 (default) disables the bonus pass.
Persisted and in-use filters still apply — set -IncludePersisted/-IncludeInUse to
override them. Useful when the cache is over its configured cap and the days-based
filter alone can't free enough space.

.PARAMETER IncludePersisted
Also remove entries flagged PersistInCache = $true. Off by default.

.PARAMETER IncludeInUse
Also attempt to remove entries with ReferenceCount > 0. Off by default. The COM
interface may still refuse the delete if the content is actively locked.

.PARAMETER PassThru
Emit per-item records (Timestamp, Status, Path, SizeKB, Source, Reason, Error) on
the success stream. Off by default to keep host output quiet.

.PARAMETER LogPath
Override the CMTrace log file location. Default: read from CCM client registry, fall
back to %SystemRoot%\CCM\Logs\ClearCache.log, then to %TEMP%\ClearCache.log.

.PARAMETER NoLog
Disable file logging entirely. Records are still kept in memory for the summary and
-PassThru output.

.EXAMPLE
.\Clear-CCMCache.ps1
Cleans entries unreferenced for more than 30 days, skipping persisted/in-use.

.EXAMPLE
.\Clear-CCMCache.ps1 -Days 14 -WhatIf
Previews a 14-day cleanup without making changes; shows projected reclaimable size.

.EXAMPLE
.\Clear-CCMCache.ps1 -PassThru | Export-Csv .\cleanup.csv -NoTypeInformation
Cleans and exports a per-item audit record to CSV.

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

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxSizeMB = 0,

    [switch]$IncludePersisted,
    [switch]$IncludeInUse,
    [switch]$PassThru,

    [string]$LogPath,
    [switch]$NoLog
)

$ErrorActionPreference = 'Stop'
$CcmNamespace = 'ROOT\ccm\SoftMgmtAgent'

# --- helpers -----------------------------------------------------------------

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

# COM CacheElementId is '{UPPER-WITH-BRACES}'; CIM CacheId is 'lower-no-braces'.
# Normalize through [Guid] so the join key is identical on both sides.
function ConvertTo-NormalizedGuid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $g = [Guid]::Empty
    if ([Guid]::TryParse($Value, [ref]$g)) { return $g.ToString('D') }
    return $null
}

function Get-CcmLogPath {
    try {
        $reg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global' -ErrorAction Stop
        if ($reg.LogDirectory -and (Test-Path -LiteralPath $reg.LogDirectory)) {
            return Join-Path $reg.LogDirectory 'ClearCache.log'
        }
    } catch { }
    $default = Join-Path $env:SystemRoot 'CCM\Logs'
    if (Test-Path -LiteralPath $default) { return Join-Path $default 'ClearCache.log' }
    return Join-Path $env:TEMP 'ClearCache.log'
}

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Severity = 'Info'
    )
    if ($Script:NoLog -or $Script:LogDisabled) { return }
    $type = switch ($Severity) { 'Info' { 1 } 'Warning' { 2 } 'Error' { 3 } }
    $now  = [datetime]::Now
    $tz   = [int][System.TimeZoneInfo]::Local.GetUtcOffset($now).TotalMinutes
    $tzs  = "{0:+000;-000}" -f $tz
    $time = "$($now.ToString('HH:mm:ss.fff'))$tzs"
    $date = $now.ToString('MM-dd-yyyy')
    $line = "<![LOG[$Message]LOG]!><time=`"$time`" date=`"$date`" component=`"ClearCCMCache`" context=`"`" type=`"$type`" thread=`"$PID`" file=`"`">"
    try {
        # -WhatIf:$false: logging is observability, not a destructive action.
        Add-Content -LiteralPath $Script:LogPath -Value $line -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Warning "CMTrace log write failed for '$Script:LogPath' - disabling further log writes. $_"
        $Script:LogDisabled = $true
    }
}

function Get-FolderSizeKB {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $bytes = (Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum).Sum
        if (-not $bytes) { return 0 }
        return [int]($bytes / 1KB)
    } catch { return 0 }
}

# Use InvariantCulture so logs are consistent regardless of the operator's locale —
# PowerShell's -f operator defaults to CurrentCulture, "$var" interpolation to
# Invariant, and that asymmetry makes mixed output (10,96 GB ... 109.6%) ugly.
$Script:InvCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Format-Size {
    param([long]$KB)
    if ($KB -ge 1GB / 1KB) { return [string]::Format($Script:InvCulture, '{0:N2} GB', $KB / (1GB / 1KB)) }
    if ($KB -ge 1MB / 1KB) { return [string]::Format($Script:InvCulture, '{0:N2} MB', $KB / (1MB / 1KB)) }
    return [string]::Format($Script:InvCulture, '{0} KB', $KB)
}

function Format-Utilization {
    param(
        [Parameter(Mandatory)][long]$DiskKB,
        [int]$ConfiguredMaxMB
    )
    $disk = Format-Size $DiskKB
    if ($ConfiguredMaxMB -gt 0) {
        $maxKB = [long]$ConfiguredMaxMB * 1024
        $pct = if ($maxKB -gt 0) { [math]::Round(($DiskKB / $maxKB) * 100, 1) } else { 0 }
        $pctStr = $pct.ToString($Script:InvCulture)
        $headroomKB = [Math]::Max(0, $maxKB - $DiskKB)
        return "$disk on disk / $(Format-Size $maxKB) max ($pctStr% used, $(Format-Size $headroomKB) headroom)"
    }
    return "$disk on disk (max not configured)"
}

function Add-Record {
    param(
        [Parameter(Mandatory)][ValidateSet('Removed', 'WouldRemove', 'Skipped', 'Failed')][string]$Status,
        [Parameter(Mandatory)][string]$Path,
        [int]$SizeKB = 0,
        [string]$Source = '',
        [string]$Reason = '',
        [string]$ErrorMsg = ''
    )
    $rec = [PSCustomObject]@{
        Timestamp = [datetime]::Now
        Status    = $Status
        Path      = $Path
        SizeKB    = $SizeKB
        Source    = $Source
        Reason    = $Reason
        Error     = $ErrorMsg
    }
    $Script:Records.Add($rec)

    $sev = if ($Status -eq 'Failed') { 'Warning' } else { 'Info' }
    $msg = "$Status $Path"
    if ($Reason)   { $msg += " [$Reason]" }
    if ($SizeKB)   { $msg += " (${SizeKB} KB)" }
    if ($ErrorMsg) { $msg += " - $ErrorMsg" }

    Write-CMTraceLog -Message $msg -Severity $sev
    Write-Verbose $msg

    if ($Script:PassThru) { Write-Output $rec }
}

# --- preflight ---------------------------------------------------------------

$svc = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
if (-not $svc) {
    Stop-WithError "CCM client service (CcmExec) is not installed on this machine."
}
if ($svc.Status -ne 'Running') {
    Stop-WithError "CCM client service (CcmExec) is '$($svc.Status)'. Start it before running - the supported COM interface needs a live client."
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

# --- initialize logging and state -------------------------------------------

$Script:NoLog       = [bool]$NoLog
$Script:LogDisabled = $false
$Script:PassThru    = [bool]$PassThru
$Script:Records     = [System.Collections.Generic.List[object]]::new()

if (-not $Script:NoLog) {
    if (-not $LogPath) { $LogPath = Get-CcmLogPath }
    $logDir = Split-Path -Parent $LogPath
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        try { New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false -Confirm:$false | Out-Null } catch { }
    }
    if (Test-Path -LiteralPath $LogPath) {
        try {
            if ((Get-Item -LiteralPath $LogPath).Length -gt 10MB) {
                $rolled = [System.IO.Path]::ChangeExtension($LogPath, '.lo_')
                Move-Item -LiteralPath $LogPath -Destination $rolled -Force -WhatIf:$false -Confirm:$false
            }
        } catch { Write-Warning "Log rotation failed: $_" }
    }
    $Script:LogPath = $LogPath
}

Write-Verbose "CCM cache path: $CachePath"
Write-Verbose "Log file: $(if ($Script:NoLog) { '(disabled)' } else { $Script:LogPath })"

$whatIf = $WhatIfPreference -eq $true -or $PSBoundParameters.ContainsKey('WhatIf')
Write-CMTraceLog -Message "=== Run started: Days=$Days IncludePersisted=$IncludePersisted IncludeInUse=$IncludeInUse WhatIf=$whatIf CachePath=$CachePath ===" -Severity Info

# Cache utilization snapshot — reads on-disk size (authoritative) and the configured
# max from CacheConfig.Size (in MB). The walk can take a moment on large caches.
$ConfiguredMaxMB = [int]($CacheConfig.Size)
$StartDiskKB = Get-FolderSizeKB -Path $CachePath
$utilLine = "Cache utilization: $(Format-Utilization -DiskKB $StartDiskKB -ConfiguredMaxMB $ConfiguredMaxMB)"
Write-Verbose $utilLine
Write-CMTraceLog -Message $utilLine -Severity Info

# --- bind to the supported COM interface -------------------------------------

try {
    $UIResMgr = New-Object -ComObject UIResource.UIResourceMgr
    $CacheCom = $UIResMgr.GetCacheInfo()
} catch {
    Stop-WithError "Failed to bind to UIResource.UIResourceMgr. Is the CCM client healthy? $_"
}

# --- gather state from COM (deletion handle) and CIM (filter properties) -----
# COM CacheElement exposes ID, Location, LastReferenceTime, ReferenceCount, ContentSize
# but NOT PersistInCache - that flag lives only on CIM CacheInfoEx.

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

# Recent-but-deletable candidates collected for the optional -MaxSizeMB bonus pass.
# Persisted/in-use entries are NOT collected: a hard size target should not override
# safety filters that prevent breaking running deployments.
$RecentCandidates = [System.Collections.Generic.List[object]]::new()

foreach ($com in $ComElements) {
    $id       = $com.CacheElementId
    $location = $com.Location
    $key      = ConvertTo-NormalizedGuid $id
    $cim      = if ($key) { $CimById[$key] } else { $null }

    if ($cim) {
        $lastRef  = $cim.LastReferenced
        $refCount = $cim.ReferenceCount
        $persist  = [bool]$cim.PersistInCache
        $sizeKB   = [int]$cim.ContentSize
        $source   = 'cim'
    } else {
        $lastRef  = $com.LastReferenceTime
        $refCount = $com.ReferenceCount
        $persist  = $false
        $sizeKB   = [int]$com.ContentSize
        $source   = 'com-only'
    }

    $idleDays = ($Now - $lastRef).Days
    if ($idleDays -le $Days) {
        $eligibleForMaxSize = -not ($persist -and -not $IncludePersisted) -and -not ($refCount -gt 0 -and -not $IncludeInUse)
        if ($MaxSizeMB -gt 0 -and $eligibleForMaxSize) {
            $RecentCandidates.Add([PSCustomObject]@{
                ComElement     = $com
                Location       = $location
                LastReferenced = $lastRef
                IdleDays       = $idleDays
                SizeKB         = $sizeKB
                Source         = $source
            })
        }
        Add-Record -Status 'Skipped' -Path $location -SizeKB $sizeKB -Source $source -Reason "recent (idle $idleDays d <= $Days)"
        continue
    }

    if ($persist -and -not $IncludePersisted) {
        Add-Record -Status 'Skipped' -Path $location -SizeKB $sizeKB -Source $source -Reason 'persisted (use -IncludePersisted)'
        continue
    }

    if ($refCount -gt 0 -and -not $IncludeInUse) {
        Add-Record -Status 'Skipped' -Path $location -SizeKB $sizeKB -Source $source -Reason "in-use ref=$refCount (use -IncludeInUse)"
        continue
    }

    $reason = "idle $idleDays d, ref=$refCount, persist=$persist, src=$source"
    $target = "$location ($reason)"

    if (-not $PSCmdlet.ShouldProcess($target, 'Delete cache element via CCM COM interface')) {
        Add-Record -Status 'WouldRemove' -Path $location -SizeKB $sizeKB -Source $source -Reason $reason
        continue
    }

    try {
        $CacheCom.DeleteCacheElement($id)
        Add-Record -Status 'Removed' -Path $location -SizeKB $sizeKB -Source $source -Reason $reason
    } catch {
        Add-Record -Status 'Failed' -Path $location -SizeKB $sizeKB -Source $source -Reason $reason -ErrorMsg $_.Exception.Message
    }
}

# --- orphan reconciliation (legacy/interrupted-state cleanup) ----------------

try {
    $RemainingCim = @(Get-CimInstance -Namespace $CcmNamespace -ClassName CacheInfoEx)
} catch {
    Write-Warning "Skipping orphan reconciliation - failed to re-query CacheInfoEx: $_"
    $RemainingCim = $null
}

try {
    $DiskFolders = @(Get-ChildItem -LiteralPath $CachePath -Directory -ErrorAction Stop |
        Select-Object -ExpandProperty FullName)
} catch {
    Write-Warning "Skipping orphan reconciliation - failed to enumerate folders under '$CachePath': $_"
    $DiskFolders = $null
}

if ($RemainingCim -ne $null -and $DiskFolders -ne $null) {
    $cmp     = [System.StringComparer]::OrdinalIgnoreCase
    $CimSet  = [System.Collections.Generic.HashSet[string]]::new([string[]]@($RemainingCim | ForEach-Object { $_.Location }), $cmp)
    $DiskSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$DiskFolders, $cmp)

    foreach ($folder in $DiskFolders) {
        if ($CimSet.Contains($folder)) { continue }
        $sizeKB = Get-FolderSizeKB -Path $folder
        if (-not (Test-PathUnder -Path $folder -Parent $CachePath)) {
            Add-Record -Status 'Skipped' -Path $folder -SizeKB $sizeKB -Source 'orphan-disk' -Reason 'outside cache root - refusing to delete'
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($folder, 'Remove orphaned folder (no CIM record)')) {
            Add-Record -Status 'WouldRemove' -Path $folder -SizeKB $sizeKB -Source 'orphan-disk' -Reason 'no CIM record'
            continue
        }
        try {
            Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
            Add-Record -Status 'Removed' -Path $folder -SizeKB $sizeKB -Source 'orphan-disk' -Reason 'no CIM record'
        } catch {
            Add-Record -Status 'Failed' -Path $folder -SizeKB $sizeKB -Source 'orphan-disk' -Reason 'no CIM record' -ErrorMsg $_.Exception.Message
        }
    }

    foreach ($entry in $RemainingCim) {
        $location = $entry.Location
        if ($DiskSet.Contains($location)) { continue }
        $sizeKB = [int]$entry.ContentSize
        if (-not $PSCmdlet.ShouldProcess($location, 'Remove orphaned CIM record (folder missing)')) {
            Add-Record -Status 'WouldRemove' -Path $location -SizeKB $sizeKB -Source 'orphan-cim' -Reason 'folder missing'
            continue
        }
        try {
            $CacheCom.DeleteCacheElement($entry.CacheId)
            Add-Record -Status 'Removed' -Path $location -SizeKB $sizeKB -Source 'orphan-cim' -Reason 'folder missing (COM)'
        } catch {
            try {
                Remove-CimInstance -InputObject $entry -ErrorAction Stop
                Add-Record -Status 'Removed' -Path $location -SizeKB $sizeKB -Source 'orphan-cim' -Reason 'folder missing (CIM fallback)'
            } catch {
                Add-Record -Status 'Failed' -Path $location -SizeKB $sizeKB -Source 'orphan-cim' -Reason 'folder missing' -ErrorMsg $_.Exception.Message
            }
        }
    }
}

# --- bonus pass: -MaxSizeMB target enforcement ------------------------------
# After the days-based main pass + orphan reconciliation, if the cache is still
# projected to exceed -MaxSizeMB, remove "recent" entries oldest-first until the
# target is met or no more eligible candidates remain. Persisted/in-use entries
# are still excluded — a hard size target should not break running deployments.

if ($MaxSizeMB -gt 0) {
    $maxKB = [long]$MaxSizeMB * 1024
    $reclaimedSoFarKB = (@($Script:Records | Where-Object Status -in 'Removed', 'WouldRemove') |
        Measure-Object SizeKB -Sum).Sum
    if (-not $reclaimedSoFarKB) { $reclaimedSoFarKB = 0 }
    $projectedDiskKB = $StartDiskKB - $reclaimedSoFarKB

    if ($projectedDiskKB -le $maxKB) {
        Write-Verbose "MaxSize: cache projected at $(Format-Size $projectedDiskKB), target $(Format-Size $maxKB) - no bonus pass needed."
    } elseif ($RecentCandidates.Count -eq 0) {
        Write-Warning "MaxSize: cache projected at $(Format-Size $projectedDiskKB) (target $(Format-Size $maxKB)) but no eligible recent candidates to remove. Consider -IncludePersisted/-IncludeInUse."
    } else {
        $excessKB = $projectedDiskKB - $maxKB
        $msg = "MaxSize: bonus pass - cache projected at $(Format-Size $projectedDiskKB), target $(Format-Size $maxKB), need to reclaim $(Format-Size $excessKB) more from $($RecentCandidates.Count) recent candidate(s)."
        Write-Verbose $msg
        Write-CMTraceLog -Message $msg -Severity Info

        $sorted = $RecentCandidates | Sort-Object LastReferenced
        $bonusReclaimedKB = 0
        foreach ($cand in $sorted) {
            if ($bonusReclaimedKB -ge $excessKB) { break }
            $reason = "max-size: idle $($cand.IdleDays) d, target $MaxSizeMB MB"
            $target = "$($cand.Location) ($reason)"
            if (-not $PSCmdlet.ShouldProcess($target, 'Delete cache element via CCM COM interface (max-size pass)')) {
                Add-Record -Status 'WouldRemove' -Path $cand.Location -SizeKB $cand.SizeKB -Source $cand.Source -Reason $reason
                $bonusReclaimedKB += $cand.SizeKB
                continue
            }
            try {
                $CacheCom.DeleteCacheElement($cand.ComElement.CacheElementId)
                Add-Record -Status 'Removed' -Path $cand.Location -SizeKB $cand.SizeKB -Source $cand.Source -Reason $reason
                $bonusReclaimedKB += $cand.SizeKB
            } catch {
                Add-Record -Status 'Failed' -Path $cand.Location -SizeKB $cand.SizeKB -Source $cand.Source -Reason $reason -ErrorMsg $_.Exception.Message
            }
        }

        if ($bonusReclaimedKB -lt $excessKB) {
            Write-Warning "MaxSize: bonus pass freed $(Format-Size $bonusReclaimedKB) but target needed $(Format-Size $excessKB). Cache remains over -MaxSizeMB."
        }
    }
}

# --- summary -----------------------------------------------------------------

$removed     = @($Script:Records | Where-Object Status -eq 'Removed')
$wouldRemove = @($Script:Records | Where-Object Status -eq 'WouldRemove')
$skipped     = @($Script:Records | Where-Object Status -eq 'Skipped')
$failed      = @($Script:Records | Where-Object Status -eq 'Failed')

$reclaimedKB = ($removed | Measure-Object SizeKB -Sum).Sum
$projectedKB = ($wouldRemove | Measure-Object SizeKB -Sum).Sum

$summary = "Run finished: removed=$($removed.Count) reclaimed=$(Format-Size $reclaimedKB)"
if ($wouldRemove.Count) { $summary += " would-remove=$($wouldRemove.Count) projected=$(Format-Size $projectedKB)" }
$summary += " skipped=$($skipped.Count) failed=$($failed.Count)"

$finalSeverity = if ($failed.Count) { 'Warning' } else { 'Info' }
Write-Verbose $summary
Write-CMTraceLog -Message "=== $summary ===" -Severity $finalSeverity
