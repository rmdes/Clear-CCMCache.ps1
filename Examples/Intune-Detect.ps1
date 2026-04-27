# ABOUTME: Intune Proactive Remediation - detection script. Pair with Intune-Remediate.ps1.
# ABOUTME: Exits 1 (needs remediation) when CCM cache exceeds the threshold; 0 otherwise.

# Edit these to match your fleet policy.
$ThresholdMB = 8192          # remediate when cache exceeds this many MB
$RequireCcmClient = $true    # if $true, machines without CCM are reported compliant (exit 0)

try {
    if ($RequireCcmClient) {
        $svc = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -ne 'Running') {
            Write-Output 'CCM client not running - skipping (compliant by policy).'
            exit 0
        }
    }

    $config = Get-CimInstance -Namespace ROOT\ccm\SoftMgmtAgent -ClassName CacheConfig `
        -Filter "ConfigKey='Cache'" -ErrorAction Stop
    $path = $config.Location
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        Write-Output 'Cache path missing - compliant.'
        exit 0
    }

    $bytes = (Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue |
        Measure-Object Length -Sum).Sum
    $sizeMB = if ($bytes) { [int]($bytes / 1MB) } else { 0 }

    if ($sizeMB -gt $ThresholdMB) {
        Write-Output "Cache size $sizeMB MB exceeds threshold $ThresholdMB MB - remediation needed."
        exit 1
    }

    Write-Output "Cache size $sizeMB MB <= threshold $ThresholdMB MB - compliant."
    exit 0
} catch {
    # Fail-safe: detection errors should NOT trigger remediation. Report compliant.
    Write-Output "Detection error: $($_.Exception.Message)"
    exit 0
}
