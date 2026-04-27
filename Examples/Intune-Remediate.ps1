# ABOUTME: Intune Proactive Remediation - remediation script. Pair with Intune-Detect.ps1.
# ABOUTME: Inlines the Clear-CCMCache.ps1 invocation parameters because Intune deploys each
# ABOUTME: script standalone; copy the canonical Clear-CCMCache.ps1 content below the marker
# ABOUTME: before deploying, OR fetch it at runtime from a trusted location.

# Edit these to match your cleanup policy.
$Days        = 30
$MaxSizeMB   = 8192

# Option A: fetch the canonical script at runtime (requires HTTPS reachability from endpoints).
# Adjust the URL to your release tag.
$ScriptUrl  = 'https://raw.githubusercontent.com/rmdes/Clear-CCMCache.ps1/main/Clear-CCMCache.ps1'
$ScriptPath = Join-Path $env:TEMP 'Clear-CCMCache.ps1'

try {
    Invoke-WebRequest -Uri $ScriptUrl -OutFile $ScriptPath -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Output "Failed to download cleanup script: $($_.Exception.Message)"
    exit 1
}

# Option B (alternative): replace the download block above with a here-string containing the
# canonical script, write it to $ScriptPath, and run from there. Useful for air-gapped fleets.

& $ScriptPath -Days $Days -MaxSizeMB $MaxSizeMB -Verbose
exit $LASTEXITCODE
