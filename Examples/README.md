# Examples

Copy-paste deployment artifacts for `Clear-CCMCache.ps1`.

| File | Purpose |
| --- | --- |
| `Intune-Detect.ps1` | Intune Proactive Remediation detection script. Exit 1 = needs remediation. |
| `Intune-Remediate.ps1` | Intune remediation script. Downloads the canonical cleanup script and runs it. |
| `Install-ScheduledTask.ps1` | Registers a weekly Windows scheduled task running as SYSTEM. |

## Exit codes used by the canonical script

The artifacts here pass `$LASTEXITCODE` through unchanged. Code meanings:

| Code | Meaning |
| --- | --- |
| `0` | Clean run. No failures, target met (if `-MaxSizeMB` was set). |
| `1` | Preflight failed (no CCM client, can't bind COM, cache path missing). |
| `2` | Run completed but one or more removals failed. |
| `3` | `-MaxSizeMB` was set and the cache is still over target after the run. |

Use these codes in monitoring / alerting / conditional retry logic.

## Intune Proactive Remediation

1. Upload `Intune-Detect.ps1` and `Intune-Remediate.ps1` to Intune Admin Center > Reports > Endpoint analytics > Proactive remediations > Create script package.
2. Set the schedule (typically daily for detect; remediate runs only if detect returns 1).
3. Tune the `$ThresholdMB` constant in `Intune-Detect.ps1` to your fleet policy.
4. The remediation script downloads `Clear-CCMCache.ps1` from this repo's `main` branch by default. Pin to a release tag for production use, or replace the download with an inline copy of the script.

## SCCM "Run Script"

The "Run Script" feature in MECM expects a single self-contained `.ps1`. Either:

- Concatenate `Clear-CCMCache.ps1` into a single deployable file, or
- Deploy `Clear-CCMCache.ps1` as a package and call it from a wrapper "Run Script."

The exit codes above are surfaced as the script result in the MECM console.

## Scheduled task

```powershell
# Default: weekly Sunday 02:00, 30-day threshold
.\Examples\Install-ScheduledTask.ps1

# Tighter policy: tuesday 03:00, 14-day threshold, target 8 GB
.\Examples\Install-ScheduledTask.ps1 -DayOfWeek Tuesday -At '03:00' -Days 14 -MaxSizeMB 8192
```

The task runs as `SYSTEM` and inherits the script's logging behavior (writes to `%SystemRoot%\CCM\Logs\ClearCache.log` by default).
