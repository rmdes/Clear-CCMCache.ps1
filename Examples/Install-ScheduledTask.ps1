# ABOUTME: Registers a weekly Windows scheduled task that runs Clear-CCMCache.ps1 as SYSTEM.
# ABOUTME: Run from an elevated PowerShell. Default: every Sunday at 02:00 local time.

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [string]$ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '..\Clear-CCMCache.ps1')).Path,
    [int]$Days       = 30,
    [int]$MaxSizeMB  = 0,
    [string]$TaskName = 'Clear-CCMCache',
    [DayOfWeek]$DayOfWeek = 'Sunday',
    [DateTime]$At = '02:00'
)

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Clear-CCMCache.ps1 not found at $ScriptPath"
}

$argList = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -Days $Days"
if ($MaxSizeMB -gt 0) { $argList += " -MaxSizeMB $MaxSizeMB" }

$action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argList
$trigger   = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $At
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings `
    -Description "Weekly CCM cache cleanup (Days=$Days, MaxSizeMB=$MaxSizeMB)" -Force | Out-Null

Write-Host "Registered scheduled task '$TaskName' running $DayOfWeek at $($At.ToString('HH:mm'))."
Write-Host "Inspect with:  Get-ScheduledTask -TaskName '$TaskName'"
Write-Host "Run on demand: Start-ScheduledTask -TaskName '$TaskName'"
