# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A single PowerShell script — `Clear-CCMCache.ps1` — that cleans the SCCM/CCM client cache on Windows endpoints. There is no build system, no test framework, no CI pipeline, and no other source files. The whole codebase is the script and the README.

This is a personal fork of `VirtualOx/Clear-CCMCache.ps1`. Backward compatibility with upstream is not a goal; breaking changes are fine.

## Validating changes

There are no unit tests. Use these in order before suggesting a change is "done":

1. **Parse check** (works on any machine, no admin, no CCM client needed):
   ```powershell
   $errors = $null; $tokens = $null
   [void][System.Management.Automation.Language.Parser]::ParseFile(
     '.\Clear-CCMCache.ps1', [ref]$tokens, [ref]$errors)
   if ($errors) { $errors } else { 'OK' }
   ```
   The script must parse cleanly under both Windows PowerShell 5.1 and PowerShell 7+.

2. **Lint** (optional, only if PSScriptAnalyzer is installed):
   ```powershell
   Invoke-ScriptAnalyzer -Path .\Clear-CCMCache.ps1
   ```

3. **End-to-end run** requires a CCM-managed Windows endpoint and an elevated shell. Always preview first:
   ```powershell
   .\Clear-CCMCache.ps1 -WhatIf -Verbose
   .\Clear-CCMCache.ps1 -Verbose       # only after preview looks right
   ```
   On a non-managed dev box the script will fail fast at the first `Get-CimInstance` because `ROOT\ccm\SoftMgmtAgent` doesn't exist — that is the correct failure path, not a bug.

## Architecture (script flow)

The script runs three passes against the CCM client cache. Reading the script top-to-bottom mirrors execution order.

1. **Resolve the cache root** from `ROOT\ccm\SoftMgmtAgent\CacheConfig` (CIM). The disk path is never hardcoded.
2. **Stale-entry pass**: enumerate `CacheInfoEx`, filter where `LastReferenced` is older than `-Days`, then for each entry delete the on-disk folder *and* the CIM record. Disk delete and CIM delete are independent try/catch blocks so a locked folder doesn't block index cleanup.
3. **Orphan reconciliation** (re-queries CIM after pass 2 to get fresh state):
   - Disk folders with no matching CIM record → deleted from disk.
   - CIM records pointing at folders that no longer exist → deleted from CIM.

## Non-obvious constraints

These are easy to break by accident — keep them in mind when editing:

- **CIM, not WMI.** Never reintroduce `Get-WmiObject` / `Remove-WmiObject`; they are removed in PS 7. The script must run unchanged on PS 5.1 and PS 7+, which is why CIM cmdlets are used and why `#Requires -Version 5.1` (not 7) is set.
- **Case-insensitive path comparison is load-bearing.** CIM returns the cache path in different casing than `Get-ChildItem` does (`C:\windows\ccmcache` vs `C:\Windows\ccmcache` on the same machine). The orphan diff uses `[StringComparer]::OrdinalIgnoreCase` `HashSet[string]` for this reason. Replacing it with `-notin` or a default-comparer hashset will cause every disk folder to look like a false orphan and get deleted.
- **`Test-PathUnder` guards every `Remove-Item`.** The script runs as administrator with `-Recurse -Force`. Any new deletion path must verify the target is under the resolved `$CachePath` before calling `Remove-Item`. Don't shortcut this.
- **Destructive operations go through `$PSCmdlet.ShouldProcess(...)`.** This is what powers `-WhatIf` / `-Confirm`. New destructive steps must follow the same pattern, not bypass it.
- **`ABOUTME:` header.** The two `# ABOUTME:` lines at the top of the script are per the maintainer's global convention. Keep them; if the script's purpose changes, update them.

## Git

Repo-local git identity is `Ricardo Mendes <rick@rmendes.net>` (set in `.git/config`, not global). Don't change it, and don't fall back to a global identity for commits in this repo.
