<#
.SYNOPSIS
    Post-extraction cleanup for a LiteTask package.
.DESCRIPTION
    Run once after extracting a LiteTask release (the updater runs it automatically on
    the extracted package). Works on the folder that contains this script, never on the
    current directory:
      1. removes the folders the build leaves behind that LiteTask does not use
         (only folders, and never runtimes, ref, res, lib, fr, LiteTaskData);
      2. removes non-Windows runtime folders;
      3. removes the Mark of the Web (Zone.Identifier) from every file, so that files of a
         zip downloaded with a browser do not trigger security prompts or execution-policy
         blocks.
    Does nothing if LiteTask.exe is not next to the script.
.NOTES
    File Name    : LiteTask-Post-Build.ps1
    Author       : SVTI.ca
    Prerequisite : PowerShell 5.1 or later
    Updated      : 2026-11-11
.EXAMPLE
    .\LiteTask-Post-Build.ps1
#>

$root = $PSScriptRoot
if (-not $root -or -not (Test-Path (Join-Path $root 'LiteTask.exe'))) {
    Write-Host "LiteTask.exe not found next to this script ($root); nothing done."
    exit 1
}

Write-Host "Please wait, while we perform the cleanup in $root ..."

# 1. Unused folders (only in the package root)
$keep = @('runtimes', 'ref', 'res', 'lib', 'fr', 'LiteTaskData')
Get-ChildItem -LiteralPath $root -Directory |
    Where-Object { $_.Name -notin $keep } |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# 2. Non-Windows runtimes
$runtimes = Join-Path $root 'runtimes'
if (Test-Path $runtimes) {
    Get-ChildItem -LiteralPath $runtimes -Directory |
        Where-Object { $_.Name -notlike 'win*' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Cleanup completed."

# 3. Mark of the Web: only files that carry it (any type: exe, dll, ps1, config, json...)
Write-Host "Please wait, while we unblock the files..."
$unblocked = 0
$failed = 0
Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    if (Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
        try {
            Unblock-File -LiteralPath $_.FullName -ErrorAction Stop
            $unblocked++
        }
        catch {
            $failed++
            Write-Host "Failed to unblock: $($_.Exception.Message)"
        }
    }
}
Write-Host "Unblocking completed: $unblocked file(s) unblocked, $failed failure(s)."
