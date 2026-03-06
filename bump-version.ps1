<#
.SYNOPSIS
    Bumps LiteTask version, commits, tags, and pushes to trigger a GitHub release.

.PARAMETER Version
    The new version number (e.g., "1.0.0.9")

.EXAMPLE
    .\bump-version.ps1 -Version "1.0.0.9"
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'

$projectFile = Join-Path $PSScriptRoot 'LiteTask.vbproj'
$manifestFile = Join-Path $PSScriptRoot 'app.manifest'

# Update LiteTask.vbproj
$proj = Get-Content $projectFile -Raw
$proj = $proj -replace '<FileVersion>[^<]+</FileVersion>', "<FileVersion>$Version</FileVersion>"
$proj = $proj -replace '<AssemblyVersion>[^<]+</AssemblyVersion>', "<AssemblyVersion>$Version</AssemblyVersion>"
$proj = $proj -replace '<Version>[^<]+</Version>', "<Version>$Version</Version>"
Set-Content $projectFile $proj -NoNewline

# Update app.manifest
$manifest = Get-Content $manifestFile -Raw
$manifest = $manifest -replace '(assemblyIdentity\s+version=")[^"]+(")', "`${1}$Version`${2}"
Set-Content $manifestFile $manifest -NoNewline

Write-Host "Version updated to $Version in:"
Write-Host "  - LiteTask.vbproj (FileVersion, AssemblyVersion, Version)"
Write-Host "  - app.manifest (assemblyIdentity)"

# Git operations
git add $projectFile $manifestFile
git commit -m "Bump version to $Version"
git tag "v$Version"
git push origin HEAD
git push origin "v$Version"

Write-Host "`nDone! Tag v$Version pushed - GitHub release workflow will trigger automatically."
