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

# Refuse to tag a dirty tree or a branch other than main: the tag would publish a
# release that does not match what is on GitHub.
if (git status --porcelain) { throw 'Working tree has uncommitted changes; commit or stash them first.' }
$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne 'main') { throw "Releases are tagged from main (current branch: $branch)." }
git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) { throw 'git pull --ff-only failed; resolve it before bumping the version.' }

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
