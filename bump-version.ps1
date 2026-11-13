<#
.SYNOPSIS
    Bumps LiteTask version, commits, tags and pushes to publish a GitHub release.

.DESCRIPTION
    Updates LiteTask.vbproj (FileVersion, AssemblyVersion, Version) and
    app.manifest, commits "Bump version to X" on main, tags v<X> and pushes both.

    Locally, pushing the tag triggers the "Build and Release" workflow. The
    "Release" workflow (Actions tab > Release > Run workflow) runs this script
    with -Bump or -Version and then builds the release itself, because a tag
    pushed from a workflow does not trigger other workflows.

.PARAMETER Version
    The new version, 4 parts (e.g. "1.3.0.4").

.PARAMETER Bump
    Instead of -Version: increment the current version, LiteTask style
    (patch = last part, 1.3.0.3 -> 1.3.0.4; minor = 1.3.0.3 -> 1.4.0.0;
    major -> 2.0.0.0).

.EXAMPLE
    .\bump-version.ps1 -Bump patch
.EXAMPLE
    .\bump-version.ps1 -Version "1.3.0.4"
#>
[CmdletBinding(DefaultParameterSetName = 'Explicit')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Explicit')]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$Version,

    [Parameter(Mandatory = $true, ParameterSetName = 'Bump')]
    [ValidateSet('patch', 'minor', 'major')]
    [string]$Bump
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

# Current version (<Version> of the project)
$proj = Get-Content $projectFile -Raw
$m = [regex]::Match($proj, '<Version>(\d+)\.(\d+)\.(\d+)\.(\d+)</Version>')
if (-not $m.Success) { throw "<Version> not found in $projectFile" }
$current = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value).$($m.Groups[4].Value)"

if ($PSCmdlet.ParameterSetName -eq 'Bump') {
    $major = [int]$m.Groups[1].Value; $minor = [int]$m.Groups[2].Value
    $build = [int]$m.Groups[3].Value; $revision = [int]$m.Groups[4].Value
    switch ($Bump) {
        'major' { $major++; $minor = 0; $build = 0; $revision = 0 }
        'minor' { $minor++; $build = 0; $revision = 0 }
        'patch' { $revision++ }
    }
    $Version = "$major.$minor.$build.$revision"
}

$tag = "v$Version"
if ([version]$Version -le [version]$current) { throw "New version $Version must be greater than the current $current." }
git fetch --tags --quiet origin
if (git tag --list $tag) { throw "Tag $tag already exists." }

Write-Host "Bumping $current -> $Version (tag $tag)"

# Update LiteTask.vbproj
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
if ($LASTEXITCODE -ne 0) { throw 'git commit failed.' }
git tag $tag
git push origin HEAD
if ($LASTEXITCODE -ne 0) { throw 'git push of main failed (branch protection?).' }
git push origin $tag
if ($LASTEXITCODE -ne 0) { throw "git push of $tag failed." }

# Expose the result to a calling workflow
if ($env:GITHUB_OUTPUT) {
    Add-Content -Path $env:GITHUB_OUTPUT -Value "version=$Version"
    Add-Content -Path $env:GITHUB_OUTPUT -Value "tag=$tag"
}

Write-Host "`nDone! $tag pushed."
