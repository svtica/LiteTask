#Requires -Version 5.1
<#
.SYNOPSIS
    Installs PowerShell modules into LiteTask's private module folder.

.DESCRIPTION
    Downloads the requested modules from the PowerShell Gallery with Save-Module into
    LiteTaskData\Modules (the folder LiteTask adds to PSModulePath for its tasks),
    then removes the Mark of the Web from the saved files. Called by
    Tools > Install PowerShell Modules (ModulesForm) and usable on its own.

.PARAMETER Modules
    Module names. Accepts an array or a single comma-separated string ("Az,ImportExcel"),
    which is how LiteTask passes them. Defaults to $DefaultModules below.

.PARAMETER Destination
    Folder the modules are saved to. Defaults to LiteTaskData\Modules next to this script.

.PARAMETER LogPath
    Log file. Defaults to LiteTaskData\logs\ModuleInstall_<timestamp>.log.

.PARAMETER Force
    Saves the module again even if it is already present in Destination (update).

.PARAMETER ShowProgress
    Shows a progress bar.

.EXAMPLE
    .\InstallModules.ps1 -Modules "Az.Accounts,ImportExcel" -ShowProgress

.NOTES
    File Name : InstallModules.ps1
    Author    : SVTI
    Version   : 3.0 (2026-11-11)
#>
[CmdletBinding()]
param (
    [string[]]$Modules,
    [string]$Destination,
    [string]$LogPath,
    [switch]$Force,
    [switch]$ShowProgress
)

$ErrorActionPreference = 'Stop'

# Default modules when -Modules is not given. AzureAD and MSOnline were retired by
# Microsoft: use Microsoft.Graph instead.
$DefaultModules = @(
    'Az',                # Azure
    'Microsoft.Graph',   # Microsoft 365 / Entra ID
    'ImportExcel',       # Excel files without Excel
    'PSWindowsUpdate'    # Windows Update management
)

$dataRoot = $PSScriptRoot   # the script lives in LiteTaskData\
if (-not $Destination) { $Destination = Join-Path $dataRoot 'Modules' }
if (-not $LogPath) {
    $LogPath = Join-Path (Join-Path $dataRoot 'logs') ("ModuleInstall_{0:yyyyMMddHHmmss}.log" -f (Get-Date))
}

# LiteTask passes one comma-separated string; split it into names
if (-not $Modules -or $Modules.Count -eq 0) { $Modules = $DefaultModules }
$Modules = @($Modules | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

function Write-LogMessage {
    param([string]$Message, [ValidateSet('INFO', 'WARNING', 'ERROR')][string]$Level = 'INFO')
    $line = "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f (Get-Date), $Level, $Message
    switch ($Level) {
        'ERROR'   { [Console]::Error.WriteLine($Message) }
        'WARNING' { Write-Warning $Message }
        default   { Write-Host $Message }
    }
    try {
        $dir = Split-Path $LogPath -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -Path $LogPath -Value $line
    } catch { }
}

function Initialize-Gallery {
    # TLS 1.2 for Windows PowerShell 5.1
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    if ($PSVersionTable.PSVersion.Major -lt 6) {
        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if (-not $nuget -or $nuget.Version -lt [Version]'2.8.5.201') {
            Write-LogMessage 'Installing the NuGet package provider...'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
    }
    if (-not (Get-Command Save-Module -ErrorAction SilentlyContinue)) {
        throw 'Save-Module (PowerShellGet) is not available.'
    }
}

function Test-ModuleInDestination([string]$Name) {
    # Only what is already in LiteTask's folder counts: a copy elsewhere on the machine
    # is not visible to LiteTask's tasks.
    return Test-Path (Join-Path $Destination $Name)
}

function Save-ModuleWithRetry([string]$Name, [int]$MaxRetries = 3) {
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            if (-not $Force -and (Test-ModuleInDestination $Name)) {
                Write-LogMessage "$Name is already present in $Destination (use -Force to update)."
                return $true
            }
            Write-LogMessage "Downloading $Name (attempt $attempt of $MaxRetries)..."
            Save-Module -Name $Name -Path $Destination -Repository PSGallery -Force
            if (-not (Test-ModuleInDestination $Name)) { throw "$Name was not found in $Destination after Save-Module." }
            Write-LogMessage "$Name saved."
            return $true
        } catch {
            Write-LogMessage "Attempt $attempt for $Name failed: $($_.Exception.Message)" -Level WARNING
            if ($attempt -lt $MaxRetries) { Start-Sleep -Seconds ([Math]::Pow(2, $attempt)) }
        }
    }
    Write-LogMessage "Failed to install $Name after $MaxRetries attempts." -Level ERROR
    return $false
}

$failed = @()
try {
    Write-LogMessage "Installing modules into $Destination : $($Modules -join ', ')"
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Initialize-Gallery

    for ($i = 0; $i -lt $Modules.Count; $i++) {
        $name = $Modules[$i]
        if ($ShowProgress) {
            Write-Progress -Activity 'Installing PowerShell modules' -Status $name -PercentComplete ($i / $Modules.Count * 100)
        }
        if (-not (Save-ModuleWithRetry $name)) { $failed += $name }
    }
    if ($ShowProgress) { Write-Progress -Activity 'Installing PowerShell modules' -Completed }

    # Remove the Mark of the Web from everything saved, so the modules load under any
    # execution policy (RemoteSigned refuses unsigned scripts that carry it).
    $unblocked = 0
    Get-ChildItem -Path $Destination -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        if (Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
            Unblock-File -LiteralPath $_.FullName
            $unblocked++
        }
    }
    if ($unblocked -gt 0) { Write-LogMessage "Removed the Mark of the Web from $unblocked file(s)." }
} catch {
    Write-LogMessage "Module installation failed: $($_.Exception.Message)" -Level ERROR
    Write-LogMessage "Log file: $LogPath"
    exit 1
}

if ($failed.Count -gt 0) {
    Write-LogMessage "Failed: $($failed -join ', ')" -Level WARNING
    Write-LogMessage "Log file: $LogPath"
    exit 1
}
Write-LogMessage 'All modules installed successfully.'
Write-LogMessage "Log file: $LogPath"
exit 0
