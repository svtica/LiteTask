#Requires -Version 5.0
<#
.SYNOPSIS
    Automated PowerShell module installation script with enhanced NuGet management.

.DESCRIPTION
    This script automates the installation of PowerShell modules with robust error handling,
    logging, and progress reporting. It includes automated NuGet provider installation and
    verification, and supports custom module lists and installation locations.

.PARAMETER Modules
    Array of module names to install. If not specified, uses the default list defined at the top of the script.

.PARAMETER LogPath
    Path to the log file. Defaults to %TEMP%\LiteTask_ModuleInstall.log

.PARAMETER Destination
    Installation path for modules. Defaults to .\LiteTaskData\Modules

.PARAMETER ShowProgress
    Switch to show installation progress bar

.EXAMPLE
    .\InstallModules.ps1
    Installs default modules with no progress bar

.EXAMPLE
    .\InstallModules.ps1 -ShowProgress
    Installs default modules with progress bar

.EXAMPLE
    .\InstallModules.ps1 -Modules @("Az", "AzureAD") -ShowProgress
    Installs specific modules with progress bar

.NOTES
    File Name      : InstallModules.ps1
    Author         : SVTI
    Prerequisite   : PowerShell 5.0 or higher
    Version        : 2.0
    Last Modified  : 2024-11-12
    
.LINK
    https://www.powershellgallery.com/
#>

# Default module configuration
# Edit this section to modify the default modules to install
$DEFAULT_MODULES = @(
    # Microsoft Azure Modules
    "Az",                  # Azure Resource Management
    "AzureAD",            # Azure Active Directory
    "MSOnline",           # Microsoft 365
    
    # System Management
    "PSWindowsUpdate",     # Windows Update Management
    
    # Security
    "PSBlueTeam",         # Security Best Practices
    
    # Add additional modules here
    # "ModuleName",       # Description
)

# NuGet configuration
$NUGET_CONFIG = @{
    MinimumVersion = "2.8.5.201"
    Repository = "PSGallery"
    InstallationPolicy = "Trusted"
}

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string[]]$Modules = $DEFAULT_MODULES,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = ".\LiteTaskData\temp\LiteTask_ModuleInstall.log",
    
    [Parameter(Mandatory=$false)]
    [string]$Destination = ".\LiteTaskData\Modules",
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowProgress
)

# Initialize logging with structured output
function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$NoOutput
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    
    if (-not $NoOutput) {
        switch ($Level) {
            "ERROR"   { Write-Error $Message }
            "WARNING" { Write-Warning $Message }
            default   { Write-Host $Message }
        }
    }
    
    try {
        if ($LogPath) {
            $logDir = Split-Path $LogPath -Parent
            if (-not (Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
        }
    }
    catch {
        $tempLogPath = ".\LiteTaskData\temp\LiteTask_ModuleInstall.log"
        Add-Content -Path $tempLogPath -Value $logMessage
        if ($LogPath -ne $tempLogPath) {
            Write-Warning "Failed to write to $LogPath, using $tempLogPath instead"
            $script:LogPath = $tempLogPath
        }
    }
}

# Enhanced NuGet verification and installation
function Initialize-NuGetProvider {
    try {
        Write-LogMessage "Verifying NuGet package provider..."
        
        # Check if NuGet is installed
        $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        
        $needsInstall = $false
        if (-not $nuget) {
            Write-LogMessage "NuGet provider not found. Installing..."
            $needsInstall = $true
        }
        elseif ($nuget.Version -lt [Version]$NUGET_CONFIG.MinimumVersion) {
            Write-LogMessage "NuGet provider needs update. Current: $($nuget.Version), Required: $($NUGET_CONFIG.MinimumVersion)"
            $needsInstall = $true
        }
        
        if ($needsInstall) {
            # Force TLS 1.2
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            # Install/Update NuGet
            Install-PackageProvider -Name NuGet -MinimumVersion $NUGET_CONFIG.MinimumVersion -Force | Out-Null
            
            # Verify installation
            $nuget = Get-PackageProvider -Name NuGet -ErrorAction Stop
            if ($nuget.Version -lt [Version]$NUGET_CONFIG.MinimumVersion) {
                throw "NuGet provider installation failed to meet minimum version requirement"
            }
        }
        
        # Configure PSGallery
        if ((Get-PSRepository -Name $NUGET_CONFIG.Repository).InstallationPolicy -ne $NUGET_CONFIG.InstallationPolicy) {
            Set-PSRepository -Name $NUGET_CONFIG.Repository -InstallationPolicy $NUGET_CONFIG.InstallationPolicy
            Write-LogMessage "Updated PSGallery installation policy to $($NUGET_CONFIG.InstallationPolicy)"
        }
        
        Write-LogMessage "NuGet provider verified (Version: $($nuget.Version))"
        return $true
    }
    catch {
        Write-LogMessage "NuGet initialization failed: $_" -Level "ERROR"
        return $false
    }
}

# Verify and prepare environment
function Initialize-Environment {
    try {
        # Resolve full paths
        $script:Destination = [System.IO.Path]::GetFullPath($Destination)
        $script:LogPath = [System.IO.Path]::GetFullPath($LogPath)
        
        # Create destination directory
        if (-not (Test-Path $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
            Write-LogMessage "Created destination directory: $Destination"
        }
        
        # Verify write permissions
        $testFile = Join-Path $Destination "test.txt"
        try {
            [System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item $testFile -Force
        }
        catch {
            throw "No write permission in destination directory: $Destination"
        }
        
        # Update PSModulePath
        $currentModulePath = $env:PSModulePath -split [System.IO.Path]::PathSeparator
        if ($currentModulePath -notcontains $Destination) {
            $env:PSModulePath = $Destination + [System.IO.Path]::PathSeparator + $env:PSModulePath
            Write-LogMessage "Added $Destination to PSModulePath"
        }
        
        return $true
    }
    catch {
        Write-LogMessage "Environment initialization failed: $_" -Level "ERROR"
        return $false
    }
}

# Verify and install prerequisites
function Install-Prerequisites {
    try {
        # Check PowerShell version
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "PowerShell 5.0 or higher is required"
        }
        
        # Initialize NuGet
        if (-not (Initialize-NuGetProvider)) {
            throw "NuGet provider initialization failed"
        }
        
        # Update PowerShellGet if needed
        $psGet = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ($psGet.Version -lt [Version]"2.0") {
            Write-LogMessage "Updating PowerShellGet..."
            Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser
            Write-LogMessage "PowerShellGet updated - please restart PowerShell and run this script again"
            return $false
        }
        
        return $true
    }
    catch {
        Write-LogMessage "Prerequisites installation failed: $_" -Level "ERROR"
        return $false
    }
}

# Enhanced module installation with progress and verification
function Install-ModuleWithRetry {
    param (
        [string]$ModuleName,
        [int]$MaxRetries = 3,
        [int]$CurrentModuleIndex,
        [int]$TotalModules
    )
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            if ($ShowProgress) {
                $progressParams = @{
                    Activity = "Installing PowerShell Modules"
                    Status = "Installing $ModuleName (Attempt $i of $MaxRetries)"
                    PercentComplete = ($CurrentModuleIndex / $TotalModules * 100)
                }
                Write-Progress @progressParams
            }
            
            # Check existing installation
            $existingModule = Get-Module -ListAvailable -Name $ModuleName
            if ($existingModule) {
                Write-LogMessage "Module $ModuleName is already installed (Version: $($existingModule.Version))"
                return $true
            }
            
            # Install module
            $installParams = @{
                Name = $ModuleName
                Force = $true
                AllowClobber = $true
                Scope = 'CurrentUser'
                ErrorAction = 'Stop'
                Destination = $Destination
            }
            
            Install-Module @installParams
            
            # Verify installation
            $verifyModule = Get-Module -ListAvailable -Name $ModuleName
            if (-not $verifyModule) {
                throw "Module installation verification failed"
            }
            
            Write-LogMessage "Successfully installed $ModuleName (Version: $($verifyModule.Version))"
            return $true
        }
        catch {
            Write-LogMessage "Attempt $i failed for $ModuleName`: $_" -Level "WARNING" -NoOutput
            if ($i -eq $MaxRetries) {
                Write-LogMessage "Failed to install $ModuleName after $MaxRetries attempts" -Level "ERROR"
                return $false
            }
            Start-Sleep -Seconds ([Math]::Pow(2, $i)) # Exponential backoff
        }
    }
    return $false
}

# Main execution block
try {
    Write-LogMessage "Starting PowerShell module installation process"
    
    # Initialize environment
    if (-not (Initialize-Environment)) {
        throw "Environment initialization failed"
    }
    
    # Install prerequisites
    if (-not (Install-Prerequisites)) {
        throw "Prerequisites installation failed"
    }
    
    # Install modules
    $successCount = 0
    $failureCount = 0
    $installedModules = @()
    $failedModules = @()
    
    for ($i = 0; $i -lt $Modules.Count; $i++) {
        $module = $Modules[$i]
        Write-LogMessage "Processing module: $module"
        
        if (Install-ModuleWithRetry -ModuleName $module -CurrentModuleIndex $i -TotalModules $Modules.Count) {
            $successCount++
            $installedModules += $module
        }
        else {
            $failureCount++
            $failedModules += $module
        }
    }
    
    # Clear progress bar
    if ($ShowProgress) {
        Write-Progress -Activity "Installing PowerShell Modules" -Completed
    }
    
    # Report results
    Write-LogMessage "Installation Summary:"
    Write-LogMessage "Successfully installed ($successCount): $($installedModules -join ', ')"
    if ($failedModules.Count -gt 0) {
        Write-LogMessage "Failed ($failureCount): $($failedModules -join ', ')" -Level "WARNING"
        exit 1
    }
    
    Write-LogMessage "All modules installed successfully."
    exit 0
}
catch {
    Write-LogMessage "Critical error during execution: $_" -Level "ERROR"
    exit 1
}
finally {
    Write-LogMessage "Log file location: $LogPath"
}