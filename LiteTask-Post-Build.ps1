<#
.SYNOPSIS
    Post-build cleanup script for LiteTask application directories and files
.DESCRIPTION
    This PowerShell script manages the cleanup of LiteTask application folders,
    keeping only essential language and runtime directories and unblock files.
.NOTES
    File Name      : LiteTask-Post-Build.ps1
    Author         : SVTI.ca
    Prerequisite   : PowerShell 5.0 or later
    Created Date   : 2025-01-16
.EXAMPLE
    .\LiteTask-Post-Build.ps1
#>

Write-Host "Please wait, while we perform the cleanup..."

# Remove all folders except the specified ones in root directory
Get-ChildItem -Directory | Where-Object { $_.Name -notin @("runtimes", "ref", "res", "lib", "fr", "LiteTaskData") } | Remove-Item -Recurse -Force

# Remove unused runtimes folders
Get-ChildItem "runtimes" -Directory | Where-Object { !($_.Name -like "win*") } | Remove-Item -Recurse -Force

Write-Host "Cleanup completed."

Write-Host "Please wait, while we unblock the files..."

# Define the path to the folder you want to process
# $folderPath = "C:\Users\yourusername\Downloads"
# $folderPath = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path

# Extracted zip content folder
$folderPath = ".\"

# Define file extensions that need to be unblocked (customize as needed)
$includedExtensions = @(".exe", ".dll", ".msi", ".zip", ".ps1", ".bat")

# Get all files in the folder and subfolders with the specified extensions
$files = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object { $includedExtensions -contains $_.Extension }

# Unblock each file
foreach ($file in $files) {
    try {
        Unblock-File -Path $file.FullName -ErrorAction Stop
        Write-Host "Unblocked: $($file.FullName)"
    }
    catch {
        Write-Host "Failed to unblock: $($file.FullName). Error: $($_.Exception.Message)"
    }
}

Write-Host "Unblocking completed."

