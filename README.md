# LiteTask

[![License: Unlicense](https://img.shields.io/badge/license-Unlicense-green.svg)](LICENSE)
[![GitHub release](https://img.shields.io/github/v/release/svtica/LiteTask)](https://github.com/svtica/LiteTask/releases/latest)
[![Release](https://github.com/svtica/LiteTask/actions/workflows/Release.yml/badge.svg)](https://github.com/svtica/LiteTask/actions/workflows/Release.yml)

**Lightweight alternative to Windows Task Scheduler with advanced PowerShell integration and secure credential management.**

LiteTask is a comprehensive task scheduling application that can run as both a desktop application and Windows service, providing flexible automation capabilities with built-in tool management.

## Features

### Core Scheduling
- **Multiple Task Types**: PowerShell scripts, batch files, SQL scripts, and executables
- **Advanced Scheduling**: One-time, interval-based, daily with multiple time slots, monthly, and task chaining
- **Action System**: Multiple actions per task, with dependencies between them
- **Secure Execution**: Credential management, Windows/SQL authentication, and elevation control

### PowerShell Integration
- Built-in module management with isolated environments
- Support for popular modules (Azure, AWS, VMware, etc.)
- PowerShell remoting with CredSSP support
- Remote execution capabilities

### Tool Management
- Optional tools (sqlcmd, LitePM, LiteDeploy) downloaded on demand from their GitHub releases (Tools > Manage Tools)
- sqlcmd downloaded automatically the first time a SQL task needs it
- Built-in process monitoring through LitePM

### Notifications & Logging
- Smart notification batching with email alerts
- Comprehensive logging with rotation policies
- Error tracking and diagnostic capabilities

## Installation

0. **Prerequisite**: [.NET 10 Desktop Runtime (x64)](https://dotnet.microsoft.com/download/dotnet/10.0) (`Microsoft.WindowsDesktop.App` 10.x)
1. Download the latest release
2. Extract files to desired location
3. After a manual extraction, run `LiteTask-Post-Build.ps1` once from the LiteTask folder: it removes the Mark of the
   Web from the extracted files and deletes the build folders LiteTask does not use (the built-in updater runs it
   automatically)
4. Run `LiteTask.exe`
5. **Optional - Install as Windows Service**:
   ```cmd
   LiteTask.exe -register
   ```
   *Requires administrative privileges and "Log on as a service" rights*

## Usage

### Desktop Application
- Launch `LiteTask.exe` for the graphical interface
- Configure tasks, schedules, and credentials through the UI
- Monitor execution status and logs

### Command Line Interface
```cmd
LiteTask.exe [options]
  -service          Run as Windows service
  -register         Register Windows service (requires admin)
  -unregister       Unregister Windows service (requires admin)
  -runtask <name>   Run the specified task
  -debug            Run in debug mode
```

### Configuration
1. **First Launch**: Application creates necessary directories and default configuration
2. **Tool Setup**: Go to Tools > Check Tools to verify utilities
3. **PowerShell Modules**: Tools > Install PowerShell Modules for additional functionality
4. **Credentials**: File > Credential Manager to add authentication details (account type *LiteSuite*: shared with LiteDeploy)

## Directory Structure

```
LiteTaskData/
├── logs/              # Application and task logs
├── Modules/           # PowerShell modules  
├── tools/             # Downloaded tools (sqlcmd.exe, LitePM\, LiteDeploy\)
├── temp/              # Temporary files
└── settings.xml       # Configuration file
```

## Technology Stack

- **Platform**: .NET 10 (LTS) Windows, PowerShell SDK 7.6 (LTS)
- **Language**: VB.NET
- **UI**: Windows Forms
- **Dependencies**: PowerShell SDK, ServiceController, [LiteSuite.Common](https://github.com/svtica/LiteSuite.Common)
  (NuGet package from GitHub Packages, shared with LiteDeploy: set the user environment variable
  `LITESUITE_PACKAGES_TOKEN` to a personal access token with `read:packages` only to build locally)

## Releases

*Actions > Release > Run workflow* (`patch` by default): `bump-version.ps1` raises the version
(patch = last part, e.g. 1.3.0.3 -> 1.3.0.4), commits and tags it on `main`, then the build signs
and publishes `LiteTask.zip`. Same use as LitePM and LiteDeploy. Locally:
`.\bump-version.ps1 -Bump patch` (or `-Version 1.3.0.4`).

## License

This software is released under [The Unlicense](LICENSE) - public domain.

---

For detailed help and troubleshooting, see [help.md](help.md).

## Part of LiteSuite

This tool is part of **LiteSuite** - a comprehensive collection of lightweight Windows administration tools.

### Other Tools in the Suite:
- **[LiteTask](https://github.com/svtica/LiteTask)** - Advanced Task Scheduler Alternative
- **[LitePM](https://github.com/svtica/LitePM)** - Process Manager with System Monitoring
- **[LiteDeploy](https://github.com/svtica/LiteDeploy)** - Network Deployment and Management
- **[LiteRun](https://github.com/svtica/LiteRun)** - Remote Command Execution Utility (archived)
- **[LiteSrv](https://github.com/svtica/LiteSrv)** - Windows Service Wrapper (archived)
