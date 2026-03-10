# LiteTask

[![GitHub release](https://img.shields.io/github/v/release/svtica/LiteTask)](https://github.com/svtica/LiteTask/releases/latest)
[![Release](https://github.com/svtica/LiteTask/actions/workflows/Release.yml/badge.svg)](https://github.com/svtica/LiteTask/actions/workflows/Release.yml)
[![License: Unlicense](https://img.shields.io/badge/license-Unlicense-green.svg)](LICENSE)

**Lightweight alternative to Windows Task Scheduler with advanced PowerShell integration and secure credential management.**

LiteTask is a comprehensive task scheduling application that can run as both a desktop application and Windows service, providing flexible automation capabilities with built-in tool management.

## Features

### Core Scheduling
- **Multiple Task Types**: PowerShell scripts, batch files, SQL scripts, and executables
- **Advanced Scheduling**: One-time, interval-based, daily with multiple time slots, monthly, and task chaining
- **Action System**: Multiple actions per task with dependencies, retry mechanisms, and parallel execution
- **Secure Execution**: Credential management, Windows/SQL authentication, and elevation control

### PowerShell Integration
- Built-in module management with isolated environments
- Support for popular modules (Azure, AWS, VMware, etc.)
- PowerShell remoting with CredSSP support
- Remote execution capabilities

### Tool Management
- Automatic detection and updates of integrated tools
- Built-in process monitoring and diagnostic utilities
- Integrated LiteDeploy and LitePM tools

### Notifications & Logging
- Smart notification batching with email alerts
- Comprehensive logging with rotation policies
- Error tracking and diagnostic capabilities

## Installation

1. Download the latest release
2. Extract files to desired location
3. On first extraction, run LiteTask-Post-Build.ps1 to remove Mark of the Web on files and remove useless files that were gererated automatically by the compiler
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
  -runtask <n>   Run specified task
  -debug            Run in debug mode
```

### Configuration
1. **First Launch**: Application creates necessary directories and default configuration
2. **Tool Setup**: Go to Tools > Check Tools to verify utilities
3. **PowerShell Modules**: Tools > Install PowerShell Modules for additional functionality
4. **Credentials**: File > Credential Manager to add authentication details

## Directory Structure

```
LiteTaskData/
├── logs/              # Application and task logs
├── Modules/           # PowerShell modules  
├── tools/             # Integrated utilities (LiteDeploy, LitePM)
├── temp/              # Temporary files
└── settings.xml       # Configuration file
```

## Technology Stack

- **Platform**: .NET 8.0 Windows
- **Language**: VB.NET
- **UI**: Windows Forms
- **Dependencies**: PowerShell SDK, ServiceController

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
