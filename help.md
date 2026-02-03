# LiteTask Help

This document provides detailed information on how to use LiteTask, a lightweight task scheduling application for Windows.

## Table of Contents

1. [Task Types](#task-types)
2. [Creating Tasks](#creating-tasks)
3. [Task Chaining](#task-chaining)
4. [Scheduling Options](#scheduling-options)
5. [Monthly Scheduling](#monthly-scheduling)
6. [Credential Management](#credential-management)
7. [Remote Execution](#remote-execution)
8. [Command-line Interface](#command-line-interface)
9. [Tool Management](#tool-management)
10. [PowerShell Module Management](#powershell-module-management)
11. [Troubleshooting](#troubleshooting)

## Task Types

LiteTask supports the following task types:

1. **PowerShell**: Execute PowerShell scripts with support for parameter filtering and CredSSP authentication.
2. **Batch**: Run batch files or command scripts.
3. **SQL**: Execute SQL scripts using OSQL with support for Windows and SQL authentication.
4. **Remote Execution**: Run tasks on remote machines using various methods.
5. **Executable**: Run any executable file with specified arguments.

## Scheduling Options

LiteTask supports the following scheduling types:

- **One-time**: Run the task once at a specified date and time
- **Interval**: Repeat the task at a fixed interval (in minutes) from a start time
- **Daily**: Run the task at one or more specific times every day
- **Monthly**: Run the task on a specific day of the month at a given time

Select "Recurring" in the Schedule tab to enable interval, daily, or monthly scheduling.

## Monthly Scheduling

Monthly scheduling allows you to run a task on a specific day each month at a chosen time.

### Configuration
1. Open the Create or Edit Task form
2. Go to the **Schedule** tab
3. Check **Recurring**
4. Select **Monthly** from the recurrence type dropdown
5. Set the **Day of month** (1-31) and the execution **Time**

### Day-of-month handling
If the selected day exceeds the number of days in a given month (e.g., day 31 in February), the task will run on the last day of that month instead. This ensures the task always executes once per month.

## Service Installation

### Requirements
- Administrative privileges for installation
- User account must have "Log on as a service" rights
- SeImpersonatePrivilege for certain operations

### Installation Steps
1. Open command prompt as Administrator
2. Navigate to LiteTask directory
3. Run `LiteTask.exe -register`
4. Enter your Windows password when prompted
5. Service runs under your user account

### Service Operation
- Start: `LiteTask.exe -start`
- Stop: `LiteTask.exe -stop`
- Uninstall: `LiteTask.exe -unregister`

## Actions System

LiteTask uses a powerful Action system that allows complex task execution flows:

### Action Properties
- **Name**: Unique identifier for the action
- **Type**: The type of action (PowerShell, Batch, SQL, etc.)
- **Target**: Path or target of the action
- **Parameters**: Additional parameters for execution
- **Requires Elevation**: Whether the action needs administrative privileges

### Action Dependencies
Actions can be configured with dependencies:
- **Depends On**: Specify other tasks that must complete before this action
- **Wait for Completion**: Whether to wait for dependent tasks to finish
- **Timeout**: Maximum time to wait for completion (in minutes)
- **Retry Count**: Number of retry attempts if action fails
- **Retry Delay**: Time to wait between retries (in minutes)
- **Continue on Error**: Whether to continue with next action if this one fails

### Log Management
- Automatic log rotation based on size
- Compression of old logs
- Configurable retention period
- Structured logging with levels

### Notification System
- Smart batching of related notifications
- Configurable thresholds
- Email notifications
- Priority levels
- Error tracking and reporting

## Tool Management

LiteTask includes built-in tool management:

1. Access the Tools tab in the ribbon menu
2. Tools are automatically checked and updated
3. Missing tools are downloaded automatically
4. Process monitoring available through LitePM

Available tools:
- Process Explorer (procexp.exe)
- LiteRun
- OSQL
- PsExec

## PowerShell Module Management

LiteTask now includes PowerShell module management capabilities:

1. Access Tools > Install PowerShell Modules
2. Choose from available modules:
   - Az: Azure PowerShell
   - AzureAD
   - MSOnline
   - PSWindowsUpdate
   - PSBlueTeam
   - Pester
   - ImportExcel
   - VMware.PowerCLI
   - SqlServer
   - AWS.Tools.Common

## Troubleshooting

Common issues and solutions:

### Service Issues
1. Verify user account has "Log on as a service" rights
2. Check Windows Event Log for service errors
3. Verify SeImpersonatePrivilege is granted
4. Check service account permissions

### Task Execution Issues
1. Review task-specific logs
2. Check credential configuration
3. Verify file permissions
4. Monitor action dependencies
5. Review timeout settings

### Remote Execution Issues
1. Check network connectivity
2. Verify remote credentials
3. Test remote PowerShell access
4. Confirm firewall settings

### Logging Issues
1. Check disk space for logs
2. Verify log rotation settings
3. Review compression settings
4. Monitor log retention
