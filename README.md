# LiteTask

LiteTask is a lightweight task scheduling application for Windows that provides a flexible and secure alternative to Windows Task Scheduler.

## Features

- **Multiple Task Types**
  - PowerShell scripts with module support
  - Batch files
  - SQL scripts
  - Executable files

- **Action System**
  - Multiple actions per task
  - Action dependencies and ordering
  - Retry mechanisms with configurable delays
  - Error handling and continuation options
  - Parallel execution support
  - Action-level elevation control

- **Advanced Scheduling**
  - One-time execution
  - Interval-based scheduling
  - Daily scheduling with multiple time slots
  - Task chaining

- **Security**
  - Secure credential management
  - Windows authentication
  - SQL authentication
  - Service account support
  - Elevation control
  - User-context service execution

- **PowerShell Integration**
  - Built-in module management
  - Isolated module environment
  - Support for popular modules (Azure, AWS, VMware, etc.)
  - Remote execution support

- **Remote Execution**
  - PowerShell remoting
  - CredSSP support
  - SQL remote execution

- **Tool Management**
  - Automatic tool detection and updates
  - Built-in updating system
  - Process monitoring
  - Diagnostic utilities

- **Notification System**
  - Smart notification batching
  - Email notifications
  - Configurable alert levels
  - Error tracking

## Installation

1. Download the latest release from the releases page
2. Extract the files to your desired location
3. Run LiteTask.exe
4. Optional: Install as a Windows service using administrative privileges:
   ```
   LiteTask.exe -register
   ```
   - You will be prompted for your Windows password
   - The service will run under your user account
   - Your account must have Log on as a service rights

## Configuration

1. **First Launch**
   - The application will create necessary directories
   - Initial configuration will be generated
   - Default logging will be enabled with rotation policies

2. **Tool Setup**
   - Go to Tools > Check Tools to verify required utilities
   - Tools are automatically updated when newer versions are available
   - Missing components are downloaded automatically

3. **PowerShell Module Setup**
   - Go to Tools > Install PowerShell Modules
   - Select desired modules from the list
   - Modules are installed to LiteTaskData\Modules

4. **Credential Management**
   - Open File > Credential Manager
   - Add necessary credentials for task execution
   - Supports Windows and SQL authentication

## Command Line Interface

```
LiteTask.exe [options]
  -service           Run as Windows service
  -register          Register Windows service (requires admin, runs as current user)
  -unregister        Unregister Windows service (requires admin)
  -start            Start service
  -stop             Stop service
  -runtask <name>   Run specified task
  -debug            Run in debug mode
```

## Directory Structure

```
LiteTaskData/
├── logs/           # Application, task, and action logs
│   ├── app_log.txt    # Main application log
│   └── *.log.gz       # Rotated and compressed logs
├── Modules/        # PowerShell modules
├── temp/           # Temporary files
├── tools/          # Required utilities
└── settings.xml    # Application configuration
```

## Troubleshooting

1. Check app_log.txt for error messages
2. Verify service account permissions
3. Check module installation logs
4. Verify file permissions
5. Check network connectivity for remote tasks
6. Verify credential configuration
7. Check action-specific logs
8. Verify action dependencies
9. Monitor action execution order
10. Review service account permissions

For detailed help, please check the [help](help.md) file.

## Support

- Check documentation for common solutions

## License

This software is provided under [Unlicence](LICENSE) terms.
