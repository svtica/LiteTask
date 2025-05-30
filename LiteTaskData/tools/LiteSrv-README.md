# LiteSRV v1.01

LiteSRV is a Windows service wrapper utility that allows you to run any program as a Windows NT service, with environment customization and extensive process management capabilities.

## Important Warning

**Modify Windows NT service settings at your own peril.**  
Making incorrect changes can render Windows unusable.  
If you don't know what you are doing, don't mess with services!

## Quick Start Guide

### Basic Setup

1. Unpack the LiteSRV executables to a directory in your PATH:
   - litesrv.exe - Main executable
   - litesrv.dll - Core library 
   - logger.dll - Logger library

2. Create a configuration file (e.g., `C:\LITESRV\litesrv.ini`):
   ```ini
   [MY_SERVICE]
   startup=C:\MYPROG\MYPROG.EXE -a -b -c xxxx yyyy zzz
   ```
   Always use full pathnames in the configuration file.

3. Install your service:
   ```
   litesrv install MY_SERVICE -c C:\LITESRV\litesrv.ini
   ```

4. Start the service:
   ```
   net start MY_SERVICE
   ```

5. Use Control Panel | Services to set the appropriate startup mode if needed.

## Features

- Run any program as a Windows NT service
- Run programs in command-line mode
- Configure custom environment variables
- Map network and local drives
- Control process execution priority
- Configure automatic process restart
- Multiple shutdown methods (graceful command, window message, or force kill)
- Detailed logging capabilities
- Desktop interaction support
- Service installation and removal utilities

## Installation

Simply copy LiteSRV.exe to your desired location. No additional installation is required.

## Usage

### Basic Syntax

```
LiteSRV <mode> <name> [options] <command> [program_parameters...]
```

### Modes

- `cmd` - Run in command mode (creates a window)
- `svc` - Run as NT service
- `any` - Try service mode first, fallback to command mode
- `install` - Install as NT service
- `install_desktop` - Install as NT service with desktop interaction
- `remove` - Remove NT service

### Options

#### Environment Options
- `-e var=value` - Set environment variable
- `-p path` - Set PATH environment variable
- `-l libdir` - Set LIB environment variable
- `-s sybase` - Set SYBASE environment variable
- `-q sybase` - Set default PATH based on SYBASE value

#### Execution Options
- `-x priority` - Set execution priority (normal/high/real/idle)
- `-w` - Start in new window (command mode only)
- `-m` - Start minimized (command mode only)
- `-y sec` - Startup delay in seconds (service mode only)
- `-t sec` - Process status check interval (service mode only)

#### Debug Options
- `-d level` - Debug level (0/1/2)
- `-o target` - Debug output target (filename/stdout/LOG)
- `-c ctrlfile` - Load options from control file
- `-h` - Display help message

### Examples

1. Run as command in new window:
```
LiteSRV cmd MyApp -w c:\apps\myapp.exe -param1 -param2
```

2. Install as service:
```
LiteSRV install MyService -c c:\config\service.conf
```

3. Run as service with environment:
```
LiteSRV svc MyService -e APP_HOME=c:\apps -e LOG_DIR=c:\logs c:\apps\myapp.exe
```

### Control File Format

Control files use an INI-style format with sections and key=value pairs:

```ini
[ServiceName]
startup=c:\apps\myapp.exe
startup_dir=c:\apps
env=APP_HOME=c:\apps
env=LOG_DIR=c:\logs
debug=1
debug_out=c:\logs\service.log
auto_restart=yes
restart_interval=60
shutdown_method=command
shutdown=c:\apps\stop.cmd
```

#### Available Control File Options

- `startup` - Command to run
- `startup_dir` - Working directory
- `env` - Environment variables (multiple allowed)
- `debug` - Debug level (0/1/2)
- `debug_out` - Debug output location
- `wait` - Status check command
- `wait_time` - Status check interval (seconds)
- `auto_restart` - Enable automatic restart (yes/no)
- `restart_interval` - Seconds between restart attempts
- `shutdown_method` - Shutdown method (command/winmessage/kill)
- `shutdown` - Shutdown command (if method=command)
- `network_drive` - Map network drive (X=\\server\share)
- `local_drive` - Map local drive (X=c:\path)

### Parameter Substitution

LiteSRV supports two types of parameter substitution in commands:

1. Environment Variables: Use %VARIABLE% syntax
```
startup=%APP_HOME%\myapp.exe
```

2. Prompted Values: Use {prompt} or {prompt:default} syntax
```
startup=myapp.exe -u {username} -p {-password}
```
Note: Adding `-` before the prompt name hides the input (for passwords)

### Shutdown Methods

LiteSRV supports three ways to stop a service:

1. `command` - Executes a command (e.g., stop script)
2. `winmessage` - Sends WM_CLOSE to application windows
3. `kill` - Forces process termination (TerminateProcess)

Configure using either:
- Control file: `shutdown_method=command`
- Command line: Not directly settable, use control file

### Logging

Debug output can be directed to:
- Standard output (-o -)
- Windows Event Log (-o LOG)
- File (-o filename)

Log levels:
- 0: Errors only
- 1: Errors and information
- 2: All debug messages

### Best Practices

1. Always use control files for services
2. Set appropriate shutdown methods using `shutdown_method` directive
3. Use auto-restart for critical services
4. Configure proper logging with -o option
5. Test in command mode before installing as service
6. Use desktop interaction only when necessary
7. Always use full pathnames in configuration files
8. Don't rely on PATH environment variable for finding executables
9. Verify service account permissions
10. Test shutdown behavior before deploying to production

### Control File Configuration

Control files use an INI-style format with sections for each service:

```ini
[ServiceName]
startup=c:\path\to\program.exe
startup_dir=c:\working\directory
debug=1
debug_out=c:\logs\service.log
env=APP_HOME=c:\apps
env=LOG_DIR=c:\logs
auto_restart=yes
restart_interval=60
shutdown_method=command
shutdown=c:\path\to\stop.cmd
```

### Parameter Substitution

Two types of substitutions are supported:

1. Environment Variables:
   ```
   startup=%APP_HOME%\myapp.exe
   ```

2. Runtime Prompts:
   ```
   startup=myapp.exe -u{username} -p{-password}
   ```
   Note: Adding `-` before the prompt hides input for passwords.

### Advanced Features

#### Service Modes
- `cmd` - Run in command mode (creates a window)
- `svc` - Run as NT service
- `any` - Try service mode first, fallback to command mode
- `install` - Install as NT service
- `install_desktop` - Install as NT service with desktop interaction
- `remove` - Remove NT service

#### Shutdown Methods
LiteSRV supports three shutdown methods:
1. `command` - Executes a shutdown command
2. `winmessage` - Sends WM_CLOSE to application windows
3. `kill` - Forces process termination (TerminateProcess)

#### Environment Control
- Set custom environment variables
- Map network and local drives
- Define startup directory
- Control process priority

#### Process Management
- Automatic restart on crash
- Configurable restart intervals
- Process status monitoring
- Startup delay support

### Limitations and Requirements

- Service account needs appropriate permissions
- Desktop interaction requires specific security settings
- Network drive mapping requires stored credentials
- No support for pausing/resuming services
- Must connect to SCM within 1 second of startup

## Troubleshooting

### Common Error Messages

1. "Error 1067: The process terminated unexpectedly"
   - Verify service name matches in configuration file and command line
   - Check parameter case sensitivity (e.g., use -c not -C)
   - Ensure full pathnames in startup directive
   - Validate configuration file structure

2. "Error 2140: An internal Windows NT error occurred"
   - Usually appears when starting service from Control Panel
   - Check event log for detailed error message
   - Verify all file paths and permissions

3. "Invalid parameter was supplied"
   - Check configuration file syntax
   - Verify all required directives present
   - Ensure startup directive is in lowercase

4. "Exception 6 in Class 'CmdRunner'"
   - Invalid parameter supplied
   - Check startup directive exists
   - Verify service name spelling

### Common Issues

1. Service fails to start
   - Check service account permissions
   - Verify all paths exist
   - Check the event log for errors
   - Test in command mode first

2. Service stops unexpectedly
   - Enable debug logging with -d 2
   - Check application error handling
   - Consider using auto_restart=y
   - Review startup_delay setting

3. Network drives unavailable
   - Service must run as named user, not LocalSystem
   - Verify service account has network access
   - Check stored credentials
   - Ensure network is available before mapping

4. Shutdown hangs
   - Try different shutdown_method values
   - Set appropriate timeout values
   - Check application cleanup routines
   - Consider using kill method as last resort

### Debug Logging

Enable debug logging with:
```
debug=2
debug_out=c:\logs\service.log
```

Or command line:
```
-d 2 -o c:\logs\service.log
```

Debug levels:
- 0: Errors only
- 1: Errors and information
- 2: All debug messages

## Support

For issues and questions:
1. Check the Windows Event Log
2. Enable debug logging
3. Test in command mode
4. Verify configuration file syntax

## License

This software is provided under [Unlicence](LICENSE) terms.

## Version History

1.01 - Current version
- Initial public release
- Complete service wrapper functionality
- Environment and drive mapping support
- Multiple shutdown methods
- Logging capabilities
