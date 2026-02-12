Imports System.Reflection

Namespace LiteTask

    Module Program
        Private _mutex As Mutex = Nothing
        Private Const MutexName As String = "Global\LiteTaskApplication"
        Private ReadOnly LogBasePath As String = Path.Combine(Application.StartupPath, "LiteTaskData", "logs")
        Private ReadOnly ServiceName As String = "LiteTaskService"
        Private _isServiceMode As Boolean = False
        Private _logger As Logger

        <STAThread()>
        Public Sub Main(args As String())
            Try
                ' Register legacy codepage encodings (e.g. 863) required by process output redirection
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

                ' Ensure log directory exists
                EnsureLogDirectory()

                ' Set up assembly resolution before anything else
                AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveAssembly

                ' Check if running as service
                _isServiceMode = args.Length > 0 AndAlso args(0).Equals("-service", StringComparison.OrdinalIgnoreCase)

                If Not _isServiceMode Then
                    Application.EnableVisualStyles()
                    Application.SetCompatibleTextRenderingDefault(False)
                End If

                ' Initialize container
                InitializeContainer()
                
                ' Clean up any orphaned temp files from previous runs
                Try
                    Dim logger = ApplicationContainer.GetService(Of Logger)()
                    logger.CleanupAllTempFiles()
                    
                    ' Also cleanup config files and backups
                    Dim xmlManager = ApplicationContainer.GetService(Of XMLManager)()
                    xmlManager.CleanupConfigFiles()
                Catch ex As Exception
                    ' Log cleanup failure but don't stop application startup
                    Console.WriteLine($"Warning: Failed to cleanup temp files: {ex.Message}")
                End Try

                ' Set up global exception handlers
                AddHandler Application.ThreadException, AddressOf Application_ThreadException
                AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException

                If args.Length > 0 Then
                    HandleCommandLineArguments(args)
                Else
                    ' Check for single instance only in UI mode
                    Dim createdNew As Boolean
                    _mutex = New Mutex(True, MutexName, createdNew)
                    If Not createdNew Then
                        MessageBox.Show("Another instance of LiteTask is already running.", "LiteTask",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    ' Create and show the main application context
                    Application.Run(New ApplicationContext())
                End If

            Catch ex As Exception
                ShowDetailedError(ex)
            Finally
                If _mutex IsNot Nothing Then
                    _mutex.Close()
                    _mutex.Dispose()
                End If
            End Try
        End Sub

        Private Sub Application_ThreadException(sender As Object, e As ThreadExceptionEventArgs)
            ShowDetailedError(e.Exception)
        End Sub

        Private Sub CurrentDomain_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
            ShowDetailedError(DirectCast(e.ExceptionObject, Exception))
        End Sub

        Private Sub EnsureLogDirectory()
            Try
                ' Create base LiteTaskData directory first
                Dim dataPath = Path.Combine(Application.StartupPath, "LiteTaskData")
                If Not Directory.Exists(dataPath) Then
                    Directory.CreateDirectory(dataPath)
                End If

                ' Create logs directory
                If Not Directory.Exists(LogBasePath) Then
                    Directory.CreateDirectory(LogBasePath)
                End If
            Catch ex As Exception
                ' If we can't create log directory, we'll fall back to app directory
                Console.WriteLine($"Warning: Could not create log directory: {ex.Message}")
            End Try
        End Sub

        Private Sub ExecuteServiceCommand(command As String, arguments As String)
            Try
                Using process As New Process()
                    process.StartInfo = New ProcessStartInfo With {
                    .FileName = "sc.exe",
                    .Arguments = $"{command} {arguments}",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }

                    process.Start()
                    Dim output = process.StandardOutput.ReadToEnd()
                    Dim err = process.StandardError.ReadToEnd()
                    process.WaitForExit()

                    If process.ExitCode <> 0 Then
                        Throw New Exception($"Service command failed: {err}")
                    End If
                End Using
            Catch ex As Exception
                LogServiceError($"Error executing service command: {command}", ex)
                Throw
            End Try
        End Sub

        Private Function GetLogPath(logName As String) As String
            Try
                Return Path.Combine(LogBasePath, logName)
            Catch
                ' Fallback to application directory if there's any issue
                Return Path.Combine(Application.StartupPath, logName)
            End Try
        End Function

        'Private Sub GrantServicePrivileges(accountName As String)
        '    Try
        '        ' Grant required privileges using subinacl.exe
        '        Dim startInfo As New ProcessStartInfo() With {
        '    .FileName = "subinacl.exe",
        '    .Arguments = $"/service LiteTaskService /grant={accountName}=F",
        '    .UseShellExecute = False,
        '    .RedirectStandardOutput = True,
        '    .RedirectStandardError = True,
        '    .CreateNoWindow = True
        '}

        '        Using process As New Process() With {.StartInfo = startInfo}
        '            process.Start()
        '            process.WaitForExit()

        '            If process.ExitCode <> 0 Then
        '                '_logger?.LogWarning($"Failed to grant service privileges to {accountName}")
        '            End If
        '        End Using

        '    Catch ex As Exception
        '        '_logger?.LogError($"Error granting service privileges: {ex.Message}")
        '        ' Continue execution as this is not critical
        '    End Try
        'End Sub

        Private Sub HandleCommandLineArguments(args As String())
            If Not IsElevated() AndAlso Array.Exists(args, Function(arg) arg.ToLower() = "-register" OrElse
                                                                arg.ToLower() = "-unregister" OrElse
                                                                arg.ToLower() = "-service") Then
                RestartAsAdmin(String.Join(" ", args))
                Return
            End If

            Select Case args(0).ToLower()
                Case "-service"
                    RunAsService()
                Case "-register"
                    RegisterService()
                Case "-unregister"
                    UnregisterService()
                Case "-runtask"
                    If args.Length > 1 Then
                        RunTaskFromCommandLine(args(1))
                    Else
                        ShowHelp()
                    End If
                Case "-debug"
                    RunInDebugMode()
                Case "-elevated"
                    ' Handle elevated mode specific operations
                    HandleElevatedMode()
                Case Else
                    ShowHelp()
            End Select
        End Sub

        Private Sub HandleCriticalError(ex As Exception)
            Try
                ' Attempt to get logger if available
                Dim logger = TryGetService(Of Logger)()
                logger?.LogCritical($"Critical application error: {ex.Message}", ex)

                If Environment.UserInteractive Then
                    MessageBox.Show($"A critical error occurred: {ex.Message}{Environment.NewLine}{Environment.NewLine}Error Details: {ex.ToString()}",
                                  "Critical Error",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch criticalEx As Exception
                ' Last resort error handling
                Console.WriteLine($"Critical Error: {criticalEx.Message}")
                MessageBox.Show($"Critical Error: {criticalEx.Message}", "Critical Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub HandleElevatedMode()
            Try
                ' Perform elevated operations here
                _logger?.LogInfo("Application running in elevated mode")

                ' Initialize with elevated privileges
                InitializeContainer()

                ' Run the application
                Application.Run(New ApplicationContext())

            Catch ex As Exception
                ShowDetailedError(ex)
            End Try
        End Sub

        Private Sub InitializeEventLogSource()
            If Not EventLog.SourceExists(ServiceName) Then
                EventLog.CreateEventSource(ServiceName, "Application")
            End If
        End Sub

        Private Function IsElevated() As Boolean
            Try
                Dim identity = WindowsIdentity.GetCurrent()
                Dim principal = New WindowsPrincipal(identity)
                Return principal.IsInRole(WindowsBuiltInRole.Administrator)
            Catch
                Return False
            End Try
        End Function

        Public Function IsUserAdministrator() As Boolean
            Try
                Dim identity = WindowsIdentity.GetCurrent()
                Dim principal = New WindowsPrincipal(identity)
                Return principal.IsInRole(WindowsBuiltInRole.Administrator)
            Catch
                Return False
            End Try
        End Function

        Private Sub LogAssemblyResolution(message As String)
            Try
                Dim logPath = GetLogPath("assembly_resolution.log")
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}")
            Catch
                ' Ignore logging errors
            End Try
        End Sub

        Public Sub LogServiceError(message As String, ex As Exception)
            Try
                Dim logPath = GetLogPath("service_error.log")
                Dim entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}" &
                       $"Error: {ex.Message}{Environment.NewLine}" &
                       $"Stack Trace: {ex.StackTrace}{Environment.NewLine}"

                File.AppendAllText(logPath, entry)

                If EventLog.SourceExists(ServiceName) Then
                    EventLog.WriteEntry(ServiceName, entry, EventLogEntryType.Error)
                End If
            Catch
                ' Ignore logging errors in error handler
            End Try
        End Sub

        Public Sub InitializeContainer()
            Try
                ApplicationContainer.Initialize()
            Catch ex As Exception
                LogAssemblyResolution($"Container initialization failed: {ex.Message}")
                If ex.InnerException IsNot Nothing Then
                    LogAssemblyResolution($"Inner exception: {ex.InnerException.Message}")
                End If
                Throw New Exception("Failed to initialize application services", ex)
            End Try
        End Sub

        Private Sub RegisterService()
            Try
                ' Use InstallUtil to register the service
                If Not IsUserAdministrator() Then
                    RestartAsAdmin("-register")
                    Return
                End If

                Dim identity = WindowsIdentity.GetCurrent()
                Dim currentUser = identity.Name
                Dim exePath = Assembly.GetExecutingAssembly().Location
                Dim installUtilPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "InstallUtil.exe")

                ' Configure service account and permissions
                Dim startInfo = New ProcessStartInfo With {
            .FileName = installUtilPath,
            .Arguments = $"/ServiceAccount=LocalSystem /elevated ""{exePath}""",
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True,
            .Verb = "runas"  ' Run elevated
        }

                Using process As New Process With {.StartInfo = startInfo}
                    process.Start()
                    Dim output = process.StandardOutput.ReadToEnd()
                    Dim err = process.StandardError.ReadToEnd()
                    process.WaitForExit()

                    If process.ExitCode <> 0 Then
                        Throw New Exception($"Service installation failed. Error: {err}")
                    End If
                End Using

                MessageBox.Show("Service registered successfully.", "Service Installation",
                      MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show($"Error registering service: {ex.Message}", "Service Installation Error",
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function ResolveAssembly(sender As Object, args As ResolveEventArgs) As Assembly
            Try
                ' Get the assembly name
                Dim assemblyName = New AssemblyName(args.Name)

                ' Skip resource assembly resolution attempts
                If assemblyName.Name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If

                ' First try the lib folder
                Dim libPath As String = Path.Combine(Application.StartupPath, "lib")
                Dim assemblyPath As String = Path.Combine(libPath, assemblyName.Name & ".dll")

                'LogAssemblyResolution($"Trying to resolve assembly: {assemblyName.Name}")
                'LogAssemblyResolution($"Looking in lib folder: {assemblyPath}")

                If File.Exists(assemblyPath) Then
                    'LogAssemblyResolution($"Found assembly in lib folder: {assemblyPath}")
                    Return Assembly.LoadFrom(assemblyPath)
                End If

                ' If not found in lib, try the application root
                assemblyPath = Path.Combine(Application.StartupPath, assemblyName.Name & ".dll")
                'LogAssemblyResolution($"Looking in root folder: {assemblyPath}")

                If File.Exists(assemblyPath) Then
                    'LogAssemblyResolution($"Found assembly in root folder: {assemblyPath}")
                    Return Assembly.LoadFrom(assemblyPath)
                End If

                ' Log failure to find assembly
                'LogAssemblyResolution($"Failed to find assembly: {assemblyName.Name}")
                Return Nothing

            Catch ex As Exception
                LogAssemblyResolution($"Error resolving assembly: {ex.Message}")
                Return Nothing
            End Try
        End Function


        Private Sub RestartAsAdmin(arguments As String)
            Try
                Dim startInfo As New ProcessStartInfo() With {
            .UseShellExecute = True,
            .WorkingDirectory = Environment.CurrentDirectory,
            .FileName = Application.ExecutablePath,
            .Verb = "runas"
        }

                If Not String.IsNullOrEmpty(arguments) Then
                    startInfo.Arguments = arguments
                End If

                Process.Start(startInfo)
                Application.Exit()
            Catch ex As Exception
                MessageBox.Show("This operation requires administrative privileges.",
                       "Elevation Required",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Warning)
            End Try
        End Sub

        Public Sub RunAsService()
            Try
                If Not IsUserAdministrator() Then
                    RestartAsAdmin("-service")
                    Return
                End If

                InitializeEventLogSource()
                EventLog.WriteEntry(ServiceName, "Starting service...", EventLogEntryType.Information)

                ' Initialize app with secure defaults
                InitializeContainer()
                
                ' Clean up any orphaned temp files from previous runs
                Try
                    Dim logger = ApplicationContainer.GetService(Of Logger)()
                    logger.CleanupAllTempFiles()
                    
                    ' Also cleanup config files and backups
                    Dim xmlManager = ApplicationContainer.GetService(Of XMLManager)()
                    xmlManager.CleanupConfigFiles()
                Catch ex As Exception
                    EventLog.WriteEntry(ServiceName, $"Warning: Failed to cleanup temp files: {ex.Message}", EventLogEntryType.Warning)
                End Try
                
                Dim service = ApplicationContainer.GetService(Of LiteTaskService)()

                Dim servicesToRun() As ServiceBase = {service}
                ServiceBase.Run(servicesToRun)

            Catch ex As Exception
                LogServiceError("Error starting service", ex)
                EventLog.WriteEntry(ServiceName, $"Error starting service: {ex.Message}",
                EventLogEntryType.Error)
                Throw
            End Try
        End Sub

        Private Sub RunInDebugMode()

            Try
                Console.WriteLine("Running in debug mode...")
                Dim service As New ServiceController("LiteTaskService")
                Dim logger = ApplicationContainer.GetService(Of Logger)()
                logger.LogInfo("Starting debug mode")
                If service.Status = ServiceControllerStatus.Stopped Then
                    service.Start()
                    service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10))
                    Console.WriteLine("Service started successfully.")
                Else
                    Console.WriteLine("Service is already running.")
                End If
            Catch ex As Exception
                HandleCriticalError(ex)
            End Try

        End Sub

        Private Sub RunTaskFromCommandLine(taskName As String)
            Try
                Dim scheduler = ApplicationContainer.GetService(Of CustomScheduler)()
                Dim logger = ApplicationContainer.GetService(Of Logger)()

                Dim task = scheduler.GetTask(taskName)
                If task Is Nothing Then
                    logger.LogError($"Task '{taskName}' not found")
                    Environment.Exit(1)
                    Return
                End If

                scheduler.RunTaskAsync(task).Wait()
                Environment.Exit(0)
            Catch ex As Exception
                Dim logger = ApplicationContainer.GetService(Of Logger)()
                logger.LogError($"Error running task from command line: {ex.Message}")
                Environment.Exit(1)
            End Try
        End Sub

        Private Sub ShowDetailedError(ex As Exception)
            Dim errorBuilder As New System.Text.StringBuilder()
            errorBuilder.AppendLine("A fatal error occurred while starting the application:")
            errorBuilder.AppendLine()
            errorBuilder.AppendLine($"Error: {ex.Message}")

            If ex.InnerException IsNot Nothing Then
                errorBuilder.AppendLine()
                errorBuilder.AppendLine($"Details: {ex.InnerException.Message}")
            End If

            errorBuilder.AppendLine()
            errorBuilder.AppendLine("Stack Trace:")
            errorBuilder.AppendLine(ex.StackTrace)

            If ex.InnerException IsNot Nothing Then
                errorBuilder.AppendLine()
                errorBuilder.AppendLine("Inner Exception Stack Trace:")
                errorBuilder.AppendLine(ex.InnerException.StackTrace)
            End If

            errorBuilder.AppendLine()
            errorBuilder.AppendLine("The application will now close.")

            ' Log error to file
            Try
                Dim logPath = GetLogPath("startup_error.log")
                File.WriteAllText(logPath, errorBuilder.ToString())
            Catch
                ' Ignore logging errors
            End Try

            ' Show message box only if running in interactive mode
            If Environment.UserInteractive Then
                MessageBox.Show(errorBuilder.ToString(), "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' Log to Windows Event Log when running as service
                Try
                    EventLog.WriteEntry("LiteTaskService", errorBuilder.ToString(), EventLogEntryType.Error)
                Catch
                    ' Ignore event log errors
                End Try
            End If

            Environment.Exit(1)
        End Sub

        Private Sub ShowHelp()
            Console.WriteLine("LiteTask - Command Line Options:")
            Console.WriteLine("  -service    - Run as Windows service")
            Console.WriteLine("  -register   - Register the Windows service")
            Console.WriteLine("  -unregister - Unregister the Windows service")
            Console.WriteLine("  -runtask    - Run a specific task (e.g., -runtask TaskName)")
            Console.WriteLine("  -debug      - Run in debug mode")
            Console.WriteLine("  -help       - Show this help message")
        End Sub

        'Private Sub ShowFatalError(ex As Exception)
        '    Dim errorMessage = $"A fatal error occurred while starting the application:{Environment.NewLine}{Environment.NewLine}" &
        '                     $"Error: {ex.Message}{Environment.NewLine}{Environment.NewLine}"

        '    If ex.InnerException IsNot Nothing Then
        '        errorMessage &= $"Details: {ex.InnerException.Message}{Environment.NewLine}"
        '    End If

        '    errorMessage &= $"{Environment.NewLine}The application will now close."

        '    MessageBox.Show(errorMessage, "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Environment.Exit(1)
        'End Sub

        Private Function TryGetService(Of T)() As T
            Try
                Return ApplicationContainer.GetService(Of T)()
            Catch
                Return Nothing
            End Try
        End Function

        Private Sub UnregisterService()
            Try
                ' Use InstallUtil to unregister the service
                Dim exePath = Assembly.GetExecutingAssembly().Location
                Dim installUtilPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "InstallUtil.exe")

                Dim startInfo = New ProcessStartInfo With {
                    .FileName = installUtilPath,
                    .Arguments = $"/u ""{exePath}""",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }

                Using process As New Process With {.StartInfo = startInfo}
                    process.Start()
                    Dim output = process.StandardOutput.ReadToEnd()
                    Dim err = process.StandardError.ReadToEnd()
                    process.WaitForExit()

                    If process.ExitCode <> 0 Then
                        Throw New Exception($"Service uninstallation failed. Error: {err}")
                    End If
                End Using

                MessageBox.Show("Service unregistered successfully.", "Service Uninstallation",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show($"Error unregistering service: {ex.Message}", "Service Uninstallation Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

    End Module
End Namespace
