Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Timers


Namespace LiteTask
    Public Class LiteTaskService
        Inherits ServiceBase
        Private WithEvents _timer As Timer
        Private _customScheduler As CustomScheduler
        Private _credentialManager As CredentialManager
        Private _xmlManager As XMLManager
        Private _logger As Logger
        Private ReadOnly _toolManager As ToolManager

        Private Const SC_MANAGER_ALL_ACCESS As Integer = &HF003F
        Private Const SERVICE_ALL_ACCESS As Integer = &HF01FF
        Private Const SC_STATUS_PROCESS_INFO As Integer = 0
        Private Const PROCESS_ALL_ACCESS As Integer = &H1F0FFF
        Private Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
        Private Const TOKEN_QUERY As Integer = &H8
        Private Const SE_PRIVILEGE_ENABLED As Integer = &H2
        Private Const SE_IMPERSONATE_NAME As String = "SeImpersonatePrivilege"
        Private Const ERROR_NOT_ALL_ASSIGNED As Integer = 1300

        <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function OpenService(hSCManager As IntPtr, lpServiceName As String, dwDesiredAccess As Integer) As IntPtr
        End Function

        <DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function QueryServiceStatusEx(hService As IntPtr, infoLevel As Integer, ByRef lpBuffer As SERVICE_STATUS_PROCESS, dwBufSize As Integer, ByRef pcbBytesNeeded As Integer) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function OpenProcess(ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Integer) As IntPtr
        End Function

        <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Private Shared Function OpenSCManager(
        ByVal machineName As String,
        ByVal databaseName As String,
        ByVal desiredAccess As Integer
    ) As IntPtr
        End Function

        <DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function AdjustTokenPrivileges(
        ByVal TokenHandle As IntPtr,
        ByVal DisableAllPrivileges As Boolean,
        ByRef NewState As TOKEN_PRIVILEGES,
        ByVal BufferLength As Integer,
        ByRef PreviousState As TOKEN_PRIVILEGES,
        ByRef ReturnLength As Integer
    ) As Boolean
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Private Structure LUID
            Public LowPart As UInteger
            Public HighPart As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure LUID_AND_ATTRIBUTES
            Public Luid As LUID
            Public Attributes As UInteger
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure TOKEN_PRIVILEGES
            Public PrivilegeCount As Integer
            Public Privileges As LUID_AND_ATTRIBUTES
        End Structure

        <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Private Shared Function LookupPrivilegeValue(lpSystemName As String, lpName As String, ByRef lpLuid As LUID) As Boolean
        End Function

        <DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function AdjustTokenPrivileges(TokenHandle As IntPtr, DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, BufferLength As Integer, PreviousState As IntPtr, ReturnLength As IntPtr) As Boolean
        End Function

        <DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function OpenProcessToken(ProcessHandle As IntPtr, DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function CloseHandle(hObject As IntPtr) As Boolean
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Private Structure SERVICE_STATUS_PROCESS
            Public dwServiceType As Integer
            Public dwCurrentState As Integer
            Public dwControlsAccepted As Integer
            Public dwWin32ExitCode As Integer
            Public dwServiceSpecificExitCode As Integer
            Public dwCheckPoint As Integer
            Public dwWaitHint As Integer
            Public dwProcessId As Integer
            Public dwServiceFlags As Integer
        End Structure


        Private ReadOnly _cancellationTokenSource As New CancellationTokenSource()
        Private _serviceTask As Task

        Public Sub New(customScheduler As CustomScheduler, credentialManager As CredentialManager,
                  xmlManager As XMLManager, toolManager As ToolManager)
            InitializeComponent()

            ServiceName = "LiteTaskService"
            CanStop = True
            CanPauseAndContinue = False
            AutoLog = True

            _customScheduler = customScheduler
            _credentialManager = credentialManager
            _xmlManager = xmlManager
            _toolManager = toolManager
            _logger = ApplicationContainer.GetService(Of Logger)()
        End Sub

        Public Sub GrantServiceImpersonatePrivilege(serviceName As String)
            Const SE_PRIVILEGE_ENABLED As Integer = &H2
            Const SE_IMPERSONATE_NAME As String = "SeImpersonatePrivilege"

            Try
                ' Open the service
                Dim scManager = OpenSCManager(Nothing, Nothing, SC_MANAGER_ALL_ACCESS)
                If scManager = IntPtr.Zero Then
                    Throw New ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
                End If

                Dim service = OpenService(scManager, serviceName, SERVICE_ALL_ACCESS)
                If service = IntPtr.Zero Then
                    Throw New ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
                End If

                ' Get the service process ID
                Dim serviceStatus As New SERVICE_STATUS_PROCESS()
                Dim bytesNeeded As UInteger
                If Not QueryServiceStatusEx(service, SC_STATUS_PROCESS_INFO, serviceStatus, Marshal.SizeOf(serviceStatus), bytesNeeded) Then
                    Throw New ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
                End If

                ' Open the service process
                Dim processHandle = OpenProcess(PROCESS_ALL_ACCESS, False, serviceStatus.dwProcessId)
                If processHandle = IntPtr.Zero Then
                    Throw New ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
                End If

                ' Get the process token
                Dim tokenHandle As IntPtr
                If Not OpenProcessToken(processHandle, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, tokenHandle) Then
                    Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
                End If

                ' Enable the SeImpersonatePrivilege
                Dim tp As New TOKEN_PRIVILEGES()
                tp.PrivilegeCount = 1
                tp.Privileges.Attributes = SE_PRIVILEGE_ENABLED
                If Not LookupPrivilegeValue(Nothing, SE_IMPERSONATE_NAME, tp.Privileges.Luid) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

                Dim previousState As TOKEN_PRIVILEGES
                Dim returnLength As Integer

                If Not AdjustTokenPrivileges(tokenHandle, False, tp, Marshal.SizeOf(tp), previousState, returnLength) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

                ' Check if the privilege was actually adjusted
                If Marshal.GetLastWin32Error() = ERROR_NOT_ALL_ASSIGNED Then
                    Throw New Win32Exception(ERROR_NOT_ALL_ASSIGNED, "The token does not have the specified privilege.")
                End If

                _logger.LogInfo($"SeImpersonatePrivilege granted to service: {serviceName}")
            Catch ex As Exception
                _logger.LogError($"Failed to grant SeImpersonatePrivilege to service {serviceName}: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub InitializeComponent()
            ServiceName = "LiteTaskService"
            CanStop = True
            CanPauseAndContinue = False
            AutoLog = True
        End Sub

        Protected Overrides Sub OnStart(args() As String)
            Try
                _logger.LogInfo("LiteTaskService is starting.")

                _customScheduler.ClearTaskStates()
                _customScheduler.LoadTasks()

                _serviceTask = Task.Run(Async Function()
                                            Try
                                                While Not _cancellationTokenSource.Token.IsCancellationRequested
                                                    Try
                                                        _customScheduler.CheckAndExecuteTasks()
                                                    Catch ex As Exception
                                                        _logger.LogError($"Error checking tasks: {ex.Message}")
                                                    End Try
                                                    Await Task.Delay(60000, _cancellationTokenSource.Token)
                                                End While
                                            Catch ex As Exception When TypeOf ex Is TaskCanceledException
                                                ' Normal service stop
                                            Catch ex As Exception
                                                _logger.LogError($"Service loop error: {ex.Message}")
                                                Throw
                                            End Try
                                        End Function, _cancellationTokenSource.Token)

                Task.Run(Sub()
                             Try
                                 If IsUserAdministrator() Then
                                     GrantServiceImpersonatePrivilege("LiteTaskService")
                                     _logger.LogInfo("Successfully granted service impersonation privileges")
                                 Else
                                     _logger.LogWarning("Service not running with admin privileges")
                                 End If
                             Catch ex As Exception
                                 _logger.LogWarning($"Could not grant privileges: {ex.Message}")
                             End Try
                         End Sub)

                _logger.LogInfo("LiteTaskService started successfully.")

            Catch ex As Exception
                _logger.LogError($"Error in OnStart: {ex.Message}")
                EventLog.WriteEntry("LiteTaskService", $"Error in OnStart: {ex.Message}", EventLogEntryType.Error)
                Throw
            End Try
        End Sub

        Protected Overrides Sub OnStop()
            Try
                _logger.LogInfo("Service stopping...")
                _cancellationTokenSource.Cancel()

                If _serviceTask IsNot Nothing Then
                    _serviceTask.Wait(TimeSpan.FromSeconds(30))
                End If

                _customScheduler.SaveTasks()
                _logger.LogInfo("Service stopped successfully")

            Catch ex As Exception
                _logger.LogError($"Error stopping service: {ex.Message}")
            Finally
                _cancellationTokenSource.Dispose()
            End Try
        End Sub

        Private Function IsServiceInstalled() As Boolean
            Try
                Using sc = New ServiceController("LiteTaskService")
                    Return True
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Function GetServiceStatus() As ServiceControllerStatus
            Try
                Using sc = New ServiceController("LiteTaskService")
                    Return sc.Status
                End Using
            Catch
                Return ServiceControllerStatus.Stopped
            End Try
        End Function

        Private Sub HandleServiceOperation(operation As String)
            If Not IsElevated() Then
                RestartAsAdmin($"-{operation}")
                Return
            End If

            Try
                If Not IsServiceInstalled() AndAlso operation <> "install" Then
                    MessageBox.Show("Service is not installed.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Select Case operation.ToLower()
                    Case "start"
                        If GetServiceStatus() = ServiceControllerStatus.Running Then
                            MessageBox.Show("Service is already running.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Return
                        End If
                        Using sc = New ServiceController("LiteTaskService")
                            sc.Start()
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30))
                        End Using
                        MessageBox.Show("Service started successfully.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case "stop"
                        If GetServiceStatus() = ServiceControllerStatus.Stopped Then
                            MessageBox.Show("Service is already stopped.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Return
                        End If
                        Using sc = New ServiceController("LiteTaskService")
                            sc.Stop()
                            sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30))
                        End Using
                        MessageBox.Show("Service stopped successfully.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Case "install"
                        If IsServiceInstalled() Then
                            MessageBox.Show("Service is already installed.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If
                        RegisterService()

                    Case "uninstall"
                        If GetServiceStatus() = ServiceControllerStatus.Running Then
                            Using sc = New ServiceController("LiteTaskService")
                                sc.Stop()
                                sc.WaitForStatus(ServiceControllerStatus.Stopped)
                            End Using
                        End If
                        Process.Start("sc.exe", "delete LiteTaskService").WaitForExit()
                        MessageBox.Show("Service uninstalled successfully.", "Service Operation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                End Select

            Catch ex As Exception
                _logger?.LogError($"Error in HandleServiceOperation: {ex.Message}")
                MessageBox.Show($"Error performing service operation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function IsElevated() As Boolean
            Try
                Dim identity = WindowsIdentity.GetCurrent()
                Dim principal = New WindowsPrincipal(identity)
                Return principal.IsInRole(WindowsBuiltInRole.Administrator)
            Catch ex As Exception
                _logger?.LogError($"Error checking elevation: {ex.Message}")
                Return False
            End Try
        End Function

        Private Sub RestartAsAdmin(arguments As String)
            Try
                Dim startInfo As New ProcessStartInfo() With {
            .UseShellExecute = True,
            .WorkingDirectory = Environment.CurrentDirectory,
            .FileName = Application.ExecutablePath,
            .Arguments = arguments,
            .Verb = "runas"
        }

                Process.Start(startInfo)
                Application.Exit()
            Catch ex As Exception
                MessageBox.Show("Failed to restart with elevated privileges. Please run as administrator.",
                       "Elevation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Public Sub EnsureRequiredPermissions()
            Try
                Dim identity = WindowsIdentity.GetCurrent()
                Dim principal = New WindowsPrincipal(identity)

                If Not principal.IsInRole(WindowsBuiltInRole.Administrator) Then
                    Throw New UnauthorizedAccessException("The service requires administrative privileges.")
                End If

                ' Check for SeImpersonatePrivilege
                If Not HasSeImpersonatePrivilege() Then
                    RequestSeImpersonatePrivilege()
                End If

                ' Verify network service permissions if needed
                VerifyNetworkServicePermissions()

            Catch ex As Exception
                _logger?.LogError($"Error ensuring permissions: {ex.Message}")
                Throw New UnauthorizedAccessException("Failed to verify or obtain required permissions.", ex)
            End Try
        End Sub

        Private Function HasSeImpersonatePrivilege() As Boolean
            Dim tokenHandle As IntPtr = IntPtr.Zero
            Try
                If Not OpenProcessToken(Process.GetCurrentProcess().Handle, TOKEN_QUERY, tokenHandle) Then
                    Return False
                End If

                Dim tp As New TOKEN_PRIVILEGES()
                Dim luid As New LUID()

                If Not LookupPrivilegeValue(Nothing, "SeImpersonatePrivilege", luid) Then
                    Return False
                End If

                tp.PrivilegeCount = 1
                tp.Privileges.Luid = luid
                tp.Privileges.Attributes = SE_PRIVILEGE_ENABLED

                Return AdjustTokenPrivileges(tokenHandle, False, tp, 0, IntPtr.Zero, 0)
            Finally
                If tokenHandle <> IntPtr.Zero Then
                    CloseHandle(tokenHandle)
                End If
            End Try
        End Function

        Private Sub RequestSeImpersonatePrivilege()
            Dim tokenHandle As IntPtr = IntPtr.Zero
            Try
                If Not OpenProcessToken(Process.GetCurrentProcess().Handle,
                              TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY,
                              tokenHandle) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

                Dim tp As New TOKEN_PRIVILEGES()
                Dim luid As New LUID()

                If Not LookupPrivilegeValue(Nothing, "SeImpersonatePrivilege", luid) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

                tp.PrivilegeCount = 1
                tp.Privileges.Luid = luid
                tp.Privileges.Attributes = SE_PRIVILEGE_ENABLED

                If Not AdjustTokenPrivileges(tokenHandle, False, tp, 0, IntPtr.Zero, 0) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If
            Finally
                If tokenHandle <> IntPtr.Zero Then
                    CloseHandle(tokenHandle)
                End If
            End Try
        End Sub

        Private Sub VerifyNetworkServicePermissions()
            Dim testPaths() As String = {
        Path.Combine(Application.StartupPath, "LiteTaskData"),
        Path.Combine(Application.StartupPath, "LiteTaskData", "logs"),
        Path.Combine(Application.StartupPath, "LiteTaskData", "temp")
    }

            For Each _path In testPaths
                Try
                    Dim testFile = Path.Combine(_path, "permission_test.tmp")
                    File.WriteAllText(testFile, "test")
                    File.Delete(testFile)
                Catch ex As Exception
                    Throw New UnauthorizedAccessException(
                $"NetworkService account lacks required permissions on path: {_path}", ex)
                End Try
            Next
        End Sub

        Private Sub RegisterService()
            Try
                Dim exePath = Assembly.GetExecutingAssembly().Location
                Dim accountName = "NT AUTHORITY\NetworkService"

                Using serviceController = New ServiceController("LiteTaskService")
                    If serviceController.Status <> ServiceControllerStatus.Stopped Then
                        serviceController.Stop()
                        serviceController.WaitForStatus(ServiceControllerStatus.Stopped)
                    End If
                End Using

                Try
                    Process.Start("sc.exe", $"delete LiteTaskService").WaitForExit()
                    Thread.Sleep(1000)
                Catch
                End Try

                Dim createArgs = $"create LiteTaskService binPath= ""{exePath} -service"" " &
                                $"start= auto " &
                                $"obj= {accountName} " &
                                $"DisplayName= ""LiteTask Scheduler Service"""

                Process.Start("sc.exe", createArgs).WaitForExit()
                Process.Start("sc.exe", "description LiteTaskService ""Task scheduling and automation service""").WaitForExit()

                Dim sid = GetServiceSid("LiteTaskService")
                GrantServiceLogonRight(accountName)
                GrantSeImpersonatePrivilege(sid)
                GrantSeServiceLogonRight(sid)

                Process.Start("sc.exe", "failure LiteTaskService reset= 86400 actions= restart/60000/restart/60000/restart/60000").WaitForExit()

                MessageBox.Show("Service registered successfully.", "Service Installation",
                               MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show($"Error registering service: {ex.Message}", "Service Installation Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function GetServiceSid(serviceName As String) As String
            Dim output As String = ""
            Dim startInfo As New ProcessStartInfo("sc.exe", $"showsid {serviceName}") With {
                .RedirectStandardOutput = True,
                .UseShellExecute = False
            }

            Using process As New Process() With {.StartInfo = startInfo}
                process.Start()
                output = process.StandardOutput.ReadToEnd()
                process.WaitForExit()
            End Using

            Dim match = Regex.Match(output, "SERVICE SID:\s+(\S+)")
            Return If(match.Success, match.Groups(1).Value, String.Empty)
        End Function

        Private Sub GrantServiceLogonRight(account As String)
            Dim startInfo As New ProcessStartInfo("ntrights.exe", $"+r SeServiceLogonRight -u ""{account}""") With {
                .UseShellExecute = False
            }
            Process.Start(startInfo).WaitForExit()
        End Sub

        Private Sub GrantSeImpersonatePrivilege(sid As String)
            Process.Start("subinacl.exe", $"/service LiteTaskService /grant={sid}=TO").WaitForExit()
        End Sub

        Private Sub GrantSeServiceLogonRight(sid As String)
            Process.Start("subinacl.exe", $"/service LiteTaskService /grant={sid}=R").WaitForExit()
        End Sub

        Private Sub RunAsService()
            Try
                If Not EventLog.SourceExists("LiteTaskService") Then
                    EventLog.CreateEventSource("LiteTaskService", "Application")
                End If

                InitializeContainer()

                Using service = ApplicationContainer.GetService(Of LiteTaskService)()
                    EventLog.WriteEntry("LiteTaskService", "Starting service...", EventLogEntryType.Information)

                    service.EnsureRequiredPermissions()

                    Dim servicesToRun() As ServiceBase = {service}
                    ServiceBase.Run(servicesToRun)
                End Using

            Catch ex As Exception
                LogServiceError("Error starting service", ex)
                EventLog.WriteEntry("LiteTaskService", $"Error starting service: {ex.Message}", EventLogEntryType.Error)
                Throw
            Finally
                Try
                    ApplicationContainer.Dispose()
                Catch disposeEx As Exception
                    LogServiceError("Error disposing container", disposeEx)
                End Try
            End Try
        End Sub
    End Class
End Namespace
