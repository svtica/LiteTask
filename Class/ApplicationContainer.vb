Namespace LiteTask
    Public NotInheritable Class ApplicationContainer
        Private Sub New()
        End Sub

        Private Shared _serviceProvider As IServiceProvider
        Private Shared ReadOnly _lock As New Object()
        Private Shared ReadOnly _enableInitializationLogging As Boolean = False

        Public Shared Sub Dispose()
            If _serviceProvider IsNot Nothing Then
                If TypeOf _serviceProvider Is IDisposable Then
                    DirectCast(_serviceProvider, IDisposable).Dispose()
                End If
                _serviceProvider = Nothing
            End If
        End Sub

        Private Shared Function GetLogPath(logName As String) As String
            Try
                Dim logPath = Path.Combine(Application.StartupPath, "LiteTaskData", "logs", logName)
                ' Ensure directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(logPath))
                Return logPath
            Catch
                ' Fallback to application directory if there's any issue
                Return Path.Combine(Application.StartupPath, logName)
            End Try
        End Function

        Public Shared Function GetService(Of T)() As T
            If _serviceProvider Is Nothing Then
                Throw New InvalidOperationException("Service provider is not initialized. Call Initialize() first.")
            End If
            Return _serviceProvider.GetRequiredService(Of T)()
        End Function

        Public Shared Sub Initialize()
            SyncLock _lock
                If _serviceProvider IsNot Nothing Then
                    Return
                End If

                Dim services = New ServiceCollection()

                Try
                    ' Define paths
                    Dim appDataPath = Path.Combine(Application.StartupPath, "LiteTaskData")
                    Dim toolsPath = Path.Combine(appDataPath, "tools")
                    Dim tempPath = Path.Combine(appDataPath, "temp")
                    Dim logsPath = Path.Combine(appDataPath, "logs")
                    Dim settingsPath = Path.Combine(appDataPath, "settings.xml")
                    LogInitialization($"Paths defined. AppDataPath: {appDataPath}")

                    ' Create directories
                    Try
                        Directory.CreateDirectory(appDataPath)
                        Directory.CreateDirectory(toolsPath)
                        Directory.CreateDirectory(tempPath)
                        Directory.CreateDirectory(logsPath)
                        LogInitialization("Directories created successfully")
                    Catch dirEx As Exception
                        Throw New InvalidOperationException("Failed to create required directories", dirEx)
                    End Try

                    ' Register configuration services
                    Try
                        ' Register PathConfiguration
                        services.AddSingleton(New PathConfiguration With {
                    .AppDataPath = appDataPath,
                    .ToolsPath = toolsPath,
                    .TempPath = tempPath,
                    .LogsPath = logsPath,
                    .SettingsPath = settingsPath
                })
                        LogInitialization("PathConfiguration registered")

                        ' Register XMLManager first with path validation
                        services.AddSingleton(Of XMLManager)(Function(sp)
                                                                 Try
                                                                     If Not Directory.Exists(Path.GetDirectoryName(settingsPath)) Then
                                                                         Directory.CreateDirectory(Path.GetDirectoryName(settingsPath))
                                                                     End If
                                                                     Return New XMLManager(settingsPath)
                                                                 Catch ex As Exception
                                                                     LogInitialization($"XMLManager initialization failed: {ex.Message}")
                                                                     Throw
                                                                 End Try
                                                             End Function)
                        LogInitialization("XMLManager registered")

                        ' Register Logger with delayed XMLManager configuration
                        services.AddSingleton(Of Logger)(Function(sp)
                                                             Try
                                                                 Dim logPath = Path.Combine(logsPath, "app_log.txt")
                                                                 If Not Directory.Exists(logsPath) Then
                                                                     Directory.CreateDirectory(logsPath)
                                                                 End If
                                                                 Return New Logger(sp.GetRequiredService(Of XMLManager)(), logPath)
                                                             Catch ex As Exception
                                                                 LogInitialization($"Logger initialization failed: {ex.Message}")
                                                                 Throw
                                                             End Try
                                                         End Function)
                        LogInitialization("Logger registered")

                        ' Register core services
                        services.AddSingleton(Of ToolManager)(Function(sp) New ToolManager(toolsPath))
                        services.AddSingleton(Of CredentialManager)(Function(sp) New CredentialManager(sp.GetRequiredService(Of Logger)()))
                        LogInitialization("Core services registered")

                        ' Register remaining services
                        RegisterRemainingServices(services)
                        LogInitialization("All services registered")

                    Catch regEx As Exception
                        Throw New InvalidOperationException("Failed to register services", regEx)
                    End Try

                    Try
                        _serviceProvider = services.BuildServiceProvider(New ServiceProviderOptions With {
                    .ValidateOnBuild = True,
                    .ValidateScopes = True
                })
                        LogInitialization("ServiceProvider built successfully")

                        ' Validate core services
                        ValidateRequiredServices()
                        LogInitialization("Service validation completed")

                    Catch buildEx As Exception
                        Throw New InvalidOperationException("Failed to build service provider", buildEx)
                    End Try

                Catch ex As Exception
                    _serviceProvider = Nothing
                    LogInitializationError(ex)
                    Throw
                End Try
            End SyncLock
        End Sub

        Private Shared Sub RegisterRemainingServices(services As IServiceCollection)
            ' Register TaskRunner
            services.AddSingleton(Of TaskRunner)(Function(sp)
                                                     Dim pathConfig = sp.GetRequiredService(Of PathConfiguration)()
                                                     Return New TaskRunner(
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of CredentialManager)(),
            sp.GetRequiredService(Of ToolManager)(),
            pathConfig.LogsPath,
            sp.GetRequiredService(Of XMLManager)()
        )
                                                 End Function)

            ' Register NotificationManager
            services.AddSingleton(Of NotificationManager)(Function(sp)
                                                              Return New NotificationManager(
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of XMLManager)()
        )
                                                          End Function)

            ' Register TranslationManager
            services.AddSingleton(Of TranslationManager)(Function(sp)
                                                             Dim translationManager = New TranslationManager(
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of XMLManager)()
        )
                                                             TranslationManager.Initialize(
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of XMLManager)()
        )
                                                             Return translationManager
                                                         End Function)

            ' Register CustomScheduler
            services.AddSingleton(Of CustomScheduler)(Function(sp)
                                                          Return New CustomScheduler(
            sp.GetRequiredService(Of CredentialManager)(),
            sp.GetRequiredService(Of XMLManager)(),
            sp.GetRequiredService(Of ToolManager)(),
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of TaskRunner)()
        )
                                                      End Function)

            ' Register PowerShellPathManager
            services.AddSingleton(Of PowerShellPathManager)(Function(sp)
                                                                Return New PowerShellPathManager(
            sp.GetRequiredService(Of Logger)()
        )
                                                            End Function)

            ' Register EmailUtils
            services.AddSingleton(Of EmailUtils)(Function(sp)
                                                     Return New EmailUtils(
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of XMLManager)()
        )
                                                 End Function)

            ' Register RunTab with proper dependencies
            services.AddSingleton(Of RunTab)(Function(sp)
                                                 Return New RunTab(
            sp.GetRequiredService(Of CredentialManager)(),
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of TaskRunner)(),
            sp.GetRequiredService(Of XMLManager)(),
            sp.GetRequiredService(Of CustomScheduler)()
        )
                                             End Function)

            ' Register SqlTab with proper dependencies
            services.AddSingleton(Of SqlTab)(Function(sp)
                                                 Return New SqlTab(
            sp.GetRequiredService(Of CredentialManager)(),
            sp.GetRequiredService(Of Logger)(),
            sp.GetRequiredService(Of TaskRunner)(),
            sp.GetRequiredService(Of XMLManager)(),
            sp.GetRequiredService(Of PathConfiguration)().LogsPath
        )
                                             End Function)

            ' Register LiteTaskService with proper dependencies
            services.AddSingleton(Of LiteTaskService)(Function(sp)
                                                          Return New LiteTaskService(
            sp.GetRequiredService(Of CustomScheduler)(),
            sp.GetRequiredService(Of CredentialManager)(),
            sp.GetRequiredService(Of XMLManager)(),
            sp.GetRequiredService(Of ToolManager)()
        )
                                                      End Function)
        End Sub

        Private Shared Sub ValidateRequiredServices()
            Dim requiredServices = New Type() {
                GetType(PathConfiguration),
                GetType(XMLManager),
                GetType(Logger),
                GetType(CredentialManager),
                GetType(TranslationManager)
            }

            For Each serviceType In requiredServices
                Try
                    Dim service = _serviceProvider.GetService(serviceType)
                    If service Is Nothing Then
                        Throw New InvalidOperationException($"Required service {serviceType.Name} could not be resolved")
                    End If
                    LogInitialization($"Validated service: {serviceType.Name}")
                Catch ex As Exception
                    Throw New InvalidOperationException($"Failed to validate {serviceType.Name}", ex)
                End Try
            Next
        End Sub

        Private Shared Sub LogInitialization(message As String)
            If Not _enableInitializationLogging Then Return
            Try
                Dim logPath = GetLogPath("initialization.log")
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}")
            Catch
                ' Ignore logging errors
            End Try
        End Sub

        Private Shared Sub LogInitializationError(ex As Exception)
            If Not _enableInitializationLogging Then Return
            Try
                Dim logPath = GetLogPath("initialization_error.log")
                Dim errorMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Initialization Error:{Environment.NewLine}" &
                         $"Message: {ex.Message}{Environment.NewLine}" &
                         $"Stack Trace: {ex.StackTrace}{Environment.NewLine}"
                If ex.InnerException IsNot Nothing Then
                    errorMessage &= $"Inner Exception: {ex.InnerException.Message}{Environment.NewLine}" &
                          $"Inner Stack Trace: {ex.InnerException.StackTrace}{Environment.NewLine}"
                End If
                File.AppendAllText(logPath, errorMessage)
            Catch
                ' Ignore logging errors
            End Try
        End Sub

    End Class
End Namespace