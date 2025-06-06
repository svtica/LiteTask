Imports System.Collections.Concurrent
Imports System.IO.Compression


Namespace LiteTask
    Public Class Logger
        Implements IDisposable

        Private ReadOnly _logQueue As New ConcurrentQueue(Of LogEntry)
        Private ReadOnly _processorCancellation As New CancellationTokenSource()
        Private _processorTask As Task
        Public _logFile As String
        Public _logFolder As String
        Private _logLevel As LogLevel
        Private _maxLogSize As Long = 10 * 1024 * 1024
        Private _maxLogAge As TimeSpan = TimeSpan.FromDays(30)
        Private _rotationLock As New ReaderWriterLockSlim()
        Private _maxRetries As Integer = 3
        Private _retryDelayMs As Integer = 1000
        Private ReadOnly _lock As New Object()
        Private _disposed As Boolean = False
        Private _xmlManager As XMLManager
        Private Shared _instance As Logger

        Public Enum LogLevel
            [Debug] = 0
            [Info] = 1
            [Warning] = 2
            [Error] = 3
            [Critical] = 4
        End Enum

        Public Sub New(xmlManager As XMLManager, logFilePath As String)
            _logFile = logFilePath
            _xmlManager = xmlManager

            ' Load log settings first
            Dim logSettings = xmlManager.GetLogSettings()
            _logLevel = DirectCast([Enum].Parse(GetType(LogLevel), logSettings("LogLevel")), LogLevel)

            ' Start processor only after level is set
            StartLogProcessor()
            InitializeLogRotation()
        End Sub

        Private Async Function CleanupOldLogsAsync() As Task
            Try
                _rotationLock.EnterWriteLock()

                Dim directory = New DirectoryInfo(_logFolder)
                Dim cutoffDate = DateTime.Now.Subtract(_maxLogAge)

                For Each file In directory.GetFiles($"{Path.GetFileNameWithoutExtension(_logFile)}_*.log*")
                    If file.CreationTime < cutoffDate Then
                        Try
                            file.Delete()
                        Catch ex As Exception
                            Debug.WriteLine($"Error deleting old log {file.Name}: {ex.Message}")
                        End Try
                    End If
                Next
            Finally
                _rotationLock.ExitWriteLock()
            End Try
        End Function

        Private Shared Sub CompressLog(logPath As String)
            Try
                Using originalFile As FileStream = File.OpenRead(logPath)
                    Using compressedFile As FileStream = File.Create($"{logPath}.gz")
                        Using gzipStream As New GZipStream(compressedFile, CompressionLevel.Optimal)
                            originalFile.CopyTo(gzipStream)
                        End Using
                    End Using
                End Using
                File.Delete(logPath)
            Catch ex As Exception
                Debug.WriteLine($"Error compressing log: {ex.Message}")
            End Try
        End Sub

        Private Function CompressLogAsync(logPath As String) As Task(Of Boolean)
            Return Task.Run(Function()
                                For retryCount = 0 To _maxRetries
                                    Try
                                        Using originalFile As FileStream = File.OpenRead(logPath)
                                            Using compressedFile As FileStream = File.Create($"{logPath}.gz")
                                                Using gzipStream As New GZipStream(compressedFile, CompressionLevel.Optimal)
                                                    originalFile.CopyTo(gzipStream)
                                                End Using
                                            End Using
                                        End Using

                                        File.Delete(logPath)
                                        Return True

                                    Catch ex As IOException When retryCount < _maxRetries
                                        Thread.Sleep(_retryDelayMs)
                                    End Try
                                Next
                                Return False
                            End Function)
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    _processorCancellation.Cancel()
                    Try
                        _processorTask?.Wait(TimeSpan.FromSeconds(5))
                    Catch ex As Exception
                        Debug.WriteLine($"Error waiting for log processor to complete: {ex.Message}")
                    End Try
                    _processorCancellation.Dispose()
                End If
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private Sub InitializeLogRotation()
            Task.Run(Async Function()
                         While Not _processorCancellation.Token.IsCancellationRequested
                             Try
                                 Await RotateLogsAsync()
                                 Await CleanupOldLogsAsync()
                                 Await Task.Delay(TimeSpan.FromHours(1), _processorCancellation.Token)
                             Catch ex As OperationCanceledException
                                 ' Normal cancellation, exit gracefully
                                 Return
                             Catch ex As Exception
                                 Debug.WriteLine($"Error in log rotation: {ex.Message}")
                             End Try
                         End While
                     End Function, _processorCancellation.Token)
        End Sub

        Public Event LogEntryAdded As EventHandler(Of LogEntryEventArgs)

        Public Sub LogDebug(message As String, Optional exception As Exception = Nothing)
            QueueLogEntry(LogLevel.Debug, message, exception)
        End Sub

        Public Sub LogInfo(message As String, Optional exception As Exception = Nothing)
            QueueLogEntry(LogLevel.Info, message, exception)
        End Sub


        Public Sub LogError(message As String, Optional exception As Exception = Nothing)
            QueueLogEntry(LogLevel.Error, message, exception)
        End Sub

        Public Sub LogCritical(message As String, Optional exception As Exception = Nothing)
            QueueLogEntry(LogLevel.Critical, message, exception)
        End Sub

        Public Sub LogWarning(message As String, Optional exception As Exception = Nothing)
            QueueLogEntry(LogLevel.Warning, message, exception)
        End Sub

        Private Async Function RotateLogsAsync() As Task
            Try
                _rotationLock.EnterWriteLock()

                If Not File.Exists(_logFile) Then Return

                Dim fileInfo = New FileInfo(_logFile)
                If fileInfo.Length >= _maxLogSize Then
                    Dim timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                    Dim tempPath = Path.GetTempFileName()
                    Dim archivePath = Path.Combine(_logFolder, $"{Path.GetFileNameWithoutExtension(_logFile)}_{timestamp}{Path.GetExtension(_logFile)}")

                    ' Create atomic rotation
                    Try
                        ' Copy current log to temp file
                        File.Copy(_logFile, tempPath, True)

                        ' Truncate current log
                        File.WriteAllText(_logFile, String.Empty)

                        ' Move temp to archive
                        File.Move(tempPath, archivePath)

                        ' Compress archive
                        Await CompressLogAsync(archivePath)
                    Finally
                        ' Cleanup temp file if exists
                        If File.Exists(tempPath) Then
                            File.Delete(tempPath)
                        End If
                    End Try
                End If
            Finally
                _rotationLock.ExitWriteLock()
            End Try
        End Function

        Public Sub SetLogFolder(folder As String)
            _logFolder = folder
            _logFile = Path.Combine(_logFolder, Path.GetFileName(_logFile))
        End Sub

        Public Sub SetLogLevel(level As LogLevel)
            _logLevel = level
        End Sub

        Private Sub StartLogProcessor()
            _processorTask = Task.Run(Async Function()
                                          While Not _processorCancellation.Token.IsCancellationRequested
                                              Try
                                                  Dim entry As LogEntry = Nothing
                                                  If _logQueue.TryDequeue(entry) Then
                                                      Await WriteLogEntryAsync(entry)
                                                  Else
                                                      Await Task.Delay(100, _processorCancellation.Token)
                                                  End If
                                              Catch ex As Exception
                                                  ' Last resort error handling
                                                  Debug.WriteLine($"Error processing log entry: {ex.Message}")
                                              End Try
                                          End While
                                      End Function, _processorCancellation.Token)
        End Sub

        Private Sub QueueLogEntry(level As LogLevel, message As String, exception As Exception)
            ' Adjust log levels for routine operations
            If message.StartsWith("Getting all tasks") OrElse
       message.StartsWith("Checking tasks at") OrElse
       message.StartsWith("Checking task:") OrElse
       message.StartsWith("Processing action:") OrElse
       message.StartsWith("Action type:") OrElse
       message.StartsWith("Action target:") Then
                level = LogLevel.Debug
            End If

            ' Only queue if message level meets or exceeds configured log level
            Dim messageLevel = CInt(level)
            Dim currentLogLevel = CInt(_logLevel)

            If messageLevel >= currentLogLevel Then
                ' Parse process output messages to adjust log level
                If message.StartsWith("Process output:") OrElse
          message.StartsWith("Process debug:") Then
                    level = LogLevel.Info
                End If

                _logQueue.Enqueue(New LogEntry With {
           .Level = level,
           .Message = message,
           .Exception = exception,
           .Timestamp = DateTime.Now
       })
            End If
        End Sub

        Private Async Function WriteLogEntryAsync(entry As LogEntry) As Task
            Dim logEntry = $"{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{entry.Level}] {entry.Message}"
            If entry.Exception IsNot Nothing Then
                logEntry &= $"{Environment.NewLine}Exception: {entry.Exception.Message}"
                logEntry &= $"{Environment.NewLine}StackTrace: {entry.Exception.StackTrace}"
            End If
            logEntry &= Environment.NewLine

            Try
                SyncLock _lock
                    Using writer As New StreamWriter(_logFile, True, Encoding.UTF8)
                        ' Use sync write inside lock
                        writer.Write(logEntry)
                    End Using
                End SyncLock

                ' Write to debug output for development
                Debug.WriteLine(logEntry)

                ' Notify UI if needed
                RaiseEvent LogEntryAdded(Me, New LogEntryEventArgs(entry))

            Catch ex As Exception
                Debug.WriteLine($"Error writing log entry: {ex.Message}")
            End Try
        End Function

    End Class

    Public Class LogEntry
        Public Property Timestamp As DateTime
        Public Property Level As Logger.LogLevel
        Public Property Message As String
        Public Property Exception As Exception
    End Class

    Public Class LogEntryEventArgs
        Inherits EventArgs

        Public Property LogEntry As LogEntry

        Public Sub New(entry As LogEntry)
            LogEntry = entry
        End Sub
    End Class

End Namespace

