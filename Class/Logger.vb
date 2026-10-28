Imports System.Collections.Concurrent
Imports System.IO.Compression
Imports System.Diagnostics
Imports System.IO


Namespace LiteTask
    Public Class Logger
        Implements IDisposable

        Private ReadOnly _logQueue As New ConcurrentQueue(Of LogEntry)
        Private ReadOnly _processorCancellation As New CancellationTokenSource()
        Private _processorTask As Task
        Private _rotationTask As Task
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
        Private Const MAX_QUEUE_SIZE As Integer = 10000
        ' Maximum number of queued entries written with a single file open.
        Private Const MAX_WRITE_BATCH As Integer = 256

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

            ' Default the log folder to the log file's directory. Until now it was
            ' only set through SetLogFolder (GUI code paths), so under the Windows
            ' service _logFolder stayed Nothing and the rotation/cleanup task threw
            ' on every iteration.
            Dim folder = Path.GetDirectoryName(logFilePath)
            If String.IsNullOrEmpty(folder) Then
                folder = AppDomain.CurrentDomain.BaseDirectory
            End If
            _logFolder = folder

            ' Load log settings first
            Dim logSettings = xmlManager.GetLogSettings()
            _logLevel = DirectCast([Enum].Parse(GetType(LogLevel), logSettings("LogLevel")), LogLevel)
            ApplyRotationSettings(logSettings)

            ' Start processor only after level is set
            StartLogProcessor()
            InitializeLogRotation()
        End Sub

        ''' <summary>
        ''' Applies Logging/MaxLogSize (MB) and Logging/LogRetentionDays from the
        ''' settings file. These values are exposed in the Options form but were
        ''' never read by the logger, which always used 10 MB / 30 days.
        ''' Invalid or missing values keep the built-in defaults.
        ''' </summary>
        Private Sub ApplyRotationSettings(logSettings As Dictionary(Of String, String))
            If logSettings Is Nothing Then Return

            Dim maxSizeMb As Integer
            If logSettings.ContainsKey("MaxLogSize") AndAlso
               Integer.TryParse(logSettings("MaxLogSize"), maxSizeMb) AndAlso maxSizeMb >= 1 Then
                _maxLogSize = CLng(maxSizeMb) * 1024L * 1024L
            End If

            Dim retentionDays As Integer
            If logSettings.ContainsKey("LogRetentionDays") AndAlso
               Integer.TryParse(logSettings("LogRetentionDays"), retentionDays) AndAlso retentionDays >= 1 Then
                _maxLogAge = TimeSpan.FromDays(retentionDays)
            End If
        End Sub

        Private Function CleanupOldLogsAsync() As Task
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
            Return Task.CompletedTask
        End Function

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

        'Enhanced cleanup method that handles both custom temp directory and system temp directory
        Public Sub CleanupAllTempFiles()
            Try
                ' Clean up custom temp directory
                Dim customTempDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteTaskData", "temp")
                CleanupTempDirectory(customTempDir, "*.tmp")
                
                ' Clean up any orphaned files in system temp directory (legacy cleanup)
                Dim systemTempDir = Path.GetTempPath()
                CleanupTempDirectory(systemTempDir, "tmp*.tmp") ' System temp files have specific naming pattern
                CleanupTempDirectory(systemTempDir, "log_rotation_*.tmp") ' Our specific pattern
                CleanupTempDirectory(systemTempDir, "LiteTask_*.log") ' PowerShell module install logs
                
            Catch ex As Exception
                Debug.WriteLine($"Error during comprehensive temp file cleanup: {ex.Message}")
            End Try
        End Sub

        Private Sub CleanupTempDirectory(directory As String, searchPattern As String)
            Try
                If Not System.IO.Directory.Exists(directory) Then Return
                
                Dim cutoffTime = DateTime.Now.AddHours(-1) ' Files older than 1 hour
                Dim tempFiles = System.IO.Directory.GetFiles(directory, searchPattern)
                
                For Each file In tempFiles
                    Try
                        Dim fileInfo = New System.IO.FileInfo(file)
                        If fileInfo.LastWriteTime < cutoffTime Then
                            ' Additional check: ensure file is not in use
                            If Not IsFileLocked(file) Then
                                System.IO.File.Delete(file)
                                Debug.WriteLine($"Cleaned up old temp file: {file}")
                            End If
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Failed to delete temp file {file}: {ex.Message}")
                    End Try
                Next
                
            Catch ex As Exception
                Debug.WriteLine($"Error cleaning temp directory {directory}: {ex.Message}")
            End Try
        End Sub

        Private Function IsFileLocked(filePath As String) As Boolean
            Try
                Using fs = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None)
                    Return False
                End Using
            Catch
                Return True
            End Try
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
                    Try
                        _rotationTask?.Wait(TimeSpan.FromSeconds(5))
                    Catch ex As Exception
                        Debug.WriteLine($"Error waiting for log rotation to complete: {ex.Message}")
                    End Try
                    _processorCancellation.Dispose()
                    _rotationLock?.Dispose()
                End If
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private Sub InitializeLogRotation()
            _rotationTask = Task.Run(Async Function()
                         While Not _processorCancellation.Token.IsCancellationRequested
                             Try
                                 Await RotateLogsAsync()
                                 Await CleanupOldLogsAsync()
                             Catch ex As OperationCanceledException
                                 ' Normal cancellation, exit gracefully
                                 Return
                             Catch ex As Exception
                                 Debug.WriteLine($"Error in log rotation: {ex.Message}")
                             End Try

                             ' The delay lives outside the Try above: an exception in
                             ' rotation/cleanup must not turn this loop into a busy spin.
                             Try
                                 Await Task.Delay(TimeSpan.FromHours(1), _processorCancellation.Token)
                             Catch ex As OperationCanceledException
                                 Return
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
            Dim archivePath As String = Nothing

            ' ReaderWriterLockSlim is thread-affine: the thread that enters the write
            ' lock must be the one that exits it. The previous implementation awaited
            ' the compression inside the lock, resumed on another pool thread, and the
            ' release then threw - leaving the lock held forever and silently stopping
            ' rotation after the first archive. Only the synchronous copy/truncate/move
            ' is done under the lock; compression happens after it is released.
            _rotationLock.EnterWriteLock()
            Try
                If Not File.Exists(_logFile) Then Return

                Dim fileInfo = New FileInfo(_logFile)
                If fileInfo.Length >= _maxLogSize Then
                    Dim timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                    
                    ' Create a custom temp directory for log rotation
                    Dim customTempDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LiteTaskData", "temp")
                    Directory.CreateDirectory(customTempDir)  ' Ensure directory exists
                    Dim tempPath = Path.Combine(customTempDir, $"log_rotation_{Guid.NewGuid()}.tmp")
                    
                    Dim candidateArchivePath = Path.Combine(_logFolder, $"{Path.GetFileNameWithoutExtension(_logFile)}_{timestamp}{Path.GetExtension(_logFile)}")

                    ' Create atomic rotation with better error handling
                    Try
                        ' Copy current log to temp file
                        File.Copy(_logFile, tempPath, True)

                        ' Truncate current log
                        File.WriteAllText(_logFile, String.Empty)

                        ' Move temp to archive
                        File.Move(tempPath, candidateArchivePath)
                        archivePath = candidateArchivePath
                        
                    Catch ex As Exception
                        '_logger?.LogError($"Error during log rotation: {ex.Message}")
                        ' If rotation failed, restore from temp file if it exists
                        If File.Exists(tempPath) AndAlso Not File.Exists(_logFile) Then
                            Try
                                File.Move(tempPath, _logFile)
                            Catch restoreEx As Exception
                                Debug.WriteLine($"Failed to restore log file: {restoreEx.Message}")
                            End Try
                        End If
                        Throw
                    Finally
                        ' Cleanup temp file with retry logic
                        If File.Exists(tempPath) Then
                            For retry = 1 To 3
                                Try
                                    File.Delete(tempPath)
                                    Exit For
                                Catch ex As Exception
                                    If retry = 3 Then
                                        Debug.WriteLine($"Failed to delete temp file after {retry} attempts: {tempPath} - {ex.Message}")
                                    Else
                                        Thread.Sleep(100) ' Wait 100ms before retry
                                    End If
                                End Try
                            Next
                        End If
                    End Try
                End If
            Finally
                _rotationLock.ExitWriteLock()
            End Try

            ' Compress the archive outside the lock (see comment above)
            If archivePath IsNot Nothing AndAlso File.Exists(archivePath) Then
                Await CompressLogAsync(archivePath)
            End If
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
                                          Dim emptyQueueCount = 0
                                          Dim batch As New List(Of LogEntry)(MAX_WRITE_BATCH)
                                          While Not _processorCancellation.Token.IsCancellationRequested
                                              Try
                                                  ' Drain up to MAX_WRITE_BATCH entries and write them with a
                                                  ' single file open instead of one open/close per line.
                                                  batch.Clear()
                                                  Dim entry As LogEntry = Nothing
                                                  While batch.Count < MAX_WRITE_BATCH AndAlso _logQueue.TryDequeue(entry)
                                                      batch.Add(entry)
                                                  End While

                                                  If batch.Count > 0 Then
                                                      WriteLogEntries(batch)
                                                      emptyQueueCount = 0 ' Reset counter on successful dequeue
                                                  Else
                                                      emptyQueueCount += 1
                                                      ' Adaptive delay: short delay initially, longer for extended idle periods
                                                      Dim delayMs = Math.Min(100 + (emptyQueueCount * 10), 1000)
                                                      Await Task.Delay(delayMs, _processorCancellation.Token)
                                                  End If
                                              Catch ex As OperationCanceledException
                                                  ' Normal shutdown, exit gracefully
                                                  Exit While
                                              Catch ex As Exception
                                                  ' Last resort error handling
                                                  Debug.WriteLine($"Error processing log entry: {ex.Message}")
                                              End Try
                                          End While

                                          ' Best-effort flush so the last entries before shutdown are not lost
                                          Try
                                              batch.Clear()
                                              Dim remaining As LogEntry = Nothing
                                              While batch.Count < MAX_WRITE_BATCH AndAlso _logQueue.TryDequeue(remaining)
                                                  batch.Add(remaining)
                                              End While
                                              If batch.Count > 0 Then WriteLogEntries(batch)
                                          Catch ex As Exception
                                              Debug.WriteLine($"Error flushing log entries: {ex.Message}")
                                          End Try
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

                ' Drop entries when queue is full to prevent unbounded memory growth
                If _logQueue.Count < MAX_QUEUE_SIZE Then
                    _logQueue.Enqueue(New LogEntry With {
               .Level = level,
               .Message = message,
               .Exception = exception,
               .Timestamp = DateTime.Now
           })
                End If
            End If
        End Sub

        Private Shared Sub AppendFormattedEntry(builder As StringBuilder, entry As LogEntry)
            builder.Append($"{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{entry.Level}] {entry.Message}")

            If entry.Exception IsNot Nothing Then
                builder.AppendLine()
                builder.Append($"Exception: {entry.Exception.Message}")
                builder.AppendLine()
                builder.Append($"StackTrace: {entry.Exception.StackTrace}")
            End If
            builder.AppendLine()
        End Sub

        Private Sub WriteLogEntries(entries As List(Of LogEntry))
            If entries Is Nothing OrElse entries.Count = 0 Then Return

            Dim logEntryBuilder As New StringBuilder()
            For Each entry In entries
                AppendFormattedEntry(logEntryBuilder, entry)
            Next
            Dim logText = logEntryBuilder.ToString()

            Try
                ' Take the rotation lock as a reader so a rotation (writer) can never
                ' truncate the file in the middle of an append. This method is fully
                ' synchronous, so the lock is released on the thread that acquired it.
                Dim rotationLockHeld = False
                Try
                    rotationLockHeld = _rotationLock.TryEnterReadLock(TimeSpan.FromSeconds(10))
                    SyncLock _lock
                        Using writer As New StreamWriter(_logFile, True, Encoding.UTF8)
                            ' Use sync write inside lock
                            writer.Write(logText)
                        End Using
                    End SyncLock
                Finally
                    If rotationLockHeld Then _rotationLock.ExitReadLock()
                End Try

                ' Write to debug output for development (only if debugger attached)
                If Debugger.IsAttached Then
                    Debug.WriteLine(logText)
                End If

                ' Notify UI if needed
                For Each entry In entries
                    RaiseEvent LogEntryAdded(Me, New LogEntryEventArgs(entry))
                Next

            Catch ex As Exception
                Debug.WriteLine($"Error writing log entries: {ex.Message}")
            End Try
        End Sub

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

