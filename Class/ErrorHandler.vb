Imports LiteTask.LiteTask

Public Class ErrorHandler
    Private ReadOnly _logger As Logger
    Private ReadOnly _emailUtils As EmailUtils

    Public Sub New(logger As Logger, emailUtils As EmailUtils)
        _logger = logger
        _emailUtils = emailUtils
    End Sub

    Public Sub HandleException(ex As Exception, source As String, Optional notifyUser As Boolean = True)
        Try
            ' Log the error
            Dim errorMessage = $"Error in {source}: {ex.Message}"
            _logger.LogError(errorMessage)
            _logger.LogError($"Stack trace: {ex.StackTrace}")

            ' Log inner exception if present
            If ex.InnerException IsNot Nothing Then
                _logger.LogError($"Inner exception: {ex.InnerException.Message}")
                _logger.LogError($"Inner stack trace: {ex.InnerException.StackTrace}")
            End If

            ' Send email notification if enabled
            Try
                _emailUtils.SendEmailReport(
                    $"LiteTask Error: {source}",
                    $"An error occurred in LiteTask:{Environment.NewLine}{Environment.NewLine}" &
                    $"Source: {source}{Environment.NewLine}" &
                    $"Error: {ex.Message}{Environment.NewLine}{Environment.NewLine}" &
                    $"Stack Trace:{Environment.NewLine}{ex.StackTrace}")
            Catch emailEx As Exception
                _logger.LogError($"Failed to send error notification email: {emailEx.Message}")
            End Try

            ' Show message to user if requested
            If notifyUser Then
                MessageBox.Show(
                    $"An error occurred in {source}:{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End If

        Catch criticalEx As Exception
            ' Last resort error handling
            Debug.WriteLine($"Critical error in error handler: {criticalEx.Message}")
            If notifyUser Then
                MessageBox.Show(
                    "A critical error occurred. Please check the application logs.",
                    "Critical Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    Public Sub InitializeErrorHandling()
        ' Set up global exception handlers
        AddHandler Application.ThreadException, AddressOf HandleThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleUnhandledException

        _logger.LogInfo("Global error handling initialized")
    End Sub

    Private Sub HandleThreadException(sender As Object, e As ThreadExceptionEventArgs)
        HandleException(e.Exception, "Unhandled Thread Exception")
    End Sub

    Private Sub HandleUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        HandleException(DirectCast(e.ExceptionObject, Exception), "Unhandled Application Exception")
    End Sub

    Public Function ValidateSettings(settings As Dictionary(Of String, String), requiredFields As String()) As List(Of String)
        Dim errors As New List(Of String)

        For Each field In requiredFields
            If Not settings.ContainsKey(field) OrElse String.IsNullOrWhiteSpace(settings(field)) Then
                errors.Add($"Required setting '{field}' is missing or empty")
            End If
        Next

        Return errors
    End Function
End Class