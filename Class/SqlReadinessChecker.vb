Imports System.Data.SqlClient

Namespace LiteTask
    ''' <summary>
    ''' Lightweight readiness gate for SQL Server, used by the scheduler to defer
    ''' SQL-dependent tasks until the database engine accepts connections (e.g.
    ''' right after a server reboot, when the service starts before SQL Server).
    ''' Probes are throttled with exponential backoff so an extended outage does
    ''' not flood the server or the logs, and a successful probe is cached briefly
    ''' so the scheduler tick does not open a connection every 30 seconds.
    ''' </summary>
    Public Class SqlReadinessChecker
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _syncRoot As New Object()

        Private _lastResult As Boolean
        Private _lastSuccessTime As DateTime = DateTime.MinValue
        Private _nextProbeTime As DateTime = DateTime.MinValue
        Private _consecutiveFailures As Integer

        Private Const PROBE_CONNECT_TIMEOUT_SECONDS As Integer = 5
        Private Const SUCCESS_CACHE_SECONDS As Integer = 300
        Private Const INITIAL_BACKOFF_SECONDS As Integer = 15
        Private Const MAX_BACKOFF_SECONDS As Integer = 600

        Public Sub New(xmlManager As XMLManager, logger As Logger)
            _xmlManager = xmlManager
            _logger = logger
        End Sub

        ''' <summary>
        ''' Server to probe: Reliability/SqlReadinessServer if set, otherwise the
        ''' default server from SqlConfiguration. Empty means nothing to probe.
        ''' </summary>
        Private Function GetProbeServer() As String
            Try
                Dim server = _xmlManager.ReadValue("Reliability", "SqlReadinessServer", "")
                If String.IsNullOrWhiteSpace(server) Then
                    server = _xmlManager.GetSqlConfiguration()("DefaultServer")
                End If
                Return server
            Catch ex As Exception
                _logger?.LogError($"[SqlReadiness] Error reading probe server configuration: {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Returns True when SQL Server accepts connections. While unavailable,
        ''' actual probes only happen when the current backoff window has elapsed;
        ''' calls in between return the cached negative result immediately.
        ''' </summary>
        Public Function IsReady() As Boolean
            SyncLock _syncRoot
                Dim now = DateTime.Now
                Dim server = GetProbeServer()

                ' No server configured: nothing to probe, never block tasks
                If String.IsNullOrWhiteSpace(server) Then
                    Return True
                End If

                If _lastResult AndAlso (now - _lastSuccessTime).TotalSeconds < SUCCESS_CACHE_SECONDS Then
                    Return True
                End If

                If Not _lastResult AndAlso now < _nextProbeTime Then
                    Return False
                End If

                Return Probe(server, now)
            End SyncLock
        End Function

        ' Connectivity-level SqlException numbers (server unreachable). Any other
        ' SqlException (e.g. 18456 login failed, 4060 cannot open database) means
        ' the engine answered, so the gate must NOT keep blocking tasks.
        Private Shared ReadOnly NetworkErrorNumbers As Integer() = {-1, 2, 40, 53, 258, 1231, 10060, 10061, 11001}

        Private Function Probe(server As String, now As DateTime) As Boolean
            Try
                _logger?.LogInfo($"[SqlReadiness] Probing SQL Server '{server}' (attempt {_consecutiveFailures + 1})")
                Using connection As New SqlConnection($"Server={server};Database=master;Integrated Security=True;TrustServerCertificate=True;Connect Timeout={PROBE_CONNECT_TIMEOUT_SECONDS}")
                    connection.Open()
                End Using
                Return MarkReady(server, now)

            Catch ex As SqlException When Not NetworkErrorNumbers.Contains(ex.Number)
                ' The server responded but refused this probe (login/database error):
                ' the engine is up, which is all the readiness gate needs to know.
                _logger?.LogInfo($"[SqlReadiness] SQL Server '{server}' answered with error {ex.Number} ({ex.Message.Trim()}); treating engine as available")
                Return MarkReady(server, now)

            Catch ex As Exception
                _consecutiveFailures += 1
                Dim backoffSeconds = Math.Min(INITIAL_BACKOFF_SECONDS * Math.Pow(2, Math.Min(_consecutiveFailures - 1, 10)), MAX_BACKOFF_SECONDS)
                _nextProbeTime = now.AddSeconds(backoffSeconds)
                _lastResult = False
                _logger?.LogWarning($"[SqlReadiness] SQL Server '{server}' not ready (probe {_consecutiveFailures} failed: {ex.Message}). Next probe in {backoffSeconds:F0}s")
                Return False
            End Try
        End Function

        Private Function MarkReady(server As String, now As DateTime) As Boolean
            If _consecutiveFailures > 0 Then
                _logger?.LogInfo($"[SqlReadiness] SQL Server '{server}' is available again after {_consecutiveFailures} failed probe(s)")
            End If
            _lastResult = True
            _lastSuccessTime = now
            _consecutiveFailures = 0
            Return True
        End Function
    End Class
End Namespace
