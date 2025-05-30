Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Text.RegularExpressions

Namespace LiteTask
    Public Class SqlTab
        ' Values
        Private _connectionStringBase As String
        Private ReadOnly _logPath As String
        Private _tabPageInitialized As Boolean = False
        ' Components
        Private ReadOnly _credentialManager As CredentialManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _taskRunner As TaskRunner
        Private ReadOnly _powerShellManager As RunTab
        Private ReadOnly _xmlManager As XMLManager
        Private ReadOnly _emailUtils As EmailUtils
        Private ReadOnly _liteRunConfig As LiteRunConfig
        Private ReadOnly _translationManager As TranslationManager
        ' Controls
        Private _queryTextBox As TextBox
        Private _outputTextBox As TextBox
        Private _runButton As Button
        Private _queryTypeComboBox As ComboBox
        Private _credentialComboBox As ComboBox
        Public WithEvents _serverTextBox As TextBox
        Public _databaseComboBox As ComboBox
        Private _queryLabel As Label
        Private _outputLabel As Label
        Private _serverLabel As Label
        Private _databaseLabel As Label
        Private _queryTypeLabel As Label
        Private _credentialLabel As Label
        Private _tabPage As TabPage
        Private _tableLayoutPanel As TableLayoutPanel
        Private WithEvents _connectButton As Button
        Private _objectListComboBox As ComboBox
        Private _objectListLabel As Label
        Private _selectAllButton As Button
        Private _returnLastEntryButton As Button

        Public Sub New(credentialManager As CredentialManager, logger As Logger, taskRunner As TaskRunner, xmlManager As XMLManager, logPath As String)
            If credentialManager Is Nothing Then
                Throw New ArgumentNullException(NameOf(credentialManager))
            End If
            If logger Is Nothing Then
                Throw New ArgumentNullException(NameOf(logger))
            End If
            If taskRunner Is Nothing Then
                Throw New ArgumentNullException(NameOf(taskRunner))
            End If
            If xmlManager Is Nothing Then
                Throw New ArgumentNullException(NameOf(xmlManager))
            End If
            If String.IsNullOrEmpty(logPath) Then
                Throw New ArgumentNullException(NameOf(logPath))
            End If

            _credentialManager = credentialManager
            _logger = logger
            _taskRunner = taskRunner
            _xmlManager = xmlManager
            _logPath = logPath
            _emailUtils = New EmailUtils(logger, xmlManager)
            '_liteRunConfig = New LiteRunConfig(xmlManager.GetLiteRunDefaults())
            _translationManager = TranslationManager.Initialize(logger, xmlManager)
        End Sub

        Private Sub AddHandlers()
            AddHandler _runButton.Click, AddressOf RunButton_Click
            AddHandler _selectAllButton.Click, AddressOf SelectAllButton_Click
            AddHandler _returnLastEntryButton.Click, AddressOf ReturnLastEntryButton_Click
            AddHandler _queryTypeComboBox.SelectedIndexChanged, AddressOf QueryTypeComboBox_SelectedIndexChanged
            AddHandler _objectListComboBox.SelectedIndexChanged, AddressOf ObjectListComboBox_SelectedIndexChanged
            AddHandler _connectButton.Click, AddressOf ConnectButton_Click
        End Sub

        Private Sub ConnectButton_Click(sender As Object, e As EventArgs) Handles _connectButton.Click
            PopulateDatabaseComboBox()
        End Sub

        Private Sub DatabaseComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _databaseComboBox.SelectedIndex > 0 Then
                PopulateObjectListComboBox()
            End If
        End Sub

        Private Function DataTableToString(dt As DataTable) As String
            Dim result As New StringBuilder()

            ' Add headers
            result.AppendLine(String.Join(vbTab, dt.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName)))

            ' Add rows
            For Each row As DataRow In dt.Rows
                result.AppendLine(String.Join(vbTab, row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), "NULL"))))
            Next

            Return result.ToString()
        End Function

        Private Async Function ExecuteAndDisplayQuery(selectAll As Boolean, lastEntry As Boolean) As Task
            Try
                If ValidateInputs() = False Then Return

                Dim query = _queryTextBox.Text
                If Not query.Trim().ToUpper().StartsWith("SELECT") AndAlso (selectAll OrElse lastEntry) Then
                    MessageBox.Show("This option is only available for SELECT queries.", "Invalid Query", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim sqlConfig = _xmlManager.GetSqlConfiguration()
                Dim server = If(Not String.IsNullOrWhiteSpace(_serverTextBox?.Text), _serverTextBox.Text, sqlConfig("DefaultServer"))
                Dim database = If(_databaseComboBox?.SelectedItem IsNot Nothing AndAlso
                                 _databaseComboBox.SelectedItem.ToString() <> "(Select a database)",
                                 _databaseComboBox.SelectedItem.ToString(),
                                 sqlConfig("DefaultDatabase"))

                Dim selectedCredential = If(_credentialComboBox?.SelectedItem?.ToString(), "(None)")

                _logger.LogInfo($"Executing query on {server}.{database}, credential: {selectedCredential}")

                ' Execute the query
                Dim result = Await ExecuteQueryAsync(query, selectedCredential, server, database)

                If result.Success Then
                    If result.Data IsNot Nothing AndAlso result.Data.Rows.Count > 0 Then
                        Dim displayData As DataTable

                        If lastEntry Then
                            ' Create a new DataTable with just the last row
                            displayData = result.Data.Clone()
                            displayData.Rows.Add(result.Data.Rows(result.Data.Rows.Count - 1).ItemArray)
                            _logger.LogInfo("Retrieved last entry successfully")
                        Else
                            displayData = result.Data
                        End If

                        _outputTextBox.Text = DataTableToString(displayData)
                        _logger.LogInfo($"Query executed successfully. Rows returned: {displayData.Rows.Count}")

                        If selectAll Then
                            ' Offer to save as CSV
                            If MessageBox.Show("Would you like to save the results as a CSV file?", "Save Results", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                Await SaveResultsAsCsv(result.Data) ' Save all data, not just displayed data
                            End If
                        End If
                    Else
                        _outputTextBox.Text = "Query executed successfully with no output"
                        _logger.LogInfo("Query executed successfully with no output")
                    End If
                Else
                    _outputTextBox.Text = result.Message
                    _logger.LogError($"Query execution failed: {result.Message}")

                    ' Send error notification
                    Try
                        Dim notificationManager = ApplicationContainer.GetService(Of NotificationManager)()
                        notificationManager?.QueueNotification(
                            "LiteTask Error",
                            $"Error executing query on {server}.{database}: {result.Message}",
                            NotificationManager.NotificationPriority.High)
                    Catch notifyEx As Exception
                        _logger.LogError($"Failed to send error notification: {notifyEx.Message}")
                    End Try
                End If

            Catch ex As Exception
                _logger.LogError($"Error in ExecuteAndDisplayQuery: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                _outputTextBox.Text = $"Error executing query: {ex.Message}"
            End Try
        End Function

        Private Async Function ExecuteDirectSqlQuery(query As String, server As String, database As String) As Task(Of SqlExecutionResult)
            Using connection As New SqlConnection($"Server={server};Database={database};Integrated Security=True;TrustServerCertificate=True")
                Await connection.OpenAsync()

                Using command As New SqlCommand(query, connection)
                    command.CommandTimeout = 300 ' 5 minutes
                    Using reader = Await command.ExecuteReaderAsync()
                        Dim dataTable As New DataTable()
                        dataTable.Load(reader)
                        Return New SqlExecutionResult With {
                    .Success = True,
                    .Data = dataTable,
                    .RowsAffected = dataTable.Rows.Count,
                    .Message = DataTableToString(dataTable)
                }
                    End Using
                End Using
            End Using
        End Function

        Public Async Function ExecuteQueryAsync(query As String, credentialTarget As String, server As String, database As String) As Task(Of SqlExecutionResult)
            Try
                ' Check if this is a stored procedure
                If query.Trim().ToUpper().StartsWith("EXEC ") OrElse query.Trim().ToUpper().StartsWith("EXECUTE ") Then
                    Dim spMatch = Regex.Match(query, "EXEC(?:UTE)?\s+(?:@\w+\s*=\s*)?(?:\[?dbo\]?\.)?\[?(\w+)\]?", RegexOptions.IgnoreCase)
                    If spMatch.Success Then
                        Dim spName = spMatch.Groups(1).Value
                        _logger.LogInfo($"Detected stored procedure: {spName}")

                        ' Get credential if specified
                        Dim credential As CredentialInfo = Nothing
                        If Not String.IsNullOrEmpty(credentialTarget) AndAlso credentialTarget <> "(None)" Then
                            credential = _credentialManager.GetCredential(credentialTarget, "Windows Vault")
                        End If

                        ' Execute using sqlcmd
                        Dim result = Await _taskRunner.ExecuteStoredProcedureWithSqlCmd(spName, server, database, Nothing, credential)
                        Return New SqlExecutionResult With {
                    .Success = Not result.StartsWith("Error:", StringComparison.OrdinalIgnoreCase),
                    .Message = result
                }
                    End If
                End If

                ' For non-stored procedure queries, use direct SQL execution
                Return Await ExecuteDirectSqlQuery(query, server, database)
            Catch ex As Exception
                _logger.LogError($"Error executing query: {ex.Message}")
                Return New SqlExecutionResult With {
            .Success = False,
            .Message = $"Error: {ex.Message}"
        }
            End Try
        End Function

        Private Function ExtractStoredProcedureName(query As String) As String
            Dim match = Regex.Match(query, "EXEC(?:UTE)?\s+(?:@\w+\s*=\s*)?(?:\[?dbo\]?\.)?\[?(\w+)\]?", RegexOptions.IgnoreCase)
            Return If(match.Success, match.Groups(1).Value, String.Empty)
        End Function

        Public Function GetTabPage() As TabPage
            If Not _tabPageInitialized Then
                Try
                    _logger?.LogInfo("Initializing SqlTab components")
                    InitializeComponent()
                    PopulateCredentialComboBox()
                    PopulateDatabaseComboBox()
                    PopulateQueryTypeComboBox()
                    _tabPageInitialized = True
                    _logger?.LogInfo("SqlTab components initialized successfully")
                Catch ex As Exception
                    _logger?.LogError($"Error initializing SqlTab components: {ex.Message}")
                    _logger?.LogError($"StackTrace: {ex.StackTrace}")
                    Throw
                End Try
            End If
            Return _tabPage
        End Function

        Private Sub InitializeComponent()
            Try
                _logger?.LogInfo("Starting InitializeComponent for SqlTab")

                ' Initialize TabPage
                _tabPage = New TabPage("SQL")

                ' Initialize TableLayoutPanel
                _tableLayoutPanel = New TableLayoutPanel With {
            .Dock = DockStyle.Fill
            }

                ' Configure TableLayoutPanel
                _tableLayoutPanel.ColumnStyles.Clear()
                _tableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20)) ' Label column
                _tableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60)) ' Main content column
                _tableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20)) ' Extra column

                _tableLayoutPanel.RowStyles.Clear()
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Server
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Database
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 40))   ' Query
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 40))   ' Output
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Query Type
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Object List
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Credential
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 40))  ' Select All/Return Last
                _tableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 40))  ' Run Button
                _tableLayoutPanel.ColumnCount = 3
                _tableLayoutPanel.RowCount = 8

                ' Initialize controls
                _serverTextBox = New TextBox With {
            .Dock = DockStyle.Fill
        }

                _connectButton = New Button With {
            .Text = "Connect",
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .MinimumSize = New Size(80, 25),
            .MaximumSize = New Size(120, 25)
        }

                _databaseComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _queryTypeComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _objectListComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _queryTextBox = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill
        }

                _outputTextBox = New TextBox With {
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .ReadOnly = True
        }

                _credentialComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                ' Initialize Labels
                _serverLabel = New Label With {
            .Text = "Server:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _databaseLabel = New Label With {
            .Text = "Database:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _queryLabel = New Label With {
            .Text = "Query:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _outputLabel = New Label With {
            .Text = "Output:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _queryTypeLabel = New Label With {
            .Text = "Query Type:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _queryTypeComboBox = New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

                _objectListLabel = New Label With {
            .Text = "Object:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                _credentialLabel = New Label With {
            .Text = "Credential:",
            .Dock = DockStyle.Fill,
            .AutoSize = True
        }

                ' Initialize buttons with specific sizes
                _selectAllButton = New Button With {
            .Text = "Select All (Save as CSV)",
            .Dock = DockStyle.None,
            .AutoSize = True,
            .MinimumSize = New Size(120, 25),
            .MaximumSize = New Size(200, 25)
        }

                _returnLastEntryButton = New Button With {
            .Text = "Return Last Entry",
            .Dock = DockStyle.None,
            .AutoSize = True,
            .MinimumSize = New Size(120, 25),
            .MaximumSize = New Size(200, 25)
        }

                _runButton = New Button With {
            .Text = "Run Query",
            .Dock = DockStyle.None,
            .AutoSize = True,
            .MinimumSize = New Size(120, 25),
            .MaximumSize = New Size(200, 25)
        }
                ' Add translations with null checks
                Try
                    If _translationManager IsNot Nothing Then
                        _logger?.LogInfo("Applying translations to SqlTab controls")
                        _tabPage.Text = _translationManager.GetTranslation("SQLTab.Text")
                        _serverLabel.Text = _translationManager.GetTranslation("SQLTab.ServerLabel")
                        _databaseLabel.Text = _translationManager.GetTranslation("SQLTab.DatabaseLabel")
                        _queryLabel.Text = _translationManager.GetTranslation("SQLTab.QueryLabel")
                        _outputLabel.Text = _translationManager.GetTranslation("SQLTab.OutputLabel")
                        _queryTypeLabel.Text = _translationManager.GetTranslation("SQLTab.QueryTypeLabel.Text")
                        _objectListLabel.Text = _translationManager.GetTranslation("SQLTab.ObjectListLabel.Text")
                        _credentialLabel.Text = _translationManager.GetTranslation("SQLTab.CredentialLabel")
                        _connectButton.Text = _translationManager.GetTranslation("SQLTab.Connect")
                        _selectAllButton.Text = _translationManager.GetTranslation("SQLTab.SelectAll")
                        _returnLastEntryButton.Text = _translationManager.GetTranslation("SQLTab.LastEntry")
                        _runButton.Text = _translationManager.GetTranslation("SQLTab.RunButton.Text")
                    Else
                        _logger?.LogWarning("TranslationManager is not initialized, using default texts")
                    End If
                Catch translationEx As Exception
                    _logger?.LogWarning($"Error applying translations: {translationEx.Message}")
                    ' Continue with default texts if translation fails
                End Try

                ' Create button panel for better layout
                Dim buttonPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False,
            .AutoSize = True
        }

                ' Add buttons to button panel
                buttonPanel.Controls.Add(_returnLastEntryButton)
                buttonPanel.Controls.Add(New Panel With {.Width = 10}) ' Spacer
                buttonPanel.Controls.Add(_selectAllButton)
                buttonPanel.Controls.Add(New Panel With {.Width = 10}) ' Spacer
                buttonPanel.Controls.Add(_runButton)


                ' Add controls to TableLayoutPanel
                _tableLayoutPanel.Controls.Add(_serverLabel, 0, 0)
                _tableLayoutPanel.Controls.Add(_serverTextBox, 1, 0)
                _tableLayoutPanel.Controls.Add(_connectButton, 2, 0)
                _tableLayoutPanel.Controls.Add(_databaseLabel, 0, 1)
                _tableLayoutPanel.Controls.Add(_databaseComboBox, 1, 1)
                _tableLayoutPanel.Controls.Add(_queryLabel, 0, 2)
                _tableLayoutPanel.Controls.Add(_queryTextBox, 1, 2)
                _tableLayoutPanel.Controls.Add(_outputLabel, 0, 3)
                _tableLayoutPanel.Controls.Add(_outputTextBox, 1, 3)
                _tableLayoutPanel.Controls.Add(_queryTypeLabel, 0, 4)
                _tableLayoutPanel.Controls.Add(_queryTypeComboBox, 1, 4)
                _tableLayoutPanel.Controls.Add(_objectListLabel, 0, 5)
                _tableLayoutPanel.Controls.Add(_objectListComboBox, 1, 5)
                _tableLayoutPanel.Controls.Add(_credentialLabel, 0, 6)
                _tableLayoutPanel.Controls.Add(_credentialComboBox, 1, 6)

                _tableLayoutPanel.Controls.Add(buttonPanel, 1, 7)

                ' Add TableLayoutPanel to TabPage
                _tabPage.Controls.Add(_tableLayoutPanel)
                AddHandlers()
                _logger?.LogInfo("InitializeComponent completed successfully for SqlTab")

            Catch ex As Exception
                _logger?.LogError($"Error in InitializeComponent: {ex.Message}")
                _logger?.LogError($"StackTrace: {ex.StackTrace}")
                Throw
            End Try
        End Sub

        Friend Function IsExecutingFromSqlTab() As Boolean
            Return TypeOf Me Is SqlTab AndAlso Me._serverTextBox IsNot Nothing AndAlso Me._databaseComboBox?.SelectedItem IsNot Nothing
        End Function

        Private Sub ObjectListComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _objectListComboBox.SelectedIndex > 0 Then
                Dim selectedObject As String = _objectListComboBox.SelectedItem.ToString()
                Dim queryType As String = _queryTypeComboBox.SelectedItem.ToString().ToLower()

                Select Case queryType
                    Case "table", "tableau", "view", "vue"
                        _queryTextBox.Text = $"SELECT * FROM {selectedObject}"
                    Case "stored procedure", "procédure stockée"
                        _queryTextBox.Text = $"EXEC {selectedObject}"
                End Select
            End If
        End Sub

        Private Sub PopulateCredentialComboBox()
            _credentialComboBox.Items.Clear()
            _credentialComboBox.Items.Add("(None)")
            Dim targets = _credentialManager.GetAllCredentialTargets()
            For Each target In targets
                _credentialComboBox.Items.Add(target)
            Next
            _credentialComboBox.SelectedIndex = 0
        End Sub

        Private Sub PopulateDatabaseComboBox()

            _databaseComboBox.Items.Clear()
            ' Add delay-loaded connection
            _databaseComboBox.Items.Add("(Select a database)")
            _databaseComboBox.SelectedIndex = 0

            ' Load saved SQL settings
            Dim sqlConfig = _xmlManager.GetSqlConfiguration()
            If Not String.IsNullOrEmpty(sqlConfig("DefaultServer")) Then
                _serverTextBox.Text = sqlConfig("DefaultServer")
            End If

            If String.IsNullOrWhiteSpace(_serverTextBox?.Text) Then
                _databaseComboBox.SelectedIndex = 0
                Return
            End If

            Dim connectionString As String = $"Server={_serverTextBox.Text};Integrated Security=True;Connect Timeout=5"

            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Using command As New SqlCommand("SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name", connection)
                        Using reader As SqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                _databaseComboBox.Items.Add(reader("name").ToString())
                            End While
                        End Using
                    End Using
                End Using

                ' Set saved database if exists
                If _databaseComboBox.Items.Count > 1 Then
                    Dim defaultDb = sqlConfig("DefaultDatabase")
                    If Not String.IsNullOrEmpty(defaultDb) Then
                        Dim defaultIndex = _databaseComboBox.Items.IndexOf(defaultDb)
                        If defaultIndex > -1 Then
                            _databaseComboBox.SelectedIndex = defaultIndex
                            Return
                        End If
                    End If
                End If
                _databaseComboBox.SelectedIndex = 0

            Catch ex As Exception
                _logger?.LogError($"Error populating database list: {ex.Message}")
                MessageBox.Show($"Error connecting to server: {ex.Message}", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub PopulateObjectListComboBox()
            _objectListComboBox.Items.Clear()
            _objectListComboBox.Items.Add("(Select an object)")

            If _databaseComboBox.SelectedIndex <= 0 OrElse String.IsNullOrEmpty(_serverTextBox.Text) Then
                _objectListComboBox.SelectedIndex = 0
                Return
            End If

            Try
                Dim server As String = _serverTextBox.Text.Trim()
                Dim database As String = _databaseComboBox.SelectedItem.ToString()
                Dim queryType As String = _queryTypeComboBox.SelectedItem.ToString()

                Dim connectionString As String = $"Server={server};Database={database};Integrated Security=True;Connect Timeout=5"

                Using connection As New SqlConnection(connectionString)
                    connection.Open()
                    Dim query As String = ""
                    _logger?.LogInfo($"Selected query type: {queryType}")

                    Select Case queryType.ToLower()
                        Case "table", "tableau"
                            query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
                        Case "view", "vue"
                            query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS ORDER BY TABLE_NAME"
                        Case "stored procedure", "procédure stockée"
                            query = "SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_TYPE = 'PROCEDURE' ORDER BY ROUTINE_NAME"
                        Case Else
                            _logger?.LogWarning($"Unknown query type: {queryType}")
                    End Select

                    If Not String.IsNullOrEmpty(query) Then
                        Using command As New SqlCommand(query, connection)
                            Using reader As SqlDataReader = command.ExecuteReader()
                                While reader.Read()
                                    _objectListComboBox.Items.Add(reader(0).ToString())
                                End While
                            End Using
                        End Using
                    End If

                    If _objectListComboBox.Items.Count > 1 Then
                        _objectListComboBox.SelectedIndex = 0
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show($"Error populating object list: {ex.Message}", "Population Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _logger?.LogError($"Error populating object list: {ex.Message}")
            End Try
        End Sub

        Private Sub PopulateQueryTypeComboBox()
            If _queryTypeComboBox IsNot Nothing Then
                _queryTypeComboBox.Items.Clear()
                _queryTypeComboBox.Items.AddRange({
            TranslationManager.Instance.GetTranslation("SQLTab.QueryType.Select"),
            TranslationManager.Instance.GetTranslation("SQLTab.QueryType.Table"),
            TranslationManager.Instance.GetTranslation("SQLTab.QueryType.View"),
            TranslationManager.Instance.GetTranslation("SQLTab.QueryType.StoredProcedure")
        })
                _queryTypeComboBox.SelectedIndex = 0
            End If
        End Sub

        Private Sub QueryTypeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _queryTypeComboBox.SelectedIndex > 0 Then
                PopulateObjectListComboBox()
            End If
        End Sub

        Private Async Sub ReturnLastEntryButton_Click(sender As Object, e As EventArgs)
            Await ExecuteAndDisplayQuery(False, True)
        End Sub

        Private Async Sub RunButton_Click(sender As Object, e As EventArgs)
            Try
                _runButton.Enabled = False
                _outputTextBox.Text = "Executing query..."

                If Not ValidateInputs() Then Return

                Dim sqlConfig = _xmlManager.GetSqlConfiguration()
                Dim server = If(Not String.IsNullOrWhiteSpace(_serverTextBox?.Text),
                       _serverTextBox.Text,
                       sqlConfig("DefaultServer"))

                Dim database = If(_databaseComboBox?.SelectedItem IsNot Nothing AndAlso
                         _databaseComboBox.SelectedItem.ToString() <> "(Select a database)",
                         _databaseComboBox.SelectedItem.ToString(),
                         sqlConfig("DefaultDatabase"))

                Dim selectedCredential = If(_credentialComboBox?.SelectedItem?.ToString(), "(None)")
                Dim credential = If(selectedCredential = "(None)",
                          Nothing,
                          _credentialManager.GetCredential(selectedCredential, "Windows Vault"))

                _logger.LogInfo($"Executing query on {server}.{database}, credential: {selectedCredential}")

                Dim query = _queryTextBox.Text
                Dim queryType = _queryTypeComboBox.SelectedItem.ToString().ToLower()
                Dim result As String

                If queryType.Contains("stored procedure") Then
                    Dim parameters = ParseParameters(query)
                    result = Await _taskRunner.ExecuteStoredProcedureWithSqlCmd(
                ExtractStoredProcedureName(query),
                server,
                database,
                parameters,
                credential)
                Else
                    result = Await _taskRunner.ExecuteSqlCommandWithSqlCmd(
                query,
                server,
                database,
                credential)
                End If

                _outputTextBox.Text = result
                _logger.LogInfo("Query executed successfully")

            Catch ex As Exception
                _logger.LogError($"Error in RunButton_Click: {ex.Message}")
                _outputTextBox.Text = $"Error executing query: {ex.Message}"
            Finally
                _runButton.Enabled = True
            End Try
        End Sub

        Private Function ParseParameters(query As String) As Dictionary(Of String, Object)
            Dim parameters As New Dictionary(Of String, Object)

            Try
                ' Extract parameters from EXEC statement
                Dim paramMatch = Regex.Match(query, "EXEC(?:UTE)?\s+\w+\s*(.+)", RegexOptions.IgnoreCase)
                If paramMatch.Success Then
                    Dim paramString = paramMatch.Groups(1).Value
                    For Each param In paramString.Split(","c)
                        Dim parts = param.Trim().Split("="c)
                        If parts.Length = 2 Then
                            Dim key = parts(0).Trim().TrimStart("@"c)
                            Dim value = parts(1).Trim().Trim("'"c)
                            parameters.Add(key, value)
                        End If
                    Next
                End If
            Catch ex As Exception
                _logger.LogError($"Error parsing parameters: {ex.Message}")
            End Try

            Return parameters
        End Function

        Private Async Function SaveDataTableToCsvAsync(dataTable As DataTable, filePath As String) As Task
            Using writer As New StreamWriter(filePath, False, Encoding.UTF8)
                ' Write headers
                Await writer.WriteLineAsync(String.Join(",", dataTable.Columns.Cast(Of DataColumn).Select(Function(column) $"""{column.ColumnName}""")))

                ' Write rows
                For Each row As DataRow In dataTable.Rows
                    Await writer.WriteLineAsync(String.Join(",", row.ItemArray.Select(Function(field) $"""{field?.ToString().Replace("""", """""")}""")))
                Next
            End Using
        End Function

        Private Async Function SaveResultsAsCsv(data As DataTable) As Task
            Try
                Using saveFileDialog As New SaveFileDialog()
                    saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
                    saveFileDialog.DefaultExt = "csv"
                    saveFileDialog.AddExtension = True

                    If saveFileDialog.ShowDialog() = DialogResult.OK Then
                        Await SaveDataTableToCsvAsync(data, saveFileDialog.FileName)
                        MessageBox.Show("Data saved successfully.", "Save Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error saving CSV: {ex.Message}", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _logger.LogError($"Error saving CSV: {ex.Message}")
            End Try
        End Function

        Private Async Sub SelectAllButton_Click(sender As Object, e As EventArgs)
            Await ExecuteAndDisplayQuery(True, False)
        End Sub

        Private Function ValidateInputs() As Boolean
            If String.IsNullOrWhiteSpace(_queryTextBox.Text) Then
                MessageBox.Show("Please enter a valid SQL query.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            If String.IsNullOrWhiteSpace(_serverTextBox?.Text) Then
                MessageBox.Show("Server must be specified", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            If _databaseComboBox?.SelectedItem Is Nothing OrElse _databaseComboBox.SelectedItem.ToString() = "(Select a database)" Then
                MessageBox.Show("Database must be selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            Return True
        End Function

        Private Enum QueryType
            RegularSQL
            StoredProcedure
            SQLFile
        End Enum

    End Class
End Namespace