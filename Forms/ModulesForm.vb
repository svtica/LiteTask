Imports System.Drawing

Namespace LiteTask
    Public Class ModulesForm
        Inherits Form

        Private ReadOnly _logger As Logger
        Private _moduleChecklist As CheckedListBox
        Private _installButton As Button
        Private _closeButton As Button
        Private _statusLabel As Label
        Private _isInstalling As Boolean = False

        ' List of recommended PowerShell modules with descriptions
        Private ReadOnly _modules As New Dictionary(Of String, String) From {
            {"Az", "Azure PowerShell module for managing Azure resources"},
            {"AzureAD", "Azure Active Directory management"},
            {"MSOnline", "Microsoft 365 management"},
            {"PSWindowsUpdate", "Windows Update management"},
            {"PSBlueTeam", "Security and hardening tools"},
            {"Pester", "Testing framework for PowerShell"},
            {"ImportExcel", "Excel import/export capabilities"},
            {"VMware.PowerCLI", "VMware infrastructure management"},
            {"SqlServer", "SQL Server management"},
            {"AWS.Tools.Common", "AWS PowerShell tools"}
        }

        Public Sub New()
            _logger = ApplicationContainer.GetService(Of Logger)()
            InitializeComponent()
            Me.Translate()
        End Sub

        Private Sub CloseButton_Click(sender As Object, e As EventArgs)
            If _isInstalling Then
                If MessageBox.Show("Installation is in progress. Are you sure you want to close?", "Confirm Close", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                    Return
                End If
            End If
            Close()
        End Sub

        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ModulesForm))

            ' Initialize controls
            _moduleChecklist = New CheckedListBox With {
                .Dock = DockStyle.Top,
                .Height = 300,
                .CheckOnClick = True
            }

            _statusLabel = New Label With {
                .Dock = DockStyle.Bottom,
                .Height = 40,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(10, 0, 10, 0)
            }

            _installButton = New Button With {
                .Text = "Install Selected",
                .Dock = DockStyle.Bottom,
                .Height = 30
            }

            _closeButton = New Button With {
                .Text = "Close",
                .Dock = DockStyle.Bottom,
                .Height = 30
            }

            ' Set form properties
            Text = "PowerShell Module Manager"
            Size = New Size(500, 450)
            MinimumSize = New Size(400, 400)
            StartPosition = FormStartPosition.CenterParent
            Icon = CType(resources.GetObject("$this.Icon"), Icon)

            ' Add modules to checklist
            For Each kvp In _modules
                _moduleChecklist.Items.Add($"{kvp.Key} - {kvp.Value}")
            Next

            ' Add controls to form
            Controls.Add(_moduleChecklist)
            Controls.Add(_statusLabel)
            Controls.Add(_installButton)
            Controls.Add(_closeButton)

            ' Add event handlers
            AddHandler _installButton.Click, AddressOf InstallButton_Click
            AddHandler _closeButton.Click, AddressOf CloseButton_Click
            AddHandler Load, AddressOf ModulesForm_Load
        End Sub

        Private Async Sub InstallButton_Click(sender As Object, e As EventArgs)
            If _isInstalling Then
                MessageBox.Show("Installation is already in progress.", "Please Wait", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Try
                ' Get selected modules
                Dim selectedModules = _moduleChecklist.CheckedItems.Cast(Of String)().
            Select(Function(item) _modules.Keys.ElementAt(_moduleChecklist.Items.IndexOf(item))).
            Where(Function(m) Not _moduleChecklist.Items(_modules.Keys.ToList().IndexOf(m)).ToString().Contains("[Installed]")).
            ToArray()

                If selectedModules.Length = 0 Then
                    MessageBox.Show("Please select at least one module to install.", "No Modules Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                _isInstalling = True
                _installButton.Enabled = False
                _moduleChecklist.Enabled = False
                UpdateStatus("Starting installation...")

                ' Create necessary paths
                Dim baseDir = Path.Combine(Application.StartupPath, "LiteTaskData")
                Dim modulesPath = Path.Combine(baseDir, "Modules")
                Dim logsPath = Path.Combine(baseDir, "logs")

                ' Ensure directories exist
                Directory.CreateDirectory(modulesPath)
                Directory.CreateDirectory(logsPath)

                ' Create log file path with timestamp
                Dim logFile = Path.Combine(logsPath, $"ModuleInstall_{DateTime.Now:yyyyMMddHHmmss}.log")

                ' Ensure script exists
                Dim scriptPath = Path.Combine(Application.StartupPath, "tools", "InstallModules.ps1")
                If Not File.Exists(scriptPath) Then
                    Throw New FileNotFoundException("Installation script not found", scriptPath)
                End If

                Dim modulesStr = String.Join(",", selectedModules)

                ' Build PowerShell command
                Dim startInfo = New ProcessStartInfo With {
            .FileName = "powershell.exe",
            .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{scriptPath}"" -Modules ""{modulesStr}"" -Destination ""{modulesPath}"" -LogPath ""{logFile}""",
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True
        }

                Using process As New Process With {.StartInfo = startInfo}
                    AddHandler process.OutputDataReceived, Sub(s, args)
                                                               If args.Data IsNot Nothing Then
                                                                   _logger.LogInfo($"PowerShell output: {args.Data}")
                                                                   UpdateStatus(args.Data)
                                                               End If
                                                           End Sub

                    AddHandler process.ErrorDataReceived, Sub(s, args)
                                                              If args.Data IsNot Nothing Then
                                                                  _logger.LogError($"PowerShell error: {args.Data}")
                                                                  UpdateStatus($"Error: {args.Data}")
                                                              End If
                                                          End Sub

                    process.Start()
                    process.BeginOutputReadLine()
                    process.BeginErrorReadLine()
                    Await process.WaitForExitAsync()

                    If process.ExitCode = 0 Then
                        UpdateStatus("Installation completed successfully")
                        MessageBox.Show("Modules installed successfully.", "Installation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        RefreshModuleStatus()
                    Else
                        UpdateStatus("Installation completed with errors")
                        If File.Exists(logFile) Then
                            Dim logContent = Await File.ReadAllTextAsync(logFile)
                            _logger.LogError($"Installation log:{Environment.NewLine}{logContent}")
                        End If
                        MessageBox.Show("Some modules failed to install. Check the application log for details.", "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End Using

            Catch ex As Exception
                UpdateStatus("Installation failed")
                _logger.LogError($"Error installing modules: {ex.Message}")
                _logger.LogError($"StackTrace: {ex.StackTrace}")
                MessageBox.Show($"Error installing modules: {ex.Message}", "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                _isInstalling = False
                _installButton.Enabled = True
                _moduleChecklist.Enabled = True
            End Try
        End Sub

        Private Sub ModulesForm_Load(sender As Object, e As EventArgs)
            Try
                RefreshModuleStatus()
                _statusLabel.Text = "Select modules to install"
            Catch ex As Exception
                _logger.LogError($"Error in ModulesForm_Load: {ex.Message}")
                MessageBox.Show($"Error initializing form: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub RefreshModuleStatus()
            Try
                Dim modulesPath = Path.Combine(Application.StartupPath, "LiteTaskData", "Modules")
                For i As Integer = 0 To _moduleChecklist.Items.Count - 1
                    Dim moduleName = _modules.Keys(i)
                    Dim moduleDir = Path.Combine(modulesPath, moduleName)
                    If Directory.Exists(moduleDir) Then
                        _moduleChecklist.Items(i) = $"{_moduleChecklist.Items(i)} [Installed]"
                    End If
                Next
            Catch ex As Exception
                _logger.LogError($"Error refreshing module status: {ex.Message}")
            End Try
        End Sub

        Private Sub UpdateStatus(message As String)
            If Me.InvokeRequired Then
                Me.Invoke(Sub() UpdateStatus(message))
                Return
            End If

            _statusLabel.Text = message
            Application.DoEvents()
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            If _isInstalling Then
                If MessageBox.Show("Installation is in progress. Are you sure you want to close?", "Confirm Close", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                    e.Cancel = True
                    Return
                End If
            End If
            MyBase.OnFormClosing(e)
        End Sub

    End Class
End Namespace