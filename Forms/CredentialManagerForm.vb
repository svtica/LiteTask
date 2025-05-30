Imports SecurePasswordTextBox

Namespace LiteTask
    Public Class CredentialManagerForm
        Inherits Form

        Private ReadOnly _credentialManager As CredentialManager
        Private _credentialListView As ListView
        Private _targetTextBox As TextBox
        Private _usernameTextBox As TextBox
        Private _passwordTextBox As SecureTextBox
        Private _accountTypeComboBox As ComboBox
        Private _addButton As Button
        Private _updateButton As Button
        Private _deleteButton As Button
        Private _closeButton As Button
        Private _targetLabel As Label
        Private _usernameLabel As Label
        Private _passwordLabel As Label
        Private _accountTypeLabel As Label

        Public Sub New()
            InitializeComponent()
            _credentialManager = ApplicationContainer.GetService(Of CredentialManager)()
            PopulateCredentialList()
            SetupEventHandlers()
            Me.Translate()
        End Sub

        Private Sub InitializeComponent()
            Dim SecureString1 As SecureString = New SecureString()
            Dim resources As ComponentResourceManager = New ComponentResourceManager(GetType(CredentialManagerForm))
            _credentialListView = New ListView()
            _targetTextBox = New TextBox()
            _usernameTextBox = New TextBox()
            _passwordTextBox = New SecureTextBox()
            _accountTypeComboBox = New ComboBox()
            _addButton = New Button()
            _updateButton = New Button()
            _deleteButton = New Button()
            _closeButton = New Button()
            _targetLabel = New Label()
            _usernameLabel = New Label()
            _passwordLabel = New Label()
            _accountTypeLabel = New Label()
            SuspendLayout()
            ' 
            ' _credentialListView
            ' 
            _credentialListView.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            _credentialListView.FullRowSelect = True
            _credentialListView.Location = New System.Drawing.Point(25, 9)
            _credentialListView.Margin = New Padding(3, 2, 3, 2)
            _credentialListView.Name = "_credentialListView"
            _credentialListView.Size = New System.Drawing.Size(362, 138)
            _credentialListView.TabIndex = 0
            _credentialListView.UseCompatibleStateImageBehavior = False
            _credentialListView.View = View.Details
            ' 
            ' _targetTextBox
            ' 
            _targetTextBox.Location = New System.Drawing.Point(119, 166)
            _targetTextBox.Margin = New Padding(3, 2, 3, 2)
            _targetTextBox.Name = "_targetTextBox"
            _targetTextBox.Size = New System.Drawing.Size(172, 23)
            _targetTextBox.TabIndex = 2
            ' 
            ' _usernameTextBox
            ' 
            _usernameTextBox.Location = New System.Drawing.Point(119, 205)
            _usernameTextBox.Margin = New Padding(3, 2, 3, 2)
            _usernameTextBox.Name = "_usernameTextBox"
            _usernameTextBox.Size = New System.Drawing.Size(172, 23)
            _usernameTextBox.TabIndex = 4
            ' 
            ' _passwordTextBox
            ' 
            _passwordTextBox.Location = New System.Drawing.Point(119, 245)
            _passwordTextBox.Margin = New Padding(3, 2, 3, 2)
            _passwordTextBox.Name = "_passwordTextBox"
            _passwordTextBox.SecureText = SecureString1
            _passwordTextBox.Size = New System.Drawing.Size(172, 23)
            _passwordTextBox.TabIndex = 6
            _passwordTextBox.UseSystemPasswordChar = True
            ' 
            ' _accountTypeComboBox
            ' 
            _accountTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            _accountTypeComboBox.FormattingEnabled = True
            _accountTypeComboBox.Items.AddRange(New Object() {"Current User", "Service Account", "Windows Vault", "Stored Account"})
            _accountTypeComboBox.Location = New System.Drawing.Point(119, 291)
            _accountTypeComboBox.Margin = New Padding(3, 2, 3, 2)
            _accountTypeComboBox.Name = "_accountTypeComboBox"
            _accountTypeComboBox.Size = New System.Drawing.Size(172, 23)
            _accountTypeComboBox.TabIndex = 8
            ' 
            ' _addButton
            ' 
            _addButton.Location = New System.Drawing.Point(309, 166)
            _addButton.Margin = New Padding(3, 2, 3, 2)
            _addButton.Name = "_addButton"
            _addButton.Size = New System.Drawing.Size(93, 36)
            _addButton.TabIndex = 9
            _addButton.Text = "Add"
            _addButton.UseVisualStyleBackColor = True
            ' 
            ' _updateButton
            ' 
            _updateButton.Location = New System.Drawing.Point(309, 224)
            _updateButton.Margin = New Padding(3, 2, 3, 2)
            _updateButton.Name = "_updateButton"
            _updateButton.Size = New System.Drawing.Size(93, 39)
            _updateButton.TabIndex = 10
            _updateButton.Text = "Update"
            _updateButton.UseVisualStyleBackColor = True
            ' 
            ' _deleteButton
            ' 
            _deleteButton.Location = New System.Drawing.Point(309, 281)
            _deleteButton.Margin = New Padding(3, 2, 3, 2)
            _deleteButton.Name = "_deleteButton"
            _deleteButton.Size = New System.Drawing.Size(93, 35)
            _deleteButton.TabIndex = 11
            _deleteButton.Text = "Delete"
            _deleteButton.UseVisualStyleBackColor = True
            ' 
            ' _closeButton
            ' 
            _closeButton.Location = New System.Drawing.Point(161, 332)
            _closeButton.Margin = New Padding(3, 2, 3, 2)
            _closeButton.Name = "_closeButton"
            _closeButton.Size = New System.Drawing.Size(96, 25)
            _closeButton.TabIndex = 12
            _closeButton.Text = "Close"
            _closeButton.UseVisualStyleBackColor = True
            ' 
            ' _targetLabel
            ' 
            _targetLabel.AutoSize = True
            _targetLabel.Location = New System.Drawing.Point(12, 172)
            _targetLabel.Name = "_targetLabel"
            _targetLabel.Size = New System.Drawing.Size(42, 15)
            _targetLabel.TabIndex = 1
            _targetLabel.Text = "Target:"
            ' 
            ' _usernameLabel
            ' 
            _usernameLabel.AutoSize = True
            _usernameLabel.Location = New System.Drawing.Point(12, 205)
            _usernameLabel.Name = "_usernameLabel"
            _usernameLabel.Size = New System.Drawing.Size(63, 15)
            _usernameLabel.TabIndex = 3
            _usernameLabel.Text = "Username:"
            ' 
            ' _passwordLabel
            ' 
            _passwordLabel.AutoSize = True
            _passwordLabel.Location = New System.Drawing.Point(12, 248)
            _passwordLabel.Name = "_passwordLabel"
            _passwordLabel.Size = New System.Drawing.Size(60, 15)
            _passwordLabel.TabIndex = 5
            _passwordLabel.Text = "Password:"
            ' 
            ' _accountTypeLabel
            ' 
            _accountTypeLabel.AutoSize = True
            _accountTypeLabel.Location = New System.Drawing.Point(12, 291)
            _accountTypeLabel.Name = "_accountTypeLabel"
            _accountTypeLabel.Size = New System.Drawing.Size(82, 15)
            _accountTypeLabel.TabIndex = 7
            _accountTypeLabel.Text = "Account Type:"
            ' 
            ' CredentialManagerForm
            ' 
            AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
            AutoScaleMode = AutoScaleMode.Font
            ClientSize = New System.Drawing.Size(414, 366)
            Controls.Add(_credentialListView)
            Controls.Add(_targetLabel)
            Controls.Add(_targetTextBox)
            Controls.Add(_usernameLabel)
            Controls.Add(_usernameTextBox)
            Controls.Add(_passwordLabel)
            Controls.Add(_passwordTextBox)
            Controls.Add(_accountTypeLabel)
            Controls.Add(_accountTypeComboBox)
            Controls.Add(_addButton)
            Controls.Add(_updateButton)
            Controls.Add(_deleteButton)
            Controls.Add(_closeButton)
            FormBorderStyle = FormBorderStyle.FixedDialog
            Icon = CType(resources.GetObject("$this.Icon"), Drawing.Icon)
            Margin = New Padding(3, 2, 3, 2)
            MaximizeBox = False
            MinimizeBox = False
            Name = "CredentialManagerForm"
            StartPosition = FormStartPosition.CenterParent
            Text = "Credential Manager"
            ResumeLayout(False)
            PerformLayout()

        End Sub

        Private Sub AddButton_Click(sender As Object, e As EventArgs)
            If ValidateInput() Then
                Dim accountType = _accountTypeComboBox.SelectedItem.ToString()
                Dim securePassword As New Security.SecureString()
                For Each c As Char In _passwordTextBox.Text
                    securePassword.AppendChar(c)
                Next
                _credentialManager.SaveCredential(New CredentialInfo With {
            .Target = _targetTextBox.Text,
            .Username = _usernameTextBox.Text,
            .AccountType = accountType
        }, _passwordTextBox.SecureText)
                PopulateCredentialList()
                ClearFields()
                MessageBox.Show("Credential added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub ClearFields()
            _targetTextBox.Text = ""
            _usernameTextBox.Text = ""
            _passwordTextBox.SecureText.Clear()
            _accountTypeComboBox.SelectedIndex = -1
        End Sub

        Private Sub CloseButton_Click(sender As Object, e As EventArgs)
            Me.Close()
        End Sub

        Private Sub CredentialListView_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _credentialListView.SelectedItems.Count > 0 Then
                Dim selectedItem = _credentialListView.SelectedItems(0)
                _accountTypeComboBox.SelectedItem = selectedItem.Text
                _targetTextBox.Text = selectedItem.SubItems(1).Text
                _usernameTextBox.Text = selectedItem.SubItems(2).Text
                _passwordTextBox.Text = "" ' For security reasons, don't populate the password
            Else
                ClearFields()
            End If
        End Sub

        Private Sub DeleteButton_Click(sender As Object, e As EventArgs)
            If _credentialListView.SelectedItems.Count > 0 Then
                Dim accountType = _credentialListView.SelectedItems(0).Text
                Dim target = _credentialListView.SelectedItems(0).SubItems(1).Text
                If MessageBox.Show($"Are you sure you want to delete the credential for '{target}' ({accountType})?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    _credentialManager.DeleteCredential(target, accountType)
                    PopulateCredentialList()
                    ClearFields()
                    MessageBox.Show("Credential deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("Please select a credential to delete.", "Delete Credential", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub PopulateCredentialList()
            Try
                _credentialListView.Items.Clear()
                _credentialListView.Columns.Clear()
                _credentialListView.Columns.Add("Account Type", 100)
                _credentialListView.Columns.Add("Target", 150)
                _credentialListView.Columns.Add("Username", 150)

                Dim credentials = _credentialManager.GetAllCredentialTargets()

                For Each cred In credentials
                    ' Parse the combined string
                    Dim parts = cred.Trim("()".ToCharArray()).Split(",")
                    If parts.Length = 2 Then
                        Dim accountType = parts(0).Trim()
                        Dim target = parts(1).Trim()

                        Dim item = New ListViewItem(accountType)
                        item.SubItems.Add(target)

                        Dim credential = _credentialManager.GetCredential(target, accountType)
                        If credential IsNot Nothing Then
                            item.SubItems.Add(credential.Username)
                        Else
                            item.SubItems.Add("N/A")
                        End If

                        _credentialListView.Items.Add(item)
                    End If
                Next

                If _credentialListView.Items.Count = 0 Then
                    MessageBox.Show("No credentials found or unable to retrieve credentials.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show($"Error populating credential list: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub SetupEventHandlers()
            AddHandler _addButton.Click, AddressOf AddButton_Click
            AddHandler _updateButton.Click, AddressOf UpdateButton_Click
            AddHandler _deleteButton.Click, AddressOf DeleteButton_Click
            AddHandler _closeButton.Click, AddressOf CloseButton_Click
            AddHandler _credentialListView.SelectedIndexChanged, AddressOf CredentialListView_SelectedIndexChanged
        End Sub

        Private Sub UpdateButton_Click(sender As Object, e As EventArgs)
            If _credentialListView.SelectedItems.Count > 0 AndAlso ValidateInput() Then
                Dim accountType = _accountTypeComboBox.SelectedItem.ToString()
                Dim securePassword As New Security.SecureString()
                For Each c As Char In _passwordTextBox.Text
                    securePassword.AppendChar(c)
                Next
                Dim credInfo As New CredentialInfo With {
            .Target = _targetTextBox.Text,
            .Username = _usernameTextBox.Text,
            .AccountType = accountType
        }
                _credentialManager.SaveCredential(credInfo, securePassword)
                PopulateCredentialList()
                ClearFields()
                MessageBox.Show("Credential updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Please select a credential to update.", "Update Credential", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Function ValidateInput() As Boolean
            If String.IsNullOrWhiteSpace(_targetTextBox.Text) Then
                MessageBox.Show("Target is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrWhiteSpace(_usernameTextBox.Text) Then
                MessageBox.Show("Username is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If String.IsNullOrWhiteSpace(_passwordTextBox.Text) Then
                MessageBox.Show("Password is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            If _accountTypeComboBox.SelectedIndex = -1 Then
                MessageBox.Show("Please select an account type.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            Return True
        End Function

    End Class
End Namespace