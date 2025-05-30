Imports System.Security.Cryptography
Imports System.Text.RegularExpressions

Namespace LiteTask
    Public Class CredentialManager
        Private ReadOnly _logger As Logger
        Private ReadOnly _storedCredentialsPath As String
        Private ReadOnly _entropy As Byte() = Encoding.UTF8.GetBytes("LiteTask_Secure_Storage_Key")
        Private ReadOnly _credentialLock As New Object()
        Private _userToken As IntPtr = IntPtr.Zero

        ' Windows API structures and functions
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure CREDENTIAL
            Public Flags As UInteger
            Public Type As UInteger
            Public TargetName As String
            Public Comment As String
            Public LastWritten As System.Runtime.InteropServices.ComTypes.FILETIME
            Public CredentialBlobSize As UInteger
            Public CredentialBlob As IntPtr
            Public Persist As UInteger
            Public AttributeCount As UInteger
            Public Attributes As IntPtr
            Public TargetAlias As String
            Public UserName As String
        End Structure

        Private Const CRED_TYPE_GENERIC As Integer = 1
        Private Const CRED_PERSIST_LOCAL_MACHINE As Integer = 2

        Public Enum ServiceAccountType
            LocalSystem
            LocalService
            NetworkService
        End Enum


        <DllImport("Advapi32.dll", SetLastError:=True, EntryPoint:="CredReadW", CharSet:=CharSet.Unicode)>
        Private Shared Function CredRead(target As String, type As Integer, reservedFlag As Integer, ByRef credentialPtr As IntPtr) As Boolean
        End Function

        <DllImport("Advapi32.dll", SetLastError:=True, EntryPoint:="CredWriteW", CharSet:=CharSet.Unicode)>
        Private Shared Function CredWrite(ByRef credential As CREDENTIAL, flags As Integer) As Boolean
        End Function

        <DllImport("Advapi32.dll", SetLastError:=True, EntryPoint:="CredFree")>
        Private Shared Sub CredFree(cred As IntPtr)
        End Sub

        <DllImport("Advapi32.dll", SetLastError:=True, EntryPoint:="CredDeleteW", CharSet:=CharSet.Unicode)>
        Private Shared Function CredDelete(target As String, type As Integer, flags As Integer) As Boolean
        End Function

        <DllImport("Advapi32.dll", SetLastError:=True, EntryPoint:="CredEnumerateW", CharSet:=CharSet.Unicode)>
        Private Shared Function CredEnumerate(filter As String, flags As Integer, ByRef count As Integer, ByRef credentials As IntPtr) As Boolean
        End Function

        Private Const LOGON32_LOGON_INTERACTIVE As Integer = 2
        Private Const LOGON32_PROVIDER_DEFAULT As Integer = 0

        Public Sub New(logger As Logger)
            _logger = logger
            _storedCredentialsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LiteTask",
            "SecureCredentials.dat")

            ' Ensure directory exists but keep existing permissions
            Directory.CreateDirectory(Path.GetDirectoryName(_storedCredentialsPath))
        End Sub

        Private Function DecryptString(encryptedInput As String) As String
            Try
                ' Keep existing decryption for compatibility
                Dim encryptedBytes = Convert.FromBase64String(encryptedInput)
                Dim decryptedData = ProtectedData.Unprotect(encryptedBytes, _entropy, DataProtectionScope.CurrentUser)
                Return Encoding.UTF8.GetString(decryptedData)
            Catch ex As Exception
                _logger?.LogError($"Error decrypting string: {ex.Message}")
                Return String.Empty
            End Try
        End Function

        Public Function GetAllCredentialTargets() As List(Of String)
            Dim targets As New List(Of String)

            ' Add system accounts
            targets.Add("(Service Account, NT AUTHORITY\SYSTEM)")
            targets.Add("(Service Account, NT AUTHORITY\LOCAL SERVICE)")
            targets.Add("(Service Account, NT AUTHORITY\NETWORK SERVICE)")

            ' Get Current User credentials
            Dim regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\LiteTask\Credentials")
            If regKey IsNot Nothing Then
                For Each target In regKey.GetValueNames()
                    targets.Add($"(Current User, {target})")
                Next
            End If

            ' Get Service Account credentials
            Dim serviceFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ServiceCredentials.dat")
            If File.Exists(serviceFilePath) Then
                For Each line In File.ReadAllLines(serviceFilePath)
                    Dim parts = line.Split("|"c)
                    If parts.Length >= 1 Then
                        targets.Add($"(Service Account, {parts(0)})")
                    End If
                Next
            End If

            ' Get Windows Vault credentials
            Dim count As Integer = 0
            Dim pCredentials As IntPtr = IntPtr.Zero
            If CredEnumerate(Nothing, 0, count, pCredentials) Then
                Dim credentialPtrs(count - 1) As IntPtr
                Marshal.Copy(pCredentials, credentialPtrs, 0, count)
                For i As Integer = 0 To count - 1
                    Dim credential As CREDENTIAL = Marshal.PtrToStructure(Of CREDENTIAL)(credentialPtrs(i))
                    If credential.Type = CRED_TYPE_GENERIC Then
                        targets.Add($"(Windows Vault, {credential.TargetName})")
                    End If
                Next
                CredFree(pCredentials)
            End If

            ' Get Stored Account credentials
            Dim storedFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "StoredCredentials.dat")
            If File.Exists(storedFilePath) Then
                For Each line In File.ReadAllLines(storedFilePath)
                    Dim parts = line.Split("|"c)
                    If parts.Length >= 1 Then
                        targets.Add($"(Stored Account, {parts(0)})")
                    End If
                Next
            End If

            Return targets
        End Function

        Public Sub DeleteCredential(target As String, accountType As String)
            Try
                Select Case accountType.ToLower()
                    Case "windows vault"
                        DeleteFromWindowsVault(target)
                    Case "stored account"
                        DeleteStoredCredential(target)
                    Case "current user"
                        DeleteCurrentUserCredential(target)
                    Case "service account"
                        ' For service accounts, you might not need to delete anything
                        ' or you might have a special handling
                    Case Else
                        Throw New ArgumentException("Invalid account type", NameOf(accountType))
                End Select
                _logger.LogInfo($"Credential deleted successfully for target: {target}, account type: {accountType}")
            Catch ex As Exception
                _logger.LogError($"Error deleting credential: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub DeleteCurrentUserCredential(target As String)
            Dim regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\LiteTask\Credentials", True)
            regKey?.DeleteValue(target, False)
        End Sub
        Private Sub DeleteFromWindowsVault(target As String)
            If Not CredDelete(target, CRED_TYPE_GENERIC, 0) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End Sub

        Private Sub DeleteStoredCredential(target As String)
            If File.Exists(_storedCredentialsPath) Then
                Dim lines = File.ReadAllLines(_storedCredentialsPath).Where(Function(line) Not line.StartsWith(target & "|")).ToArray()
                File.WriteAllLines(_storedCredentialsPath, lines)
            End If
        End Sub

        Private Function EncryptString(input As String) As String
            Try
                ' Use existing encryption for compatibility
                Dim inputBytes = Encoding.UTF8.GetBytes(input)
                Dim encryptedData = ProtectedData.Protect(inputBytes, _entropy, DataProtectionScope.CurrentUser)
                Return Convert.ToBase64String(encryptedData)
            Catch ex As Exception
                _logger?.LogError($"Error encrypting string: {ex.Message}")
                Throw
            End Try
        End Function

        Private Sub EnsureDirectoryExists(path As String)
            If Not Directory.Exists(path) Then
                Directory.CreateDirectory(path)
            End If
        End Sub

        Private Function FormatTargetName(target As String) As String
            ' Ensure target has consistent format for Windows Vault
            If Not target.Contains("LiteTask_") Then
                Return $"LiteTask_{target}"
            End If
            Return target
        End Function

        'Public Function GetCredential(target As String, accountType As String) As CredentialInfo
        '    Try
        '        ' Support both formats:
        '        ' 1. Legacy format from XML: direct accountType
        '        ' 2. Combined format: (accountType,target)
        '        If Not target.StartsWith("(") Then
        '            ' Handle legacy format where only AccountType is provided
        '            Return GetCredentialByType(target, accountType)
        '        Else
        '            ' Handle combined format (accountType,target)
        '            Dim parts = target.Trim("()".ToCharArray()).Split(",")
        '            If parts.Length = 2 Then
        '                accountType = parts(0).Trim()
        '                target = parts(1).Trim()
        '            End If
        '            Return GetCredentialByType(target, accountType)
        '        End If
        '    Catch ex As Exception
        '        _logger?.LogError($"Error getting credential: {ex.Message}")
        '        Return Nothing
        '    End Try
        'End Function

        'Private Function GetCredentialByType(target As String, accountType As String) As CredentialInfo
        '    Select Case accountType.ToLower()
        '        Case "current user"
        '            Return GetCurrentUserCredential(target)
        '        Case "windows vault"
        '            Return GetFromWindowsVault(target)
        '        Case "stored account"
        '            Return GetStoredCredential(target)
        '        Case "service account"
        '            Return GetServiceAccountCredential(target)
        '        Case Else
        '            Return Nothing
        '    End Select
        'End Function

        'Private Function GetCurrentUserCredential(target As String) As CredentialInfo
        '    Try
        '        ' Get credentials from registry
        '        Using regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\LiteTask\Credentials")
        '            Dim value = TryCast(regKey?.GetValue(target), String)
        '            If value IsNot Nothing Then
        '                Dim parts = value.Split("|"c)
        '                If parts.Length = 2 Then
        '                    Dim decryptedPassword = DecryptString(parts(1))
        '                    Return New CredentialInfo With {
        '                .Target = target,
        '                .Username = parts(0),
        '                .Password = decryptedPassword,
        '                .AccountType = "Current User",
        '                .SecurePassword = New NetworkCredential("", decryptedPassword).SecurePassword
        '            }
        '                End If
        '            End If
        '            Return Nothing
        '        End Using
        '    Catch ex As Exception
        '        _logger.LogError($"Error getting current user credential: {ex.Message}")
        '        Return Nothing
        '    End Try
        'End Function

        'Private Function GetFromWindowsVault(target As String) As CredentialInfo
        '    Dim credentialPtr As IntPtr = IntPtr.Zero
        '    Try
        '        ValidateTarget(target)

        '        ' Format target for Windows Vault
        '        Dim formattedTarget = FormatTargetName(target)
        '        If Not CredRead(formattedTarget, CRED_TYPE_GENERIC, 0, credentialPtr) Then
        '            Return Nothing
        '        End If

        '        Dim credential As CREDENTIAL = Marshal.PtrToStructure(Of CREDENTIAL)(credentialPtr)
        '        If credential.CredentialBlobSize = 0 OrElse credential.CredentialBlob = IntPtr.Zero Then
        '            Return Nothing
        '        End If

        '        ' Securely copy password
        '        Dim passwordBytes(credential.CredentialBlobSize - 1) As Byte
        '        Marshal.Copy(credential.CredentialBlob, passwordBytes, 0, credential.CredentialBlobSize)

        '        ' Create secure string
        '        Dim securePassword As New SecureString()
        '        Dim password = Encoding.Unicode.GetString(passwordBytes)
        '        For Each c In password
        '            securePassword.AppendChar(c)
        '        Next
        '        securePassword.MakeReadOnly()

        '        ' Clear sensitive data
        '        Array.Clear(passwordBytes, 0, passwordBytes.Length)

        '        Return New CredentialInfo With {
        '        .Target = target,
        '        .Username = credential.UserName,
        '        .SecurePassword = securePassword,
        '        .AccountType = "Windows Vault"
        '    }

        '    Catch ex As Exception
        '        _logger.LogError($"Error reading from Windows Vault: {ex.Message}")
        '        Return Nothing
        '    Finally
        '        If credentialPtr <> IntPtr.Zero Then
        '            CredFree(credentialPtr)
        '        End If
        '    End Try
        'End Function

        Public Function GetCredential(target As String, accountType As String) As CredentialInfo
            Try
                If target.StartsWith("(") AndAlso target.EndsWith(")") Then
                    Dim parts = target.Trim("()".ToCharArray()).Split(",")
                    If parts.Length = 2 Then
                        accountType = parts(0).Trim()
                        target = parts(1).Trim()
                    End If
                End If

                ' Checking target and account type for Windows Vault credentials
                If accountType.ToLower() = "windows vault" Then
                    ' Remove any "LiteTask_" prefix if it exists
                    If target.StartsWith("LiteTask_") Then
                        target = target.Substring(8)
                    End If
                    ' Get from Windows Vault
                    Return GetFromWindowsVault(target)
                End If

                ' For other credential types
                Select Case accountType.ToLower()
                    Case "stored account"
                        Return GetStoredCredential(target)
                    Case "current user"
                        Return GetCurrentUserCredential(target)
                    Case "service account"
                        Return GetServiceAccountCredential(target)
                    Case Else
                        _logger?.LogWarning($"Unsupported account type: {accountType}")
                        Return Nothing
                End Select

            Catch ex As Exception
                _logger?.LogError($"Error getting credential: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Function GetCurrentUserCredential(target As String) As CredentialInfo
            Try
                Dim regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\LiteTask\Credentials")
                Dim value = TryCast(regKey?.GetValue(target), String)
                If value IsNot Nothing Then
                    Dim parts = value.Split("|"c)
                    If parts.Length = 2 Then
                        Return New CredentialInfo() With {
                            .Target = target,
                            .Username = parts(0),
                            .Password = DecryptString(parts(1)),
                            .AccountType = "Current User",
                            .SecurePassword = New NetworkCredential("", DecryptString(parts(1))).SecurePassword
                        }
                    End If
                End If
                Return Nothing
            Catch ex As Exception
                _logger.LogError($"Error getting current user credential: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Function GetFromWindowsVault(target As String) As CredentialInfo
            Dim credentialPtr As IntPtr = IntPtr.Zero
            Try
                ' Format target to ensure consistent lookup
                Dim formattedTarget = FormatTargetName(target)
                If Not CredRead(formattedTarget, CRED_TYPE_GENERIC, 0, credentialPtr) Then
                    ' Try without formatting as fallback
                    If Not CredRead(target, CRED_TYPE_GENERIC, 0, credentialPtr) Then
                        Return Nothing
                    End If
                End If

                Dim credential As CREDENTIAL = Marshal.PtrToStructure(Of CREDENTIAL)(credentialPtr)
                If credential.CredentialBlobSize = 0 OrElse credential.CredentialBlob = IntPtr.Zero Then
                    Return Nothing
                End If

                Dim passwordBytes(credential.CredentialBlobSize - 1) As Byte
                Marshal.Copy(credential.CredentialBlob, passwordBytes, 0, credential.CredentialBlobSize)
                Dim password As String = Encoding.Unicode.GetString(passwordBytes)

                Return New CredentialInfo With {
            .Target = target,
            .Username = credential.UserName,
            .Password = password,
            .AccountType = "Windows Vault",
            .SecurePassword = New NetworkCredential("", password).SecurePassword
        }
            Catch ex As Exception
                _logger.LogError($"Error reading credential from Windows Vault: {ex.Message}")
                Return Nothing
            Finally
                If credentialPtr <> IntPtr.Zero Then
                    CredFree(credentialPtr)
                End If
            End Try
        End Function

        Private Function GetStoredCredential(target As String) As CredentialInfo
            Try
                If File.Exists(_storedCredentialsPath) Then
                    For Each line In File.ReadAllLines(_storedCredentialsPath)
                        Dim parts = line.Split("|"c)
                        If parts.Length = 3 AndAlso parts(0) = target Then
                            Return New CredentialInfo() With {
                                .Target = parts(0),
                                .Username = parts(1),
                                .Password = DecryptString(parts(2)),
                                .AccountType = "Stored Account",
                                .SecurePassword = New NetworkCredential("", DecryptString(parts(2))).SecurePassword
                                 }
                        End If
                    Next
                End If
                Return Nothing
            Catch ex As Exception
                _logger.LogError($"Error getting stored credential: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Public Function GetServiceAccountCredential(target As String) As CredentialInfo
            ' Handle system accounts
            If target = "NT AUTHORITY\SYSTEM" OrElse
           target = "NT AUTHORITY\LOCAL SERVICE" OrElse
           target = "NT AUTHORITY\NETWORK SERVICE" Then
                Return New CredentialInfo() With {
                    .Target = target,
                    .Username = target,
                    .AccountType = "Service Account",
                    .Password = "",
                    .SecurePassword = New SecureString()
                }
            End If
            Return Nothing
        End Function

        Public Shared Function GetUserSidByName(username As String) As String
            Try
                Dim ntAccount = New NTAccount(username)
                Dim sid = CType(ntAccount.Translate(GetType(SecurityIdentifier)), SecurityIdentifier)
                Return sid.Value
            Catch ex As Exception
                '_logger?.LogError($"Error getting Sid: {ex.Message}") '
                Return String.Empty
            End Try
        End Function

        Public Sub SaveCredential(credInfo As CredentialInfo, password As SecureString)
            Try
                If credInfo Is Nothing OrElse password Is Nothing Then
                    Throw New ArgumentNullException()
                End If

                SyncLock _credentialLock
                    Select Case credInfo.AccountType.ToLower()
                        Case "windows vault"
                            SaveToWindowsVault(credInfo, password)
                        Case "stored account"
                            SaveStoredCredential(credInfo, password)
                        Case "current user"
                            SaveCurrentUserCredential(credInfo, password)
                        Case Else
                            Throw New ArgumentException("Invalid account type", NameOf(credInfo.AccountType))
                    End Select
                End SyncLock

                _logger?.LogInfo($"Credential saved successfully for target: {credInfo.Target}")

            Catch ex As Exception
                _logger?.LogError($"Error saving credential: {ex.Message}")
                Throw
            Finally
                If password IsNot Nothing Then
                    password.Dispose()
                End If
            End Try
        End Sub

        Private Sub SaveCurrentUserCredential(credInfo As CredentialInfo, password As SecureString)
            Dim encryptedPassword = EncryptString(New NetworkCredential(String.Empty, password).Password)
            Dim regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\LiteTask\Credentials")
            regKey.SetValue(credInfo.Target, $"{credInfo.Username}|{encryptedPassword}")
        End Sub

        Private Sub SaveStoredCredential(credInfo As CredentialInfo, password As SecureString)
            Dim encryptedPassword = EncryptString(New NetworkCredential(String.Empty, password).Password)
            File.AppendAllText(_storedCredentialsPath, $"{credInfo.Target}|{credInfo.Username}|{encryptedPassword}{Environment.NewLine}")
        End Sub

        Private Sub SaveToWindowsVault(credInfo As CredentialInfo, password As SecureString)
            Dim credentialPtr As IntPtr = IntPtr.Zero
            Try
                ValidateCredentialInfo(credInfo)

                ' Create credential structure
                Dim cred As New CREDENTIAL With {
                .Type = CRED_TYPE_GENERIC,
                .TargetName = FormatTargetName(credInfo.Target),
                .UserName = credInfo.Username,
                .Persist = CRED_PERSIST_LOCAL_MACHINE
            }

                ' Convert SecureString to encrypted bytes
                Dim bstr As IntPtr = Marshal.SecureStringToBSTR(password)
                Dim passwordStr = Marshal.PtrToStringBSTR(bstr)
                Dim passwordBytes = Encoding.Unicode.GetBytes(passwordStr)
                cred.CredentialBlobSize = CUInt(passwordBytes.Length)
                cred.CredentialBlob = Marshal.AllocHGlobal(passwordBytes.Length)
                Marshal.Copy(passwordBytes, 0, cred.CredentialBlob, passwordBytes.Length)

                ' Clear sensitive data
                Array.Clear(passwordBytes, 0, passwordBytes.Length)


                ' Write to vault
                If Not CredWrite(cred, 0) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

            Finally
                If credentialPtr <> IntPtr.Zero Then
                    Marshal.FreeHGlobal(credentialPtr)
                End If
                GC.Collect()
            End Try
        End Sub

        Private Sub ValidateTarget(target As String)
            If String.IsNullOrWhiteSpace(target) Then
                Throw New ArgumentException("Target cannot be empty")
            End If

            If Not Regex.IsMatch(target, "^[a-zA-Z0-9\\_\\-\\.\\s]+$") Then
                Throw New ArgumentException("Invalid target format")
            End If
        End Sub

        Private Sub ValidateAccountType(accountType As String)
            If String.IsNullOrWhiteSpace(accountType) Then
                Throw New ArgumentException("Account type cannot be empty")
            End If

            Dim validTypes = New String() {"windows vault", "stored account", "current user", "service account"}
            If Not validTypes.Contains(accountType.ToLower()) Then
                Throw New ArgumentException("Invalid account type")
            End If
        End Sub

        Private Sub ValidateCredentialInfo(credInfo As CredentialInfo)
            If credInfo Is Nothing Then
                Throw New ArgumentNullException(NameOf(credInfo))
            End If

            ValidateTarget(credInfo.Target)
            ValidateAccountType(credInfo.AccountType)

            If String.IsNullOrWhiteSpace(credInfo.Username) Then
                Throw New ArgumentException("Username cannot be empty")
            End If

            If Not Regex.IsMatch(credInfo.Username, "^[a-zA-Z0-9\\_\\-\\.\\@\\\\]+$") Then
                Throw New ArgumentException("Invalid username format")
            End If
        End Sub

    End Class
End Namespace
