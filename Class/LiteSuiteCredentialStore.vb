' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    LiteSuiteCredentialStore.vb
' Author:  LiteTask contributors
' Date:    2026-11-12
' Purpose: Credentials shared by the LiteSuite tools, in the Windows Credential
'          Manager of the current user (generic credentials, target
'          "LiteSuite:<name>"). Same store and format as LiteSuite.CredentialStore
'          of LiteDeploy (src/LiteSuite.Common/CredentialStore.cs): an account
'          saved in LiteDeploy (Tools > Credentials) is usable here as the
'          account type "LiteSuite", and the other way round. Windows protects
'          the passwords with the user's DPAPI keys.
' -----------------------------------------------------------------------------
Imports System.ComponentModel

Namespace LiteTask

    Public Class LiteSuiteCredential
        Public Property Name As String
        Public Property UserName As String
        Public Property Password As String
    End Class

    Public NotInheritable Class LiteSuiteCredentialStore
        Private Sub New()
        End Sub

        Public Const Prefix As String = "LiteSuite:"
        Public Const AccountType As String = "LiteSuite"

        Private Const CRED_TYPE_GENERIC As Integer = 1
        Private Const CRED_PERSIST_LOCAL_MACHINE As Integer = 2
        Private Const ERROR_NOT_FOUND As Integer = 1168

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

        <DllImport("advapi32.dll", EntryPoint:="CredWriteW", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CredWrite(ByRef credential As CREDENTIAL, flags As UInteger) As Boolean
        End Function

        <DllImport("advapi32.dll", EntryPoint:="CredReadW", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CredRead(target As String, type As Integer, flags As Integer, ByRef credential As IntPtr) As Boolean
        End Function

        <DllImport("advapi32.dll", EntryPoint:="CredDeleteW", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CredDelete(target As String, type As Integer, flags As Integer) As Boolean
        End Function

        <DllImport("advapi32.dll", EntryPoint:="CredEnumerateW", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CredEnumerate(filter As String, flags As Integer, ByRef count As Integer, ByRef credentials As IntPtr) As Boolean
        End Function

        <DllImport("advapi32.dll")>
        Private Shared Sub CredFree(buffer As IntPtr)
        End Sub

        ''' <summary>A credential name: letters, digits, spaces and . - _ @ (64 at most).</summary>
        Public Shared Function IsValidName(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) OrElse name.Length > 64 OrElse name.Trim() <> name Then Return False
            Return name.All(Function(c) Char.IsLetterOrDigit(c) OrElse c = " "c OrElse c = "."c OrElse c = "-"c OrElse c = "_"c OrElse c = "@"c)
        End Function

        ''' <summary>Creates or replaces a credential.</summary>
        Public Shared Sub Save(name As String, userName As String, password As String)
            If Not IsValidName(name) Then Throw New ArgumentException("Invalid credential name (letters, digits, spaces and . - _ @).", NameOf(name))
            If String.IsNullOrWhiteSpace(userName) Then Throw New ArgumentException("A user name is required.", NameOf(userName))
            Dim blob = Encoding.Unicode.GetBytes(If(password, String.Empty))
            Dim pBlob = Marshal.AllocHGlobal(Math.Max(1, blob.Length))
            Try
                Marshal.Copy(blob, 0, pBlob, blob.Length)
                Dim c As New CREDENTIAL With {
                    .Type = CRED_TYPE_GENERIC,
                    .TargetName = Prefix & name,
                    .Comment = "LiteSuite",
                    .CredentialBlobSize = CUInt(blob.Length),
                    .CredentialBlob = pBlob,
                    .Persist = CRED_PERSIST_LOCAL_MACHINE,
                    .UserName = userName.Trim()
                }
                If Not CredWrite(c, 0) Then Throw New Win32Exception(Marshal.GetLastWin32Error())
            Finally
                ' The password does not stay in unmanaged memory
                For i = 0 To blob.Length - 1
                    Marshal.WriteByte(pBlob, i, 0)
                Next
                Array.Clear(blob, 0, blob.Length)
                Marshal.FreeHGlobal(pBlob)
            End Try
        End Sub

        ''' <summary>The credential, or Nothing when it does not exist.</summary>
        Public Shared Function Read(name As String) As LiteSuiteCredential
            Dim p As IntPtr
            If Not CredRead(Prefix & name, CRED_TYPE_GENERIC, 0, p) Then
                Dim err = Marshal.GetLastWin32Error()
                If err = ERROR_NOT_FOUND Then Return Nothing
                Throw New Win32Exception(err)
            End If
            Try
                Dim c = Marshal.PtrToStructure(Of CREDENTIAL)(p)
                Dim password = If(c.CredentialBlobSize = 0, String.Empty, Marshal.PtrToStringUni(c.CredentialBlob, CInt(c.CredentialBlobSize \ 2)))
                Return New LiteSuiteCredential With {.Name = name, .UserName = c.UserName, .Password = password}
            Finally
                CredFree(p)
            End Try
        End Function

        ''' <summary>Deletes a credential; False when it did not exist.</summary>
        Public Shared Function Delete(name As String) As Boolean
            If CredDelete(Prefix & name, CRED_TYPE_GENERIC, 0) Then Return True
            Dim err = Marshal.GetLastWin32Error()
            If err = ERROR_NOT_FOUND Then Return False
            Throw New Win32Exception(err)
        End Function

        ''' <summary>Names and user names of the LiteSuite credentials (no password).</summary>
        Public Shared Function List() As List(Of LiteSuiteCredential)
            Dim result As New List(Of LiteSuiteCredential)
            Dim count As Integer
            Dim p As IntPtr
            If Not CredEnumerate(Prefix & "*", 0, count, p) Then
                Dim err = Marshal.GetLastWin32Error()
                If err = ERROR_NOT_FOUND Then Return result
                Throw New Win32Exception(err)
            End If
            Try
                For i = 0 To count - 1
                    Dim item = Marshal.ReadIntPtr(p, i * IntPtr.Size)
                    Dim c = Marshal.PtrToStructure(Of CREDENTIAL)(item)
                    If c.Type = CRED_TYPE_GENERIC AndAlso c.TargetName IsNot Nothing AndAlso c.TargetName.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) Then
                        result.Add(New LiteSuiteCredential With {.Name = c.TargetName.Substring(Prefix.Length), .UserName = c.UserName})
                    End If
                Next
            Finally
                CredFree(p)
            End Try
            Return result.OrderBy(Function(c) c.Name, StringComparer.OrdinalIgnoreCase).ToList()
        End Function
    End Class
End Namespace
