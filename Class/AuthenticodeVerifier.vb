' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    AuthenticodeVerifier.vb
' Author:  LiteTask contributors
' Date:    2026-10-28
' Purpose: Authenticode verification of files extracted from an update package.
'          Uses WinVerifyTrust to confirm that the embedded signature matches
'          the file content, then reads the signer certificate so the caller
'          can pin the publisher by thumbprint.
'
'          The publisher certificate (res\certificate.cer, CN=SVTI.ca) is
'          self-signed, so the chain normally terminates in a root that is not
'          in the machine store. CERT_E_UNTRUSTEDROOT is therefore reported as
'          a *valid* signature (the digest was verified) and the identity check
'          is left to the thumbprint comparison performed by the caller.
' -----------------------------------------------------------------------------
Imports System.Security.Cryptography.X509Certificates

Namespace LiteTask

    Public Enum AuthenticodeStatus
        ''' <summary>Signature matches the file and the chain ends in a trusted root.</summary>
        Valid
        ''' <summary>Signature matches the file, but the chain ends in a root that is not trusted on this machine (self-signed publisher).</summary>
        ValidUntrustedRoot
        ''' <summary>The file carries no Authenticode signature.</summary>
        NotSigned
        ''' <summary>The file carries a signature that does not match its content (tampered or corrupted).</summary>
        HashMismatch
        ''' <summary>Any other failure (expired, revoked, explicitly distrusted, chain error, file missing...).</summary>
        Invalid
    End Enum

    Public Class AuthenticodeResult
        Public Property Status As AuthenticodeStatus = AuthenticodeStatus.Invalid
        ''' <summary>HRESULT returned by WinVerifyTrust (0 on success).</summary>
        Public Property ErrorCode As Integer
        ''' <summary>SHA-1 thumbprint of the signer certificate, upper-case hex without separators. Nothing when the signature is not valid.</summary>
        Public Property SignerThumbprint As String
        ''' <summary>Distinguished name of the signer certificate. Nothing when the signature is not valid.</summary>
        Public Property SignerSubject As String

        ''' <summary>True when the signature was cryptographically verified against the file content.</summary>
        Public ReadOnly Property IsSignatureValid As Boolean
            Get
                Return Status = AuthenticodeStatus.Valid OrElse Status = AuthenticodeStatus.ValidUntrustedRoot
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"{Status} (0x{ErrorCode:X8}) signer={If(SignerSubject, "<none>")} thumbprint={If(SignerThumbprint, "<none>")}"
        End Function
    End Class

    Public Class AuthenticodeVerifier

#Region "Win32"
        Private Const WTD_UI_NONE As UInteger = 2
        Private Const WTD_REVOKE_NONE As UInteger = 0
        Private Const WTD_CHOICE_FILE As UInteger = 1
        Private Const WTD_STATEACTION_VERIFY As UInteger = 1
        Private Const WTD_STATEACTION_CLOSE As UInteger = 2
        Private Const WTD_REVOCATION_CHECK_NONE As UInteger = &H10
        Private Const WTD_UICONTEXT_EXECUTE As UInteger = 0

        Public Const ERROR_SUCCESS As Integer = 0
        Public Const TRUST_E_NOSIGNATURE As Integer = &H800B0100
        Public Const TRUST_E_BAD_DIGEST As Integer = &H80096010
        Public Const CERT_E_UNTRUSTEDROOT As Integer = &H800B0109

        ' WINTRUST_ACTION_GENERIC_VERIFY_V2
        Private Shared ReadOnly WintrustActionGenericVerifyV2 As New Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE")

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure WINTRUST_FILE_INFO
            Public cbStruct As UInteger
            <MarshalAs(UnmanagedType.LPWStr)> Public pcwszFilePath As String
            Public hFile As IntPtr
            Public pgKnownSubject As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINTRUST_DATA
            Public cbStruct As UInteger
            Public pPolicyCallbackData As IntPtr
            Public pSIPClientData As IntPtr
            Public dwUIChoice As UInteger
            Public fdwRevocationChecks As UInteger
            Public dwUnionChoice As UInteger
            Public pFile As IntPtr
            Public dwStateAction As UInteger
            Public hWVTStateData As IntPtr
            Public pwszURLReference As IntPtr
            Public dwProvFlags As UInteger
            Public dwUIContext As UInteger
            Public pSignatureSettings As IntPtr
        End Structure

        <DllImport("wintrust.dll", ExactSpelling:=True)>
        Private Shared Function WinVerifyTrust(hwnd As IntPtr, ByRef pgActionID As Guid, pWVTData As IntPtr) As Integer
        End Function
#End Region

        ''' <summary>
        ''' Verifies the Authenticode signature embedded in a file (PE executable, PowerShell script, ...).
        ''' Never throws: every failure is reported through <see cref="AuthenticodeResult.Status"/>.
        ''' </summary>
        Public Shared Function Verify(filePath As String) As AuthenticodeResult
            Dim result As New AuthenticodeResult()

            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                result.Status = AuthenticodeStatus.Invalid
                result.ErrorCode = TRUST_E_NOSIGNATURE
                Return result
            End If

            Try
                result.ErrorCode = CallWinVerifyTrust(Path.GetFullPath(filePath))
            Catch
                result.Status = AuthenticodeStatus.Invalid
                Return result
            End Try

            Select Case result.ErrorCode
                Case ERROR_SUCCESS
                    result.Status = AuthenticodeStatus.Valid
                Case CERT_E_UNTRUSTEDROOT
                    result.Status = AuthenticodeStatus.ValidUntrustedRoot
                Case TRUST_E_NOSIGNATURE
                    result.Status = AuthenticodeStatus.NotSigned
                Case TRUST_E_BAD_DIGEST
                    result.Status = AuthenticodeStatus.HashMismatch
                Case Else
                    result.Status = AuthenticodeStatus.Invalid
            End Select

            If result.IsSignatureValid Then
                Try
                    Using signer = X509Certificate.CreateFromSignedFile(filePath)
                        Using cert As New X509Certificate2(signer)
                            result.SignerThumbprint = NormalizeThumbprint(cert.Thumbprint)
                            result.SignerSubject = cert.Subject
                        End Using
                    End Using
                Catch
                    ' A verified signature whose signer cannot be read is not usable for pinning: fail closed.
                    result.Status = AuthenticodeStatus.Invalid
                    result.SignerThumbprint = Nothing
                    result.SignerSubject = Nothing
                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' Normalizes a thumbprint for comparison: upper-case hex digits only (separators and spaces removed).
        ''' Returns an empty string for Nothing.
        ''' </summary>
        Public Shared Function NormalizeThumbprint(value As String) As String
            If String.IsNullOrEmpty(value) Then Return String.Empty
            Return New String(value.Where(Function(c) Uri.IsHexDigit(c)).ToArray()).ToUpperInvariant()
        End Function

        ''' <summary>
        ''' Returns True when <paramref name="thumbprint"/> matches one of <paramref name="expected"/> (after normalization).
        ''' </summary>
        Public Shared Function ThumbprintMatches(thumbprint As String, expected As IEnumerable(Of String)) As Boolean
            Dim normalized = NormalizeThumbprint(thumbprint)
            If normalized.Length = 0 OrElse expected Is Nothing Then Return False
            Return expected.Any(Function(e) NormalizeThumbprint(e) = normalized)
        End Function

        Private Shared Function CallWinVerifyTrust(filePath As String) As Integer
            Dim actionId = WintrustActionGenericVerifyV2
            Dim invalidHandle As New IntPtr(-1)

            Dim fileInfo As New WINTRUST_FILE_INFO With {
                .pcwszFilePath = filePath,
                .hFile = IntPtr.Zero,
                .pgKnownSubject = IntPtr.Zero
            }
            fileInfo.cbStruct = CUInt(Marshal.SizeOf(fileInfo))

            Dim pFileInfo As IntPtr = IntPtr.Zero
            Dim pTrustData As IntPtr = IntPtr.Zero
            Try
                pFileInfo = Marshal.AllocHGlobal(Marshal.SizeOf(fileInfo))
                Marshal.StructureToPtr(fileInfo, pFileInfo, False)

                Dim trustData As New WINTRUST_DATA With {
                    .pPolicyCallbackData = IntPtr.Zero,
                    .pSIPClientData = IntPtr.Zero,
                    .dwUIChoice = WTD_UI_NONE,
                    .fdwRevocationChecks = WTD_REVOKE_NONE,
                    .dwUnionChoice = WTD_CHOICE_FILE,
                    .pFile = pFileInfo,
                    .dwStateAction = WTD_STATEACTION_VERIFY,
                    .hWVTStateData = IntPtr.Zero,
                    .pwszURLReference = IntPtr.Zero,
                    .dwProvFlags = WTD_REVOCATION_CHECK_NONE,
                    .dwUIContext = WTD_UICONTEXT_EXECUTE,
                    .pSignatureSettings = IntPtr.Zero
                }
                trustData.cbStruct = CUInt(Marshal.SizeOf(trustData))

                pTrustData = Marshal.AllocHGlobal(Marshal.SizeOf(trustData))
                Marshal.StructureToPtr(trustData, pTrustData, False)

                Dim status = WinVerifyTrust(invalidHandle, actionId, pTrustData)

                ' Release the state data allocated by the verify call.
                trustData = Marshal.PtrToStructure(Of WINTRUST_DATA)(pTrustData)
                trustData.dwStateAction = WTD_STATEACTION_CLOSE
                Marshal.StructureToPtr(trustData, pTrustData, False)
                WinVerifyTrust(invalidHandle, actionId, pTrustData)

                Return status
            Finally
                If pFileInfo <> IntPtr.Zero Then
                    Marshal.DestroyStructure(pFileInfo, GetType(WINTRUST_FILE_INFO))
                    Marshal.FreeHGlobal(pFileInfo)
                End If
                If pTrustData <> IntPtr.Zero Then
                    Marshal.FreeHGlobal(pTrustData)
                End If
            End Try
        End Function
    End Class

    ''' <summary>Outcome of <see cref="UpdateSignaturePolicy.Evaluate"/>.</summary>
    Public Enum UpdateSignatureVerdict
        ''' <summary>The package is signed by an expected publisher: apply it.</summary>
        Accepted
        ''' <summary>The package is unsigned, but so is the running executable (development build): apply it with a warning.</summary>
        AcceptedUnsignedDevBuild
        ''' <summary>The package is unsigned while the running executable is signed: refuse it.</summary>
        RejectedUnsigned
        ''' <summary>The package signature does not match its content: refuse it.</summary>
        RejectedTampered
        ''' <summary>The package is validly signed, but not by an expected publisher: refuse it.</summary>
        RejectedPublisherMismatch
        ''' <summary>The package signature could not be validated for another reason: refuse it.</summary>
        RejectedInvalidSignature
    End Enum

    ''' <summary>
    ''' Decides whether an update package may be applied, from the Authenticode result of the
    ''' extracted LiteTask.exe, the Authenticode result of the running LiteTask.exe and the set
    ''' of expected publisher thumbprints.
    ''' </summary>
    Public Module UpdateSignaturePolicy

        Public Function Evaluate(package As AuthenticodeResult, running As AuthenticodeResult, expectedThumbprints As IEnumerable(Of String)) As UpdateSignatureVerdict
            If package Is Nothing Then Return UpdateSignatureVerdict.RejectedInvalidSignature

            If package.IsSignatureValid Then
                If AuthenticodeVerifier.ThumbprintMatches(package.SignerThumbprint, expectedThumbprints) Then
                    Return UpdateSignatureVerdict.Accepted
                End If
                Return UpdateSignatureVerdict.RejectedPublisherMismatch
            End If

            Select Case package.Status
                Case AuthenticodeStatus.NotSigned
                    ' Only an unsigned (development) build may install an unsigned package.
                    ' When the running executable cannot be assessed, fail closed.
                    If running IsNot Nothing AndAlso running.Status = AuthenticodeStatus.NotSigned Then
                        Return UpdateSignatureVerdict.AcceptedUnsignedDevBuild
                    End If
                    Return UpdateSignatureVerdict.RejectedUnsigned
                Case AuthenticodeStatus.HashMismatch
                    Return UpdateSignatureVerdict.RejectedTampered
                Case Else
                    Return UpdateSignatureVerdict.RejectedInvalidSignature
            End Select
        End Function

        ''' <summary>
        ''' Builds the set of accepted publisher thumbprints: the pinned release certificate plus,
        ''' when the running executable carries a valid signature, its own signer.
        ''' </summary>
        Public Function ExpectedThumbprints(pinnedThumbprint As String, running As AuthenticodeResult) As List(Of String)
            Dim list As New List(Of String)
            Dim pinned = AuthenticodeVerifier.NormalizeThumbprint(pinnedThumbprint)
            If pinned.Length > 0 Then list.Add(pinned)
            If running IsNot Nothing AndAlso running.IsSignatureValid Then
                Dim own = AuthenticodeVerifier.NormalizeThumbprint(running.SignerThumbprint)
                If own.Length > 0 AndAlso Not list.Contains(own) Then list.Add(own)
            End If
            Return list
        End Function
    End Module
End Namespace
