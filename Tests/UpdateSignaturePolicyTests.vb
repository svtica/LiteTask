' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    UpdateSignaturePolicyTests.vb
' Author:  LiteTask contributors
' Date:    2026-10-28
' Purpose: Unit tests for the self-update signature policy (which packages
'          UpdateManager may apply) and for thumbprint normalization.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class UpdateSignaturePolicyTests

    Private Const Pinned As String = "EC2B412C4025C2F1C8DB512CEC66632E65C017F0"
    Private Const Other As String = "0123456789ABCDEF0123456789ABCDEF01234567"

    Private Shared Function Signed(thumbprint As String, Optional status As AuthenticodeStatus = AuthenticodeStatus.ValidUntrustedRoot) As AuthenticodeResult
        Return New AuthenticodeResult With {
            .Status = status,
            .SignerThumbprint = thumbprint,
            .SignerSubject = "CN=Test"
        }
    End Function

    Private Shared Function WithStatus(status As AuthenticodeStatus) As AuthenticodeResult
        Return New AuthenticodeResult With {.Status = status, .ErrorCode = &H800B0100}
    End Function

    <TestMethod>
    Public Sub Evaluate_SignedByPinnedPublisher_Accepted()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Signed(Pinned), Signed(Pinned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.Accepted, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_SignedWithTrustedRoot_Accepted()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Signed(Pinned, AuthenticodeStatus.Valid), Nothing, {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.Accepted, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_ThumbprintComparisonIgnoresCaseAndSeparators()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Signed("ec:2b:41:2c:40:25:c2:f1:c8:db:51:2c:ec:66:63:2e:65:c0:17:f0"), Nothing, {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.Accepted, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_SignedByOtherPublisher_Rejected()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Signed(Other), Signed(Pinned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedPublisherMismatch, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_SignedByOtherPublisher_RejectedEvenOnDevBuild()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Signed(Other), WithStatus(AuthenticodeStatus.NotSigned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedPublisherMismatch, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_UnsignedPackageOnSignedInstall_Rejected()
        Dim verdict = UpdateSignaturePolicy.Evaluate(WithStatus(AuthenticodeStatus.NotSigned), Signed(Pinned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedUnsigned, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_UnsignedPackageWhenRunningSignatureUnknown_Rejected()
        Dim verdict = UpdateSignaturePolicy.Evaluate(WithStatus(AuthenticodeStatus.NotSigned), Nothing, {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedUnsigned, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_UnsignedPackageOnUnsignedDevBuild_AcceptedWithWarning()
        Dim verdict = UpdateSignaturePolicy.Evaluate(WithStatus(AuthenticodeStatus.NotSigned), WithStatus(AuthenticodeStatus.NotSigned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.AcceptedUnsignedDevBuild, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_TamperedPackage_RejectedEvenOnDevBuild()
        Dim verdict = UpdateSignaturePolicy.Evaluate(WithStatus(AuthenticodeStatus.HashMismatch), WithStatus(AuthenticodeStatus.NotSigned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedTampered, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_InvalidSignature_Rejected()
        Dim verdict = UpdateSignaturePolicy.Evaluate(WithStatus(AuthenticodeStatus.Invalid), Signed(Pinned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedInvalidSignature, verdict)
    End Sub

    <TestMethod>
    Public Sub Evaluate_NothingPackage_Rejected()
        Dim verdict = UpdateSignaturePolicy.Evaluate(Nothing, Signed(Pinned), {Pinned})
        Assert.AreEqual(UpdateSignatureVerdict.RejectedInvalidSignature, verdict)
    End Sub

    <TestMethod>
    Public Sub ExpectedThumbprints_IncludesRunningSignerWhenValid()
        Dim expected = UpdateSignaturePolicy.ExpectedThumbprints(Pinned, Signed(Other))
        CollectionAssert.AreEquivalent(New String() {Pinned, Other}, expected)
    End Sub

    <TestMethod>
    Public Sub ExpectedThumbprints_IgnoresRunningSignerWhenNotValid()
        Dim expected = UpdateSignaturePolicy.ExpectedThumbprints(Pinned, Signed(Other, AuthenticodeStatus.HashMismatch))
        CollectionAssert.AreEquivalent(New String() {Pinned}, expected)
    End Sub

    <TestMethod>
    Public Sub ExpectedThumbprints_DoesNotDuplicatePinned()
        Dim expected = UpdateSignaturePolicy.ExpectedThumbprints(Pinned, Signed(Pinned.ToLowerInvariant()))
        CollectionAssert.AreEquivalent(New String() {Pinned}, expected)
    End Sub

    <TestMethod>
    Public Sub NormalizeThumbprint_StripsSeparatorsAndUppercases()
        Assert.AreEqual("ABCDEF01", AuthenticodeVerifier.NormalizeThumbprint(" ab:cd ef-01 "))
        Assert.AreEqual(String.Empty, AuthenticodeVerifier.NormalizeThumbprint(Nothing))
    End Sub

    <TestMethod>
    Public Sub ThumbprintMatches_EmptyThumbprintNeverMatches()
        Assert.IsFalse(AuthenticodeVerifier.ThumbprintMatches(Nothing, {Pinned}))
        Assert.IsFalse(AuthenticodeVerifier.ThumbprintMatches(Pinned, Nothing))
        Assert.IsTrue(AuthenticodeVerifier.ThumbprintMatches(Pinned, {Other, Pinned}))
    End Sub
End Class
