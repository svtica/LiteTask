' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    RuntimeRequirementTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Unit tests for LiteTask.RuntimeRequirement: an update built on a new
'          .NET major version must be refused when its desktop runtime is not
'          installed, and accepted under the default roll-forward policy.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class RuntimeRequirementTests

    Private Const WinFormsConfig As String = "{
  ""runtimeOptions"": {
    ""tfm"": ""net10.0"",
    ""frameworks"": [
      { ""name"": ""Microsoft.NETCore.App"", ""version"": ""10.0.0"" },
      { ""name"": ""Microsoft.WindowsDesktop.App"", ""version"": ""10.0.0"" }
    ],
    ""configProperties"": { ""System.GC.Server"": false }
  }
}"

    Private Shared Function Req(name As String, version As String) As FrameworkReference
        Return New FrameworkReference With {.Name = name, .Version = New Version(version)}
    End Function

    <TestMethod>
    Public Sub ParseRuntimeConfig_ReadsFrameworksArray()
        Dim frameworks = RuntimeRequirement.ParseRuntimeConfig(WinFormsConfig)
        Assert.AreEqual(2, frameworks.Count)
        Assert.AreEqual("Microsoft.WindowsDesktop.App", frameworks(1).Name)
        Assert.AreEqual(New Version(10, 0, 0), frameworks(1).Version)
    End Sub

    <TestMethod>
    Public Sub ParseRuntimeConfig_ReadsSingleFrameworkAndSelfContained()
        Dim oneFramework = RuntimeRequirement.ParseRuntimeConfig("{""runtimeOptions"":{""framework"":{""name"":""Microsoft.NETCore.App"",""version"":""8.0.0""}}}")
        Assert.AreEqual(1, oneFramework.Count)
        Assert.AreEqual(8, oneFramework(0).Version.Major)

        ' Self-contained: no framework reference, nothing to check
        Assert.AreEqual(0, RuntimeRequirement.ParseRuntimeConfig("{""runtimeOptions"":{""tfm"":""net10.0""}}").Count)
    End Sub

    <TestMethod>
    Public Sub ParseVersion_IgnoresPrereleaseSuffix()
        Assert.AreEqual(New Version(11, 0, 0), RuntimeRequirement.ParseVersion("11.0.0-rc.1.26425.128"))
        Assert.IsNull(RuntimeRequirement.ParseVersion("not-a-version"))
        Assert.IsNull(RuntimeRequirement.ParseVersion(""))
    End Sub

    <TestMethod>
    Public Sub IsSatisfied_SameMajorNotOlder()
        Dim required = Req("Microsoft.WindowsDesktop.App", "10.0.0")
        Assert.IsTrue(RuntimeRequirement.IsSatisfied(required, {New Version(10, 0, 14)}))
        Assert.IsTrue(RuntimeRequirement.IsSatisfied(required, {New Version(8, 0, 22), New Version(10, 0, 0)}))
    End Sub

    <TestMethod>
    Public Sub IsSatisfied_OlderOrNewerMajor_IsMissing()
        Dim required = Req("Microsoft.WindowsDesktop.App", "10.0.0")
        ' Server still on the .NET 8 desktop runtime: the update must be refused
        Assert.IsFalse(RuntimeRequirement.IsSatisfied(required, {New Version(8, 0, 22)}))
        ' Default roll-forward (Minor) does not move to another major version
        Assert.IsFalse(RuntimeRequirement.IsSatisfied(required, {New Version(11, 0, 0)}))
        Assert.IsFalse(RuntimeRequirement.IsSatisfied(Req("X", "10.0.5"), {New Version(10, 0, 4)}))
        Assert.IsFalse(RuntimeRequirement.IsSatisfied(required, New Version() {}))
        Assert.IsFalse(RuntimeRequirement.IsSatisfied(required, Nothing))
    End Sub

End Class
