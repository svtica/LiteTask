' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    UpdateManagerTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-17
' Purpose: Unit tests for the options UpdateManager gives LiteSuite.Common's
'          SelfUpdater (repository, package, executable, files of the user)
'          and for the UpdateInfo built from a check (the version comparison
'          that replaced UpdateManager.IsNewerVersion).
' -----------------------------------------------------------------------------
Imports System.Text.RegularExpressions
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask
Imports LiteSuite

<TestClass>
Public Class UpdateManagerTests

    Private Const ExePath As String = "C:\Apps\LiteTask\LiteTask.exe"
    Private Const TempDir As String = "C:\Apps\LiteTask\LiteTaskData\temp"
    Private Const AssetUrl As String = "https://github.com/svtica/LiteTask/releases/download/v1.3.0.7/LiteTask.zip"

    Private Shared Function Options() As SelfUpdaterOptions
        Return UpdateManager.CreateOptions("1.3.0.6", ExePath, TempDir)
    End Function

    Private Shared Function Asset(name As String, url As String) As GitHubReleaseAsset
        Return New GitHubReleaseAsset With {.Name = name, .Size = 1000, .BrowserDownloadUrl = url}
    End Function

    Private Shared Function Release(tag As String, ParamArray assets As GitHubReleaseAsset()) As GitHubRelease
        Return New GitHubRelease With {.TagName = tag, .Name = "LiteTask " & tag, .Assets = assets.ToList()}
    End Function

    <TestMethod>
    Public Sub CreateOptions_TargetsSvticaLiteTask()
        Dim o = Options()
        Assert.AreEqual("svtica", o.Owner)
        Assert.AreEqual("LiteTask", o.Repository)
        Assert.AreEqual("LiteTask", o.ApplicationName)
        Assert.AreEqual("LiteTask.exe", o.ExecutableName)
        Assert.AreEqual("LiteTask.zip", o.PackageName)
        Assert.AreEqual("1.3.0.6", o.CurrentVersion)
        Assert.AreEqual(ExePath, o.CurrentExecutablePath)
        Assert.AreEqual(TempDir, o.TempDirectory)
    End Sub

    <TestMethod>
    Public Sub CreateOptions_PinsTheLiteSuitePublisher()
        Assert.AreEqual(SelfUpdaterOptions.LiteSuiteThumbprint, Options().PinnedThumbprint)
        Assert.AreEqual("D21B5D39EA4C7BDC69C94C86FC4602E8A1CA15B4", UpdateManager.ExpectedPublisherThumbprint)
    End Sub

    <TestMethod>
    Public Sub PackagePattern_AcceptsTheReleaseAssetOnly()
        Dim pattern As New Regex(Options().PackagePattern, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
        ' Asset name built by .github/workflows/Release.yml
        Assert.IsTrue(pattern.IsMatch("LiteTask.zip"))
        Assert.IsTrue(pattern.IsMatch("LiteTask-1.3.0.7.zip"))
        Assert.IsFalse(pattern.IsMatch("LitePM.zip"))
        Assert.IsFalse(pattern.IsMatch("LiteTask.zip.sha256"))
        Assert.IsFalse(pattern.IsMatch("Source code.zip"))
    End Sub

    <TestMethod>
    Public Sub FindPackage_PicksTheLiteTaskZipOfTheRepository()
        Dim o = Options()
        Dim r = Release("v1.3.0.7",
                        Asset("Other.zip", "https://github.com/svtica/LiteTask/releases/download/v1.3.0.7/Other.zip"),
                        Asset("LiteTask.zip", AssetUrl))
        Dim found = UpdateDownloadPolicy.FindPackage(r, o.Owner, o.Repository, o.PackagePattern, o.PackageName)
        Assert.IsNotNull(found)
        Assert.AreEqual("LiteTask.zip", found.Name)
    End Sub

    <TestMethod>
    Public Sub FindPackage_RefusesAnAssetOutsideTheRepository()
        Dim o = Options()
        Dim r = Release("v1.3.0.7", Asset("LiteTask.zip", "https://example.com/svtica/LiteTask/releases/download/v1.3.0.7/LiteTask.zip"))
        Assert.IsNull(UpdateDownloadPolicy.FindPackage(r, o.Owner, o.Repository, o.PackagePattern, o.PackageName))
    End Sub

    <TestMethod>
    Public Sub PreservePatterns_KeepUserDataButNotShippedFiles()
        Dim patterns = Options().PreservePatterns
        CollectionAssert.Contains(patterns, "LiteTaskData\settings.xml")
        CollectionAssert.Contains(patterns, "LiteTaskData\logs\*")
        CollectionAssert.Contains(patterns, "LiteTaskData\tools\*")
        CollectionAssert.Contains(patterns, "ServiceCredentials.dat")
        CollectionAssert.Contains(patterns, "*.log")
        ' Shipped in the package and updated by each release: never kept
        For Each shipped In {"LiteTask.exe", "LiteTask.dll", "LiteTaskData\lang\fr.xml", "LiteTaskData\InstallModules.ps1", "LiteTask-Post-Build.ps1"}
            Dim shippedFile = shipped
            Assert.IsFalse(patterns.Any(Function(p) LikeWildcard(shippedFile, p)), $"{shippedFile} must be replaced by the update")
        Next
    End Sub

    <TestMethod>
    Public Sub SelfUpdater_AcceptsTheOptions()
        Using updater As New SelfUpdater(Options())
            Assert.AreEqual("C:\Apps\LiteTask", updater.Options.ApplicationDirectory)
            StringAssert.EndsWith(updater.LogPath, "LiteTask-update.log")
        End Using
    End Sub

    <TestMethod>
    Public Sub ToUpdateInfo_NewerReleaseWithPackage()
        Dim r = Release("v1.3.0.7", Asset("LiteTask.zip", AssetUrl))
        Dim check As New UpdateCheckResult With {.CurrentVersion = "1.3.0.6", .LatestVersion = "1.3.0.7", .Release = r, .Package = r.Assets(0)}
        Dim info = UpdateManager.ToUpdateInfo(check)
        Assert.IsTrue(info.IsUpdateAvailable)
        Assert.AreEqual("1.3.0.6", info.CurrentVersion)
        Assert.AreEqual("1.3.0.7", info.RemoteVersion)
        Assert.AreEqual(AssetUrl, info.DownloadUrl)
        Assert.AreEqual("LiteTask.zip", info.AssetName)
        Assert.AreEqual("LiteTask v1.3.0.7", info.ReleaseName)
        Assert.AreSame(check, info.Check)
    End Sub

    <TestMethod>
    Public Sub ToUpdateInfo_NewerReleaseWithoutPackage_ReportedWithoutDownloadUrl()
        ' MainForm then says that the release has no package
        Dim check As New UpdateCheckResult With {.CurrentVersion = "1.3.0.6", .LatestVersion = "1.4.0.0", .Release = Release("v1.4.0.0")}
        Dim info = UpdateManager.ToUpdateInfo(check)
        Assert.IsTrue(info.IsUpdateAvailable)
        Assert.IsNull(info.DownloadUrl)
    End Sub

    <TestMethod>
    Public Sub ToUpdateInfo_SameOlderOrInvalidVersion_NoUpdate()
        ' Same rules as the former UpdateManager.IsNewerVersion
        Assert.IsFalse(InfoFor("1.3.0.6", "1.3.0.6").IsUpdateAvailable)
        Assert.IsFalse(InfoFor("1.3.0.6", "1.3.0.5").IsUpdateAvailable)
        Assert.IsFalse(InfoFor("1.3.0.6", "1.3").IsUpdateAvailable)
        Assert.IsFalse(InfoFor("1.3.0.6", "not-a-version").IsUpdateAvailable)
        Assert.IsTrue(InfoFor("1.3.0.6", "1.10").IsUpdateAvailable)
        Assert.IsTrue(InfoFor("1.3", "1.3.0.1").IsUpdateAvailable)
    End Sub

    Private Shared Function InfoFor(current As String, latest As String) As UpdateInfo
        Dim r = Release("v" & latest, Asset("LiteTask.zip", AssetUrl))
        Return UpdateManager.ToUpdateInfo(New UpdateCheckResult With {
            .CurrentVersion = current, .LatestVersion = latest, .Release = r, .Package = r.Assets(0)})
    End Function

    ''' <summary>PowerShell -like for the patterns used here ('*' matches anything, '\' included; case ignored).</summary>
    Private Shared Function LikeWildcard(value As String, pattern As String) As Boolean
        Dim expression As String = "^" & Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".") & "$"
        Return Regex.IsMatch(value, expression, RegexOptions.IgnoreCase)
    End Function
End Class
