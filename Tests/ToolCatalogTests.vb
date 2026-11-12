' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    ToolCatalogTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Unit tests for LiteTask.ToolCatalog: optional tools are downloaded
'          from GitHub releases only (HTTPS github.com release assets), the
'          right asset is picked, sqlcmd stays pinned and the tool selection
'          survives an upgrade from a package that shipped the tools.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class ToolCatalogTests

    Private Shared Function Asset(name As String, url As String) As KeyValuePair(Of String, String)
        Return New KeyValuePair(Of String, String)(name, url)
    End Function

    <TestMethod>
    Public Sub Catalog_SqlcmdIsPinnedInstalledAtToolsRootAndAutoInstalled()
        Dim sqlcmd = ToolCatalog.Find("SQLCMD")
        Assert.IsNotNull(sqlcmd)
        Assert.AreEqual("sqlcmd.exe", sqlcmd.ExecutableName)
        Assert.AreEqual("", sqlcmd.InstallFolder, "TaskRunner runs LiteTaskData\tools\sqlcmd.exe")
        Assert.IsFalse(String.IsNullOrEmpty(sqlcmd.PinnedTag), "sqlcmd arguments are pinned by tests; its version must not float")
        Assert.IsTrue(sqlcmd.AutoInstallWhenNeeded)
        Assert.IsTrue(sqlcmd.SelectedByDefault)
    End Sub

    <TestMethod>
    Public Sub GetReleaseApiUrl_PinnedTagOrLatest()
        Assert.AreEqual("https://api.github.com/repos/microsoft/go-sqlcmd/releases/tags/v1.8.2",
                        ToolCatalog.GetReleaseApiUrl(ToolCatalog.Find(ToolCatalog.SqlCmdId)))
        Assert.AreEqual("https://api.github.com/repos/svtica/LitePM/releases/latest",
                        ToolCatalog.GetReleaseApiUrl(ToolCatalog.Find(ToolCatalog.LitePMId)))
    End Sub

    <TestMethod>
    Public Sub IsSafeDownloadUrl_OnlyHttpsGithubReleaseAssets()
        Assert.IsTrue(ToolCatalog.IsSafeDownloadUrl("https://github.com/svtica/LitePM/releases/download/v1.2/LitePM-1.2-win-x64.zip"))
        Assert.IsFalse(ToolCatalog.IsSafeDownloadUrl("http://github.com/svtica/LitePM/releases/download/v1.2/LitePM-1.2-win-x64.zip"))
        Assert.IsFalse(ToolCatalog.IsSafeDownloadUrl("https://github.com.evil.example/svtica/LitePM/releases/download/v1.2/x.zip"))
        Assert.IsFalse(ToolCatalog.IsSafeDownloadUrl("https://example.com/releases/download/x.zip"))
        Assert.IsFalse(ToolCatalog.IsSafeDownloadUrl("https://github.com/svtica/LitePM/archive/refs/heads/main.zip"))
        Assert.IsFalse(ToolCatalog.IsSafeDownloadUrl("not a url"))
    End Sub

    <TestMethod>
    Public Sub SelectAsset_PicksMatchingAssetAndSkipsUnsafeUrls()
        Dim litePm = ToolCatalog.Find(ToolCatalog.LitePMId)
        Dim assets = New List(Of KeyValuePair(Of String, String)) From {
            Asset("Source code.zip", "https://github.com/svtica/LitePM/archive/v1.2.zip"),
            Asset("LitePM-1.2-win-x64.zip", "https://evil.example/LitePM-1.2-win-x64.zip"),
            Asset("LitePM-1.2-win-x64.zip", "https://github.com/svtica/LitePM/releases/download/v1.2/LitePM-1.2-win-x64.zip")
        }
        Dim selected = ToolCatalog.SelectAsset(litePm, assets)
        Assert.IsTrue(selected.HasValue)
        ' Plain "LitePM.zip" (v1.0 naming) is accepted too
        Assert.IsTrue(ToolCatalog.SelectAsset(litePm, New List(Of KeyValuePair(Of String, String)) From {
            Asset("LitePM.zip", "https://github.com/svtica/LitePM/releases/download/v1.0/LitePM.zip")}).HasValue)
        Assert.AreEqual("https://github.com/svtica/LitePM/releases/download/v1.2/LitePM-1.2-win-x64.zip", selected.Value.Value)
    End Sub

    <TestMethod>
    Public Sub SelectAsset_SqlcmdIgnoresOtherPlatforms()
        Dim sqlcmd = ToolCatalog.Find(ToolCatalog.SqlCmdId)
        Dim assets = New List(Of KeyValuePair(Of String, String)) From {
            Asset("sqlcmd-linux-amd64.tar.bz2", "https://github.com/microsoft/go-sqlcmd/releases/download/v1.8.2/sqlcmd-linux-amd64.tar.bz2"),
            Asset("sqlcmd-windows-arm64.zip", "https://github.com/microsoft/go-sqlcmd/releases/download/v1.8.2/sqlcmd-windows-arm64.zip"),
            Asset("sqlcmd-windows-amd64.zip", "https://github.com/microsoft/go-sqlcmd/releases/download/v1.8.2/sqlcmd-windows-amd64.zip")
        }
        Assert.AreEqual("sqlcmd-windows-amd64.zip", ToolCatalog.SelectAsset(sqlcmd, assets).Value.Key)
    End Sub

    <TestMethod>
    Public Sub SelectAsset_NoMatch_ReturnsNothing()
        Dim liteDeploy = ToolCatalog.Find(ToolCatalog.LiteDeployId)
        Assert.IsFalse(ToolCatalog.SelectAsset(liteDeploy, New List(Of KeyValuePair(Of String, String))).HasValue)
        Assert.IsFalse(ToolCatalog.SelectAsset(liteDeploy, Nothing).HasValue)
    End Sub

    <TestMethod>
    Public Sub Selection_RoundTripKeepsKnownIdsInCatalogOrder()
        Dim parsed = ToolCatalog.ParseSelection(" litepm ; unknown;SQLCMD,")
        Assert.AreEqual(2, parsed.Count)
        Assert.AreEqual("sqlcmd;LitePM", ToolCatalog.FormatSelection(parsed))
        Assert.AreEqual(0, ToolCatalog.ParseSelection("").Count)
    End Sub

    <TestMethod>
    Public Sub DefaultSelection_KeepsToolsAlreadyInstalledByAnOlderPackage()
        Dim selection = ToolCatalog.DefaultSelection(Function(t) t.Id = ToolCatalog.LitePMId)
        Assert.IsTrue(selection.Contains(ToolCatalog.SqlCmdId))
        Assert.IsTrue(selection.Contains(ToolCatalog.LitePMId))
        Assert.IsFalse(selection.Contains(ToolCatalog.LiteDeployId))
    End Sub

End Class
