' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ToolCatalog.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Optional external tools that LiteTask downloads on demand from their
'          GitHub releases instead of shipping them in every package (sqlcmd
'          alone was 23 MB, ~85 % of LiteTaskData). Pure data + selection rules,
'          kept free of I/O so they can be unit-tested; ToolManager performs the
'          download and installation.
' -----------------------------------------------------------------------------
Imports System.Text.RegularExpressions

Namespace LiteTask

    Public Class ToolDefinition
        ''' <summary>Stable identifier stored in settings (ToolSettings/SelectedTools).</summary>
        Public Property Id As String
        Public Property DisplayName As String
        ''' <summary>Translation key of the description shown in Tools &gt; Manage Tools.</summary>
        Public Property DescriptionKey As String
        Public Property DefaultDescription As String
        ''' <summary>GitHub "owner/repo" publishing the tool as a release asset.</summary>
        Public Property Repository As String
        ''' <summary>Release tag to install; Nothing installs the latest release.</summary>
        Public Property PinnedTag As String
        ''' <summary>Regex (case-insensitive) matched against the release asset names (.zip or .exe).</summary>
        Public Property AssetPattern As String
        Public Property ExecutableName As String
        ''' <summary>
        ''' Folder under LiteTaskData\tools. Empty: only the executable is copied to the
        ''' tools root (sqlcmd.exe, whose path TaskRunner relies on). Otherwise the whole
        ''' release content is installed in that subfolder (LitePM and its DLLs).
        ''' </summary>
        Public Property InstallFolder As String = ""
        ''' <summary>Selected on a new installation.</summary>
        Public Property SelectedByDefault As Boolean
        ''' <summary>Installed automatically, without asking, the first time a task needs it.</summary>
        Public Property AutoInstallWhenNeeded As Boolean
    End Class

    Public NotInheritable Class ToolCatalog
        Private Sub New()
        End Sub

        Public Const SqlCmdId As String = "sqlcmd"
        Public Const LitePMId As String = "LitePM"
        Public Const LiteDeployId As String = "LiteDeploy"

        Public Shared ReadOnly Property All As IReadOnlyList(Of ToolDefinition) = New List(Of ToolDefinition) From {
            New ToolDefinition With {
                .Id = SqlCmdId,
                .DisplayName = "sqlcmd",
                .DescriptionKey = "Tools.Description.sqlcmd",
                .DefaultDescription = "Microsoft go-sqlcmd. Required by SQL tasks.",
                .Repository = "microsoft/go-sqlcmd",
                .PinnedTag = "v1.8.2",
                .AssetPattern = "^sqlcmd-windows-amd64\.zip$",
                .ExecutableName = "sqlcmd.exe",
                .InstallFolder = "",
                .SelectedByDefault = True,
                .AutoInstallWhenNeeded = True
            },
            New ToolDefinition With {
                .Id = LitePMId,
                .DisplayName = "LitePM",
                .DescriptionKey = "Tools.Description.LitePM",
                .DefaultDescription = "Lite Process Manager. Used by Tools > Monitor LiteTask Processes.",
                .Repository = "svtica/LitePM",
                .AssetPattern = "^LitePM-.*-win-x64\.zip$",
                .ExecutableName = "LitePM.exe",
                .InstallFolder = "LitePM"
            },
            New ToolDefinition With {
                .Id = LiteDeployId,
                .DisplayName = "LiteDeploy",
                .DescriptionKey = "Tools.Description.LiteDeploy",
                .DefaultDescription = "Remote deployment and management utility (standalone).",
                .Repository = "svtica/LiteDeploy",
                .AssetPattern = "^LiteDeploy.*\.(zip|exe)$",
                .ExecutableName = "LiteDeploy.exe",
                .InstallFolder = "LiteDeploy"
            }
        }

        Public Shared Function Find(id As String) As ToolDefinition
            Return All.FirstOrDefault(Function(t) String.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>GitHub REST URL of the release to install (pinned tag or latest).</summary>
        Public Shared Function GetReleaseApiUrl(tool As ToolDefinition) As String
            If String.IsNullOrWhiteSpace(tool.PinnedTag) Then
                Return $"https://api.github.com/repos/{tool.Repository}/releases/latest"
            End If
            Return $"https://api.github.com/repos/{tool.Repository}/releases/tags/{Uri.EscapeDataString(tool.PinnedTag)}"
        End Function

        ''' <summary>Only HTTPS downloads from github.com release assets are accepted.</summary>
        Public Shared Function IsSafeDownloadUrl(url As String) As Boolean
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(url, UriKind.Absolute, uri) Then Return False
            Return uri.Scheme = Uri.UriSchemeHttps AndAlso
                   String.Equals(uri.Host, "github.com", StringComparison.OrdinalIgnoreCase) AndAlso
                   uri.AbsolutePath.Contains("/releases/download/")
        End Function

        ''' <summary>
        ''' First release asset whose name matches the tool's pattern and whose URL is a
        ''' github.com HTTPS release download. Nothing when the release has none.
        ''' </summary>
        ''' <param name="assets">(name, browser_download_url) pairs from the GitHub API.</param>
        Public Shared Function SelectAsset(tool As ToolDefinition, assets As IEnumerable(Of KeyValuePair(Of String, String))) As KeyValuePair(Of String, String)?
            If assets Is Nothing Then Return Nothing
            Dim pattern As New Regex(tool.AssetPattern, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)
            For Each asset In assets
                If asset.Key IsNot Nothing AndAlso pattern.IsMatch(asset.Key) AndAlso IsSafeDownloadUrl(asset.Value) Then
                    Return asset
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>Parses ToolSettings/SelectedTools ("sqlcmd;LitePM"), keeping known ids only.</summary>
        Public Shared Function ParseSelection(value As String) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(value) Then Return result
            For Each part In value.Split({";"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                Dim tool = Find(part.Trim())
                If tool IsNot Nothing Then result.Add(tool.Id)
            Next
            Return result
        End Function

        Public Shared Function FormatSelection(ids As IEnumerable(Of String)) As String
            Return String.Join(";", All.Where(Function(t) ids.Contains(t.Id, StringComparer.OrdinalIgnoreCase)).Select(Function(t) t.Id))
        End Function

        ''' <summary>
        ''' Selection used when the setting has never been saved: the tools selected by
        ''' default plus every tool already present (upgrade from a version that shipped
        ''' them), so an update never "unselects" a tool someone uses.
        ''' </summary>
        Public Shared Function DefaultSelection(isInstalled As Func(Of ToolDefinition, Boolean)) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each tool In All
                If tool.SelectedByDefault OrElse (isInstalled IsNot Nothing AndAlso isInstalled(tool)) Then
                    result.Add(tool.Id)
                End If
            Next
            Return result
        End Function
    End Class

End Namespace
