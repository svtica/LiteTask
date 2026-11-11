' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    RuntimeRequirement.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Reads the .NET shared frameworks an update package needs (its
'          LiteTask.runtimeconfig.json) and checks they are installed on this
'          machine before any file is replaced. Moving LiteTask to a new .NET
'          major version (8 -> 10) would otherwise leave a server whose desktop
'          runtime was not upgraded with an application and a service that no
'          longer start. Parsing and matching are pure so they can be unit-tested.
' -----------------------------------------------------------------------------
Imports System.Text.Json

Namespace LiteTask

    Public Class FrameworkReference
        Public Property Name As String
        Public Property Version As Version

        Public Overrides Function ToString() As String
            Return $"{Name} {Version}"
        End Function
    End Class

    Public NotInheritable Class RuntimeRequirement
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Frameworks listed in a runtimeconfig.json ("framework" or "frameworks").
        ''' Empty for a self-contained application (no framework reference).
        ''' </summary>
        Public Shared Function ParseRuntimeConfig(json As String) As List(Of FrameworkReference)
            Dim result As New List(Of FrameworkReference)
            Using doc = JsonDocument.Parse(json, New JsonDocumentOptions With {
                .CommentHandling = JsonCommentHandling.Skip, .AllowTrailingCommas = True})
                Dim options As JsonElement
                If Not doc.RootElement.TryGetProperty("runtimeOptions", options) Then Return result

                Dim frameworkElement As JsonElement
                If options.TryGetProperty("framework", frameworkElement) Then AddFramework(result, frameworkElement)
                Dim frameworkArray As JsonElement
                If options.TryGetProperty("frameworks", frameworkArray) AndAlso frameworkArray.ValueKind = JsonValueKind.Array Then
                    For Each item In frameworkArray.EnumerateArray()
                        AddFramework(result, item)
                    Next
                End If
            End Using
            Return result
        End Function

        Private Shared Sub AddFramework(result As List(Of FrameworkReference), element As JsonElement)
            Dim nameElement As JsonElement
            Dim versionElement As JsonElement
            If element.TryGetProperty("name", nameElement) AndAlso element.TryGetProperty("version", versionElement) Then
                Dim parsed = ParseVersion(versionElement.GetString())
                If parsed IsNot Nothing Then
                    result.Add(New FrameworkReference With {.Name = nameElement.GetString(), .Version = parsed})
                End If
            End If
        End Sub

        ''' <summary>"10.0.14" -&gt; 10.0.14; prerelease suffixes ("11.0.0-rc.1") are ignored.</summary>
        Public Shared Function ParseVersion(value As String) As Version
            If String.IsNullOrWhiteSpace(value) Then Return Nothing
            Dim core = value.Split({"-"c, "+"c}, 2)(0)
            Dim parsed As Version = Nothing
            Return If(Version.TryParse(core, parsed), parsed, Nothing)
        End Function

        ''' <summary>
        ''' True when one installed version can run the requirement under the default
        ''' roll-forward policy (Minor): same major version, and not older.
        ''' </summary>
        Public Shared Function IsSatisfied(required As FrameworkReference, installedVersions As IEnumerable(Of Version)) As Boolean
            Return installedVersions IsNot Nothing AndAlso
                   installedVersions.Any(Function(v) v.Major = required.Version.Major AndAlso v >= required.Version)
        End Function

        ''' <summary>
        ''' Versions of a shared framework installed under the .NET roots of this
        ''' machine (%ProgramFiles%\dotnet, DOTNET_ROOT), from the folder names in
        ''' shared\&lt;framework&gt;\.
        ''' </summary>
        Public Shared Function GetInstalledVersions(frameworkName As String) As List(Of Version)
            Dim result As New List(Of Version)
            For Each root In GetDotnetRoots()
                Dim folder = Path.Combine(root, "shared", frameworkName)
                If Not Directory.Exists(folder) Then Continue For
                For Each versionDir In Directory.GetDirectories(folder)
                    Dim parsed = ParseVersion(Path.GetFileName(versionDir))
                    If parsed IsNot Nothing Then result.Add(parsed)
                Next
            Next
            Return result
        End Function

        Private Shared Function GetDotnetRoots() As List(Of String)
            Dim roots As New List(Of String)
            Dim fromEnv = Environment.GetEnvironmentVariable("DOTNET_ROOT")
            If Not String.IsNullOrWhiteSpace(fromEnv) Then roots.Add(fromEnv)
            ' The runtime this process runs on lives in <root>\shared\<framework>\<version>\
            Try
                Dim runtimeDir = Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()
                Dim root = Directory.GetParent(runtimeDir.TrimEnd(Path.DirectorySeparatorChar))?.Parent?.Parent
                If root IsNot Nothing Then roots.Add(root.FullName)
            Catch
            End Try
            roots.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet"))
            Return roots.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        End Function

        ''' <summary>
        ''' Frameworks required by the update package in <paramref name="packageDir"/>
        ''' that are not installed here. Empty when everything is present or when the
        ''' package carries no runtimeconfig.json.
        ''' </summary>
        Public Shared Function FindMissingFrameworks(packageDir As String, executableName As String) As List(Of FrameworkReference)
            Dim configName = Path.GetFileNameWithoutExtension(executableName) & ".runtimeconfig.json"
            Dim configPath = Directory.GetFiles(packageDir, configName, SearchOption.AllDirectories).
                OrderBy(Function(p) p.Length).FirstOrDefault()
            If configPath Is Nothing Then Return New List(Of FrameworkReference)

            Return ParseRuntimeConfig(File.ReadAllText(configPath)).
                Where(Function(f) Not IsSatisfied(f, GetInstalledVersions(f.Name))).
                ToList()
        End Function
    End Class

End Namespace
