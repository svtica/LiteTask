' -----------------------------------------------------------------------------
' LiteTask - part of LiteSuite. Released under The Unlicense (see LICENSE).
' File:    LocalizationTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-24
' Purpose: Tests of the language and display choices: English by default with
'          the texts of the code, French from LiteTaskData\lang\fr.xml, the
'          LITETASK_LANG override, Quebec typography, every key of the code
'          translated and given an English text, light or dark display.
' -----------------------------------------------------------------------------
Imports System.Drawing
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class LocalizationTests

    Private Shared Function RepoRoot() As String
        Dim dir = AppDomain.CurrentDomain.BaseDirectory
        While dir IsNot Nothing AndAlso Not File.Exists(Path.Combine(dir, "LiteTask.sln"))
            dir = Path.GetDirectoryName(dir)
        End While
        Assert.IsNotNull(dir, "LiteTask.sln not found above the test folder")
        Return dir
    End Function

    Private Shared Function LangFolder() As String
        Return Path.Combine(RepoRoot(), "LiteTaskData", "lang")
    End Function

    Private Shared Function FrenchKeys() As Dictionary(Of String, String)
        Dim doc As New XmlDocument()
        doc.Load(Path.Combine(LangFolder(), "fr.xml"))
        Dim keys As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each node As XmlNode In doc.SelectNodes("/translations/translation")
            keys(node.Attributes("key").Value) = node.InnerText
        Next
        Return keys
    End Function

    ' Source files of the application (not the tests)
    Private Shared Function Sources() As IEnumerable(Of String)
        Dim root = RepoRoot()
        Return Directory.GetFiles(Path.Combine(root, "Forms"), "*.vb").
            Concat(Directory.GetFiles(Path.Combine(root, "Class"), "*.vb")).
            Concat({Path.Combine(root, "Program.vb")})
    End Function

    <TestCleanup>
    Public Sub Cleanup()
        Environment.SetEnvironmentVariable("LITETASK_LANG", Nothing)
    End Sub

    <TestMethod>
    Public Sub English_KeepsTheEnglishTextOfTheCode()
        Dim manager As New TranslationManager(Nothing, Nothing, TranslationManager.English, LangFolder())
        Assert.AreEqual(TranslationManager.English, manager.Language)
        Assert.AreEqual("Cancel", manager.GetTranslation("_cancelButton.Text", "Cancel"))
    End Sub

    <TestMethod>
    Public Sub French_TranslatesWithQuebecTypography()
        Dim manager As New TranslationManager(Nothing, Nothing, TranslationManager.French, LangFolder())
        Assert.AreEqual(TranslationManager.French, manager.Language)
        Assert.AreEqual("Annuler", manager.GetTranslation("_cancelButton.Text", "Cancel"))
        Assert.AreEqual("Nom" & ChrW(160) & ":", manager.GetTranslation("_nameLabel.Text", "Name:"))
    End Sub

    <TestMethod>
    Public Sub French_MissingKey_FallsBackToEnglish()
        Dim manager As New TranslationManager(Nothing, Nothing, TranslationManager.French, LangFolder())
        Assert.AreEqual("Untranslated", manager.GetTranslation("No.Such.Key", "Untranslated"))
    End Sub

    <TestMethod>
    Public Sub Automatic_FollowsTheWindowsDisplayLanguage()
        Dim expected = If(Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName = "fr", TranslationManager.French, TranslationManager.English)
        Assert.AreEqual(expected, TranslationManager.ResolveLanguage(TranslationManager.Automatic))
        Assert.AreEqual(expected, TranslationManager.ResolveLanguage(""))
        Assert.AreEqual(TranslationManager.English, TranslationManager.ResolveLanguage("de"), "an unknown language is English")
    End Sub

    <TestMethod>
    Public Sub EnvironmentVariable_OverridesTheChoice()
        Environment.SetEnvironmentVariable("LITETASK_LANG", "en")
        Assert.AreEqual(TranslationManager.English, TranslationManager.ResolveLanguage(TranslationManager.French))
        Environment.SetEnvironmentVariable("LITETASK_LANG", "fr")
        Assert.AreEqual(TranslationManager.French, TranslationManager.ResolveLanguage(TranslationManager.English))
    End Sub

    <TestMethod>
    Public Sub FrenchFile_KeysAreUnique()
        Dim doc As New XmlDocument()
        doc.Load(Path.Combine(LangFolder(), "fr.xml"))
        Dim duplicates = doc.SelectNodes("/translations/translation").Cast(Of XmlNode)().
            GroupBy(Function(n) n.Attributes("key").Value, StringComparer.OrdinalIgnoreCase).
            Where(Function(g) g.Count() > 1).Select(Function(g) g.Key).ToList()
        Assert.AreEqual(0, duplicates.Count, "Keys defined twice in fr.xml: " & String.Join(", ", duplicates))
    End Sub

    <TestMethod>
    Public Sub FrenchFile_NonBreakingSpaceBeforeColonQuestionExclamationSemicolon()
        Dim wrong As New List(Of String)
        For Each pair In FrenchKeys()
            ' "C:\", "https:" and "{0:N0}" are not French punctuation
            Dim text = Regex.Replace(pair.Value, "https?:|[A-Za-z]:\\|\{\d+:[^}]*\}", "")
            If Regex.IsMatch(text, "( |[^\s\u00A0])[;?!]|( |\w)[:](\s|$)") Then wrong.Add(pair.Key & " = " & pair.Value)
        Next
        Assert.AreEqual(0, wrong.Count, "Use a non-breaking space before : ? ! ; in: " & String.Join(" | ", wrong))
    End Sub

    <TestMethod>
    Public Sub Code_EveryTranslationHasAnEnglishText()
        ' GetTranslation("key") alone shows the key itself in English: always give the English text
        Dim wrong As New List(Of String)
        For Each file In Sources()
            For Each m As Match In Regex.Matches(IO.File.ReadAllText(file), "\bGetTranslation\(\s*""[^""]*""\s*\)")
                wrong.Add(Path.GetFileName(file) & ": " & m.Value)
            Next
        Next
        Assert.AreEqual(0, wrong.Count, String.Join(", ", wrong))
    End Sub

    <TestMethod>
    Public Sub Code_EveryKeyIsTranslatedInFrench()
        Dim french = FrenchKeys()
        Dim missing As New List(Of String)
        For Each file In Sources()
            Dim text = IO.File.ReadAllText(file)
            For Each m As Match In Regex.Matches(text, "(?:\bGetTranslation|\bT|\bF|\bTr)\(\s*""([^""]+)"",\s*""")
                Dim key = m.Groups(1).Value
                If Not french.ContainsKey(key) AndAlso Not missing.Contains(key) Then missing.Add(key)
            Next
        Next
        Assert.AreEqual(0, missing.Count, "Keys of the code missing from fr.xml: " & String.Join(", ", missing))
    End Sub

    <TestMethod>
    Public Sub Format_KeepsThePlaceholdersOfTheEnglishText()
        ' A French format must use the same {n} as the English one, or String.Format fails at run time
        Dim french = FrenchKeys()
        Dim wrong As New List(Of String)
        For Each file In Sources()
            For Each m As Match In Regex.Matches(IO.File.ReadAllText(file), "\bF\(\s*""([^""]+)"",\s*""((?:[^""]|"""")*)""")
                Dim key = m.Groups(1).Value
                If Not french.ContainsKey(key) Then Continue For
                Dim english = Regex.Matches(m.Groups(2).Value, "\{\d+").Cast(Of Match)().Select(Function(x) x.Value).Distinct().OrderBy(Function(x) x)
                Dim translated = Regex.Matches(french(key), "\{\d+").Cast(Of Match)().Select(Function(x) x.Value).Distinct().OrderBy(Function(x) x)
                If Not english.SequenceEqual(translated) Then wrong.Add(key)
            Next
        Next
        Assert.AreEqual(0, wrong.Count, "Placeholders differ: " & String.Join(", ", wrong))
    End Sub

    <TestMethod>
    Public Sub Options_SavedChoiceIndex()
        Dim choices = {"auto", "en", "fr"}
        Assert.AreEqual(2, OptionsForm.ChoiceIndex(choices, "fr"))
        Assert.AreEqual(1, OptionsForm.ChoiceIndex(choices, "EN"))
        Assert.AreEqual(0, OptionsForm.ChoiceIndex(choices, "xx"), "an unknown value is Automatic")
        Assert.AreEqual(0, OptionsForm.ChoiceIndex(choices, Nothing))
    End Sub

    <TestMethod>
    Public Sub Theme_PreferenceOrWindows()
        Assert.IsTrue(Theme.ResolveDark("dark", Function() False))
        Assert.IsFalse(Theme.ResolveDark("light", Function() True))
        Assert.IsTrue(Theme.ResolveDark("auto", Function() True))
        Assert.IsFalse(Theme.ResolveDark("auto", Function() False))
        Assert.IsTrue(Theme.ResolveDark(Nothing, Function() True), "no preference follows Windows")
    End Sub

    <TestMethod>
    Public Sub Theme_DarkRow_KeepsATintOfTheStateColor()
        Dim surface = Color.FromArgb(43, 43, 43)
        Assert.AreEqual(surface, Theme.DarkRowBack(SystemColors.Window, surface), "default row")
        Assert.AreEqual(surface, Theme.DarkRowBack(Color.White, surface))
        Dim tinted = Theme.DarkRowBack(Color.LightGreen, surface)
        Assert.IsTrue(tinted.GetBrightness() < 0.45F, "the row stays dark: light text remains readable")
        Assert.IsTrue(tinted.G > tinted.R, "the state color is still recognizable")
    End Sub
End Class
