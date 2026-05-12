' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ParameterParserTests.vb
' Author:  LiteTask contributors
' Date:    2026-04-25
' Purpose: Unit tests for LiteTask.ParameterParser covering the historical
'          `key=value` form, cmd-style quoting, and PowerShell-style
'          `-Name Value` / switch syntax.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class ParameterParserTests

    <TestMethod>
    Public Sub Parse_SinglePair_ReturnsKeyAndValue()
        Dim result = ParameterParser.Parse("key=value")
        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("value", result("key"))
    End Sub

    <TestMethod>
    Public Sub Parse_QuotedValueWithSpaces_KeepsValueIntact()
        Dim result = ParameterParser.Parse("key=""value with spaces""")
        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("value with spaces", result("key"))
    End Sub

    <TestMethod>
    Public Sub Parse_MultiplePairs_AllParsed()
        Dim result = ParameterParser.Parse("a=1 b=2 c=3")
        Assert.AreEqual(3, result.Count)
        Assert.AreEqual("1", result("a"))
        Assert.AreEqual("2", result("b"))
        Assert.AreEqual("3", result("c"))
    End Sub

    <TestMethod>
    Public Sub Parse_MixedQuotedAndUnquoted_BothHandled()
        Dim result = ParameterParser.Parse("a=1 b=""hello world"" c=3")
        Assert.AreEqual(3, result.Count)
        Assert.AreEqual("1", result("a"))
        Assert.AreEqual("hello world", result("b"))
        Assert.AreEqual("3", result("c"))
    End Sub

    <TestMethod>
    Public Sub Parse_EmptyQuotedValue_YieldsEmptyString()
        Dim result = ParameterParser.Parse("a=1 b="""" c=3")
        Assert.AreEqual(3, result.Count)
        Assert.AreEqual("", result("b"))
    End Sub

    <TestMethod>
    Public Sub Parse_EmptyUnquotedValue_YieldsEmptyString()
        Dim result = ParameterParser.Parse("a= b=2")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("", result("a"))
        Assert.AreEqual("2", result("b"))
    End Sub

    <TestMethod>
    Public Sub Parse_NullOrEmpty_ReturnsEmptyDictionary()
        Assert.AreEqual(0, ParameterParser.Parse(Nothing).Count)
        Assert.AreEqual(0, ParameterParser.Parse("").Count)
        Assert.AreEqual(0, ParameterParser.Parse("   ").Count)
    End Sub

    <TestMethod>
    Public Sub Parse_TokenWithoutEquals_IsSkipped()
        Dim result = ParameterParser.Parse("a=1 garbage b=2")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("1", result("a"))
        Assert.AreEqual("2", result("b"))
        Assert.IsFalse(result.ContainsKey("garbage"))
    End Sub

    <TestMethod>
    Public Sub Parse_QuotedValueFollowedByMore_ContinuesParsing()
        Dim result = ParameterParser.Parse("path=""C:\Program Files\App"" mode=fast")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("C:\Program Files\App", result("path"))
        Assert.AreEqual("fast", result("mode"))
    End Sub

    <TestMethod>
    Public Sub Parse_DuplicateKey_LastWriteWins()
        Dim result = ParameterParser.Parse("a=1 a=2 a=3")
        Assert.AreEqual(1, result.Count)
        Assert.AreEqual("3", result("a"))
    End Sub

    <TestMethod>
    Public Sub Parse_LeadingTrailingWhitespace_Tolerated()
        Dim result = ParameterParser.Parse("   a=1   b=2   ")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("1", result("a"))
        Assert.AreEqual("2", result("b"))
    End Sub

    <TestMethod>
    Public Sub Parse_PowerShellStyle_NameValue_Pairs()
        Dim result = ParameterParser.Parse("-WindowHours 96 -EmailTo ""user@example.com""")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("96", result("WindowHours"))
        Assert.AreEqual("user@example.com", result("EmailTo"))
    End Sub

    <TestMethod>
    Public Sub Parse_PowerShellStyle_SwitchBetweenParams_StoredAsNothing()
        Dim result = ParameterParser.Parse("-Force -WindowHours 96")
        Assert.AreEqual(2, result.Count)
        Assert.IsTrue(result.ContainsKey("Force"))
        Assert.IsNull(result("Force"))
        Assert.AreEqual("96", result("WindowHours"))
    End Sub

    <TestMethod>
    Public Sub Parse_PowerShellStyle_TrailingSwitch_StoredAsNothing()
        Dim result = ParameterParser.Parse("-WindowHours 96 -Verbose")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("96", result("WindowHours"))
        Assert.IsNull(result("Verbose"))
    End Sub

    <TestMethod>
    Public Sub Parse_PowerShellStyle_NameEqualsValue_Accepted()
        Dim result = ParameterParser.Parse("-Mode=fast -Count=3")
        Assert.AreEqual(2, result.Count)
        Assert.AreEqual("fast", result("Mode"))
        Assert.AreEqual("3", result("Count"))
    End Sub

    <TestMethod>
    Public Sub Parse_MixedStyles_BothParsed()
        Dim result = ParameterParser.Parse("alpha=one -Force -Beta ""two three""")
        Assert.AreEqual(3, result.Count)
        Assert.AreEqual("one", result("alpha"))
        Assert.IsNull(result("Force"))
        Assert.AreEqual("two three", result("Beta"))
    End Sub

    <TestMethod>
    Public Sub Parse_KeyLookup_IsCaseInsensitive()
        Dim result = ParameterParser.Parse("-WindowHours 96")
        Assert.AreEqual("96", result("windowhours"))
        Assert.AreEqual("96", result("WINDOWHOURS"))
    End Sub

End Class
