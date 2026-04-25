' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    ParameterParserTests.vb
' Author:  LiteTask contributors
' Date:    2026-04-25
' Purpose: Unit tests for LiteTask.ParameterParser covering the original
'          unquoted form and the new cmd-style quoting.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

Namespace LiteTask.Tests
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

    End Class
End Namespace
