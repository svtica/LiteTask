' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    SqlReadinessClassificationTests.vb
' Author:  LiteTask contributors
' Date:    2026-11-10
' Purpose: Pins the SqlReadinessChecker failure classification to its 1.2.1.12
'          deny-list (GLPI #4302, "no behaviour change"). v1.2.1.13 only adds
'          logging around it; changing the list is a separate decision.
' -----------------------------------------------------------------------------
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports LiteTask.LiteTask

<TestClass>
Public Class SqlReadinessClassificationTests

    <TestMethod>
    Public Sub DenyList_IsExactly_1_2_1_12()
        For Each number In {-1, 2, 40, 53, 258, 1231, 10060, 10061, 11001}
            Assert.IsTrue(SqlReadinessChecker.IsNetworkErrorNumber(number), $"{number} must stay classified as server unreachable")
        Next
    End Sub

    <TestMethod>
    Public Sub NumbersOutsideDenyList_StillCountAsEngineAnswered_AsIn_1_2_1_12()
        ' Documented, not endorsed: 64 is the number logged on 2026-10-27 12:30:42.
        ' Reclassifying it is out of scope for 1.2.1.13.
        For Each number In {64, -2, 121, 18456, 4060}
            Assert.IsFalse(SqlReadinessChecker.IsNetworkErrorNumber(number), $"{number} is not in the 1.2.1.12 deny-list")
        Next
    End Sub

End Class
