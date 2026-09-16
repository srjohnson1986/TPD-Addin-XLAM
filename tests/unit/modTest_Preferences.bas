Attribute VB_Name = "modTest_Preferences"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Rubberduck unit tests for modPreferences's pure functions
'  (#163): FirstNonEmptyColumnList, FirstNonEmptyValue. Not
'  part of /src - never imported by BuildAddin, never shipped
'  in the .xlam. Import this file plus modPreferences into a
'  scratch workbook that has a reference to Rubberduck, then
'  run from the VBE (Rubberduck > Run All Tests). See
'  CONTRIBUTING.md for the rest of the dev workflow.
'===========================================================

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
End Sub

Private Function JoinColl(ByVal c As Collection) As String
    Dim v As Variant
    Dim s As String
    For Each v In c
        s = s & CStr(v) & "|"
    Next v
    JoinColl = s
End Function

'@TestMethod("FirstNonEmptyColumnList")
Public Sub FirstNonEmptyColumnList_PicksFirstNonEmptyCandidate()
    On Error GoTo TestFail

    Dim result As Collection
    ' Picker precedence: LastUsed (unset) -> DefaultUser (set) -> shipped.
    Set result = modPreferences.FirstNonEmptyColumnList(Array("", "Vendor, Model"), "Vendor, Model, Qty")

    Assert.AreEqual "Vendor|Model|", JoinColl(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FirstNonEmptyColumnList")
Public Sub FirstNonEmptyColumnList_AllCandidatesEmpty_FallsBackToShipped()
    On Error GoTo TestFail

    Dim result As Collection
    Set result = modPreferences.FirstNonEmptyColumnList(Array("", ""), "Vendor, Model, Qty")

    Assert.AreEqual "Vendor|Model|Qty|", JoinColl(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FirstNonEmptyColumnList")
Public Sub FirstNonEmptyColumnList_FirstCandidateWinsOverLater()
    On Error GoTo TestFail

    Dim result As Collection
    Set result = modPreferences.FirstNonEmptyColumnList(Array("Serial", "Vendor, Model"), "Qty")

    Assert.AreEqual "Serial|", JoinColl(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FirstNonEmptyValue")
Public Sub FirstNonEmptyValue_PicksFirstNonEmptyCandidate()
    On Error GoTo TestFail

    Dim result As String
    result = modPreferences.FirstNonEmptyValue(Array("", "  ", "Vendor"), "Model")

    Assert.AreEqual "Vendor", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FirstNonEmptyValue")
Public Sub FirstNonEmptyValue_AllCandidatesEmpty_FallsBack()
    On Error GoTo TestFail

    Dim result As String
    result = modPreferences.FirstNonEmptyValue(Array("", "   "), "Vendor")

    Assert.AreEqual "Vendor", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
