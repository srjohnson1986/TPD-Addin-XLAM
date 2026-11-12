Attribute VB_Name = "modTest_Preferences"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Rubberduck unit tests for modPreferences's pure functions
'  (#163): FirstNonEmptyColumnList, FirstNonEmptyValue - and
'  the EQ List toggle table in modPreferences_Defaults (#171).
'  Not part of /src - never imported by BuildAddin, never
'  shipped in the .xlam. Import this file plus modPreferences,
'  modPreferences_Defaults and modPreferences_KeyMap into a
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

'@TestMethod("EqListToggleTable")
Public Sub EqListToggle_RemoveBlankRows_ShipsOnAndNeedsNoPurchased()
    On Error GoTo TestFail

    Assert.IsTrue modPreferences_Defaults.EqListToggleShipsOn(PREF_EQLIST_REMOVE_BLANK_ROWS)
    Assert.IsFalse modPreferences_Defaults.EqListToggleNeedsPurchased(PREF_EQLIST_REMOVE_BLANK_ROWS)
    Assert.AreEqual "Remove blank rows", modPreferences_Defaults.EqListToggleName(PREF_EQLIST_REMOVE_BLANK_ROWS)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EqListToggleTable")
Public Sub EqListToggleKeys_RemoveBlankRowsListedLast()
    On Error GoTo TestFail

    Dim keys As Variant
    keys = modPreferences_Defaults.EqListToggleKeys()

    Assert.AreEqual PREF_EQLIST_REMOVE_BLANK_ROWS, keys(UBound(keys))
    Assert.AreEqual 6&, CLng(UBound(keys) - LBound(keys) + 1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
