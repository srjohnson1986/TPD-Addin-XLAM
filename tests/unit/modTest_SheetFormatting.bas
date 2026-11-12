Attribute VB_Name = "modTest_SheetFormatting"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Rubberduck unit tests for modHelpers_SheetFormatting's
'  pure IsStructuralRow - which EQ List rows keep their source
'  fill and font when the carried-over formatting is stripped
'  (#130 / #170). Not part of /src - never imported by
'  BuildAddin, never shipped in the .xlam. Import this file
'  plus modHelpers_SheetFormatting (and the modules it calls)
'  into a scratch workbook with a reference to Rubberduck, then
'  run from the VBE (Rubberduck > Run All Tests).
'===========================================================

Option Explicit
Option Private Module

Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_ParentRow_IsStructural()
    On Error GoTo TestFail

    Assert.AreEqual True, modHelpers_SheetFormatting.IsStructuralRow(True, "PARENT", True, "", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_ParentRow_CaseAndSpacesIgnored()
    On Error GoTo TestFail

    Assert.AreEqual True, modHelpers_SheetFormatting.IsStructuralRow(True, "  parent ", True, "TAG-1", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_ItemRowWithIds_IsNotStructural()
    On Error GoTo TestFail

    Assert.AreEqual False, modHelpers_SheetFormatting.IsStructuralRow(True, "NO", True, "TAG-1", "C-1")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_ItemRowWithNoIds_IsNotStructural_170()
    On Error GoTo TestFail

    Assert.AreEqual False, modHelpers_SheetFormatting.IsStructuralRow(True, "NO", True, "", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_BlankPurchasedNoIds_PurchasedColumnStillDecides()
    On Error GoTo TestFail

    Assert.AreEqual False, modHelpers_SheetFormatting.IsStructuralRow(True, "", True, "", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_NoPurchasedColumn_BothIdsBlank_IsStructural()
    On Error GoTo TestFail

    Assert.AreEqual True, modHelpers_SheetFormatting.IsStructuralRow(False, "", True, "  ", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_NoPurchasedColumn_OneIdSet_IsNotStructural()
    On Error GoTo TestFail

    Assert.AreEqual False, modHelpers_SheetFormatting.IsStructuralRow(False, "", True, "", "C-1")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsStructuralRow")
Public Sub IsStructuralRow_NoPurchasedOrIdColumns_NothingIsStructural()
    On Error GoTo TestFail

    Assert.AreEqual False, modHelpers_SheetFormatting.IsStructuralRow(False, "", False, "", "")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
