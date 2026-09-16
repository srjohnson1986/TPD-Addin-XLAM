Attribute VB_Name = "modTest_ColumnReorderAndEQCount"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Rubberduck unit tests for the last #163 extractions:
'  modHelpers_Columns.PlanColumnMoves and
'  modMain_CountEquipmentRows.ComputeEQCountValues /
'  EQCountNumberFormat. Not part of /src - never imported by
'  BuildAddin, never shipped in the .xlam. Import this file
'  plus modHelpers_Columns, modHelpers_Headers and
'  modMain_CountEquipmentRows into a scratch workbook that has
'  a reference to Rubberduck, then run from the VBE
'  (Rubberduck > Run All Tests). See CONTRIBUTING.md for the
'  rest of the dev workflow.
'===========================================================

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
End Sub

Private Function Coll(ParamArray items() As Variant) As Collection
    Dim result As New Collection
    Dim i As Long
    For i = LBound(items) To UBound(items)
        result.Add items(i)
    Next i
    Set Coll = result
End Function

' Applies `moves` (as PlanColumnMoves returns them) to a plain array the same
' way ws.Columns(from).Cut / ws.Columns(to).Insert would to a live sheet, so
' tests can assert on the resulting order without touching Excel.
Private Function ApplyMoves(ByVal headings As Variant, ByVal moves As Collection) As String
    Dim arr() As String
    Dim n As Long, i As Long, j As Long
    Dim mv As Variant
    Dim moved As String

    n = UBound(headings) - LBound(headings) + 1
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(headings(LBound(headings) + i - 1))
    Next i

    For Each mv In moves
        moved = arr(CLng(mv(0)))
        For j = CLng(mv(0)) To CLng(mv(1)) + 1 Step -1
            arr(j) = arr(j - 1)
        Next j
        arr(CLng(mv(1))) = moved
    Next mv

    Dim s As String
    For i = 1 To n
        s = s & arr(i) & "|"
    Next i
    ApplyMoves = s
End Function

'@TestMethod("PlanColumnMoves")
Public Sub PlanColumnMoves_ReordersToMatchOrderedHeadings()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim moves As Collection

    headings = Array("Vendor", "Model", "Serial", "Qty")
    Set moves = modHelpers_Columns.PlanColumnMoves(headings, Coll("Qty", "Vendor", "Model", "Serial"))

    Assert.AreEqual "Qty|Vendor|Model|Serial|", ApplyMoves(headings, moves)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PlanColumnMoves")
Public Sub PlanColumnMoves_AlreadyInOrder_NoMoves()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim moves As Collection

    headings = Array("Vendor", "Model", "Qty")
    Set moves = modHelpers_Columns.PlanColumnMoves(headings, Coll("Vendor", "Model", "Qty"))

    Assert.AreEqual 0, moves.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PlanColumnMoves")
Public Sub PlanColumnMoves_HeadingNotOnSheet_Skipped()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim moves As Collection

    headings = Array("Vendor", "Model")
    Set moves = modHelpers_Columns.PlanColumnMoves(headings, Coll("Serial", "Model", "Vendor"))

    ' "Serial" isn't there - skipped, doesn't consume a target slot - so Model
    ' (already 2nd) doesn't move and only Vendor needs to move to the end.
    Assert.AreEqual "Model|Vendor|", ApplyMoves(headings, moves)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PlanColumnMoves")
Public Sub PlanColumnMoves_UnselectedTrailingColumnsUntouched()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim moves As Collection

    headings = Array("Vendor", "Model", "Notes")
    Set moves = modHelpers_Columns.PlanColumnMoves(headings, Coll("Model", "Vendor"))

    Assert.AreEqual "Model|Vendor|Notes|", ApplyMoves(headings, moves)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ComputeEQCountValues")
Public Sub ComputeEQCountValues_SkipsParentAndIncluded_NumbersRest()
    On Error GoTo TestFail

    Dim statuses As Variant
    Dim result As Variant

    statuses = Array("", "PARENT", "INCLUDED", "")
    result = modMain_CountEquipmentRows.ComputeEQCountValues(statuses)

    Assert.AreEqual 1, result(LBound(result))
    Assert.AreEqual "", result(LBound(result) + 1)
    Assert.AreEqual "", result(LBound(result) + 2)
    Assert.AreEqual 2, result(LBound(result) + 3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ComputeEQCountValues")
Public Sub ComputeEQCountValues_EmptyStatuses_ReturnsEmptyArray()
    On Error GoTo TestFail

    Dim result As Variant
    result = modMain_CountEquipmentRows.ComputeEQCountValues(Array())

    Assert.IsTrue IsArray(result)
    Assert.IsTrue UBound(result) < LBound(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQCountNumberFormat")
Public Sub EQCountNumberFormat_WidensToLargestNumber()
    On Error GoTo TestFail

    Dim result As String
    result = modMain_CountEquipmentRows.EQCountNumberFormat(Array(1, 2, "", 3, 4, 5, 6, 7, 8, 9, 10))

    Assert.AreEqual "00", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQCountNumberFormat")
Public Sub EQCountNumberFormat_AllBlank_ReturnsSingleZero()
    On Error GoTo TestFail

    Dim result As String
    result = modMain_CountEquipmentRows.EQCountNumberFormat(Array("", ""))

    Assert.AreEqual "0", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
