Attribute VB_Name = "modTest_HelpersColumns"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Rubberduck unit tests for modHelpers_Columns's pure
'  functions (#163): HeadingsPresentIn, UniqueValues,
'  BuildOrderedColumnMap. Not part of /src - never imported
'  by BuildAddin, never shipped in the .xlam. Import this
'  file plus modHelpers_Columns and modHelpers_Headers into a
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

Private Function Coll(ParamArray items() As Variant) As Collection
    Dim result As New Collection
    Dim i As Long
    For i = LBound(items) To UBound(items)
        result.Add items(i)
    Next i
    Set Coll = result
End Function

' BuildOrderedColumnMap's returned keys are the same index base as the
' `headings` array passed in - real callers always pass
' modHelpers_Headers.GetHeadingList's output, which is 1-based. VBA's Array()
' function is always 0-based (even under Option Base 1), so tests that assert
' on the actual key VALUES (column numbers) need a genuinely 1-based fixture.
Private Function OneBasedHeadings(ParamArray items() As Variant) As Variant
    Dim result() As Variant
    Dim i As Long
    ReDim result(1 To UBound(items) - LBound(items) + 1)
    For i = LBound(items) To UBound(items)
        result(i - LBound(items) + 1) = items(i)
    Next i
    OneBasedHeadings = result
End Function

Private Function JoinColl(ByVal c As Collection) As String
    Dim v As Variant
    Dim s As String
    For Each v In c
        s = s & CStr(v) & "|"
    Next v
    JoinColl = s
End Function

'@TestMethod("HeadingsPresentIn")
Public Sub HeadingsPresentIn_KeepsWantedOrderAndSpelling()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim result As Collection

    headings = Array("Vendor", "Model", "Serial", "Qty")
    Set result = modHelpers_Columns.HeadingsPresentIn(headings, Coll("QTY", "Missing", "vendor"))

    Assert.AreEqual "QTY|vendor|", JoinColl(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("HeadingsPresentIn")
Public Sub HeadingsPresentIn_NoneMatch_ReturnsEmpty()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim result As Collection

    headings = Array("Vendor", "Model")
    Set result = modHelpers_Columns.HeadingsPresentIn(headings, Coll("Serial", "Qty"))

    Assert.AreEqual 0, result.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueValues")
Public Sub UniqueValues_DedupesCaseInsensitiveAndSkipsBlanks()
    On Error GoTo TestFail

    Dim values As Variant
    Dim result As Collection

    values = Array("Acme", "acme", "", " ", "Beta", Empty, "Acme ")
    Set result = modHelpers_Columns.UniqueValues(values)

    ' NormalizeCellText trims but does not case-fold, so "Acme" and "acme" are
    ' distinct normalized strings - CollectionContainsText's case-insensitive
    ' compare is what actually dedupes them (matches GetUniqueValuesInColumn's
    ' existing behaviour).
    Assert.AreEqual "Acme|Beta|", JoinColl(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueValues")
Public Sub UniqueValues_AllBlank_ReturnsEmpty()
    On Error GoTo TestFail

    Dim values As Variant
    Dim result As Collection

    values = Array("", " ", Empty)
    Set result = modHelpers_Columns.UniqueValues(values)

    Assert.AreEqual 0, result.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BuildOrderedColumnMap")
Public Sub BuildOrderedColumnMap_OrdersBySelectedHeadings_NotSourceOrder()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim colMap As Object
    Dim keys As Variant

    ' Source order is Vendor, Model, Qty - selection asks for Qty before Vendor.
    headings = OneBasedHeadings("Vendor", "Model", "Qty")
    Set colMap = modHelpers_Columns.BuildOrderedColumnMap(headings, Coll("Qty", "Vendor"))

    keys = colMap.Keys
    Assert.AreEqual 2, colMap.Count
    Assert.AreEqual 3, keys(0)  ' Qty is column 3
    Assert.AreEqual 1, keys(1)  ' Vendor is column 1

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BuildOrderedColumnMap")
Public Sub BuildOrderedColumnMap_SkipsHeadingNotOnSheet()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim colMap As Object

    headings = OneBasedHeadings("Vendor", "Model")
    Set colMap = modHelpers_Columns.BuildOrderedColumnMap(headings, Coll("Vendor", "Serial"))

    Assert.AreEqual 1, colMap.Count
    Assert.IsTrue colMap.exists(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BuildOrderedColumnMap")
Public Sub BuildOrderedColumnMap_DuplicateHeading_PrefersRightmostAndDedupes()
    On Error GoTo TestFail

    Dim headings As Variant
    Dim colMap As Object
    Dim keys As Variant

    ' "Vendor" appears twice - FindHeadingIndex's default (preferRightmost:=True)
    ' means the rightmost (col 3) wins, matching the shared column-map logic
    ' used by the EQ List / Schedule / Split flows.
    headings = OneBasedHeadings("Vendor", "Model", "Vendor")
    Set colMap = modHelpers_Columns.BuildOrderedColumnMap(headings, Coll("Vendor"))

    keys = colMap.Keys
    Assert.AreEqual 1, colMap.Count
    Assert.AreEqual 3, keys(0)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
