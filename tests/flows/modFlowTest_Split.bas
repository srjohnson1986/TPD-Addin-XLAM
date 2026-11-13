Attribute VB_Name = "modFlowTest_Split"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Flow tests for Split Sheet by Column: the whole of
'  modMain_SplitSheets.SplitSheetByColumn_DoWork, splitting
'  tests/fixtures/EQ_List_Sample.xlsx on its Vendor column
'  with the shipped default columns.
'
'  The fixture has 67 rows with a Vendor, over 11 distinct
'  vendors once spaces and case are ignored. This is the
'  table in tests/README.md ("Vendor values for Split Sheet
'  by Column"), checked end to end: the unit tests cover the
'  sheet names on their own, these cover the split itself.
'
'  Not checked here: the "N sheet(s) created" summary box
'  (it is raised by the caller, not by _DoWork), column
'  widths, row heights, frozen panes.
'
'  Not part of /src. Headless only: tools\Test-Unit.ps1 -Suite flows
'===========================================================

Option Explicit
Option Private Module

Private Const FIXTURE As String = "EQ_List_Sample.xlsx"
Private Const VENDOR_SHEETS As Long = 11
Private Const ROWS_WITH_A_VENDOR As Long = 67

Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub

' Splits the fixture on Vendor. created and failures come back from _DoWork.
Private Sub RunSplitOnVendor(ByVal wb As Workbook, ByRef created As Long, ByRef failures As Collection)
    Set failures = New Collection
    created = modMain_SplitSheets.SplitSheetByColumn_DoWork( _
                  wb.Worksheets(1), "Vendor", ColumnsFromCsv(DefaultSplitColumns()), failures)
End Sub

' Asserts that the vendor's sheet exists and holds this many data rows.
Private Sub ExpectVendorSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal dataRows As Long)
    Dim ws As Worksheet

    Set ws = SheetNamed(wb, sheetName)
    If ws Is Nothing Then
        Assert.Fail "no sheet named '" & sheetName & "'"
        Exit Sub
    End If
    Assert.AreEqual dataRows, DataRowCount(ws, 1), "data rows on '" & sheetName & "'"
End Sub

'@TestMethod("Split")
Public Sub Split_OnVendor_MakesOneSheetPerVendor()
    Dim wb As Workbook
    Dim created As Long
    Dim failures As Collection
    Dim sheetsBefore As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    sheetsBefore = wb.Worksheets.Count
    RunSplitOnVendor wb, created, failures

    Assert.AreEqual VENDOR_SHEETS, created, "sheets reported as created"
    Assert.AreEqual 0, failures.Count, "no vendor should fail"
    Assert.AreEqual sheetsBefore + VENDOR_SHEETS, wb.Worksheets.Count, "sheets actually added"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Split")
Public Sub Split_OnVendor_OrdinaryVendorsGetTheirRows()
    Dim wb As Workbook
    Dim created As Long
    Dim failures As Collection
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    RunSplitOnVendor wb, created, failures

    ExpectVendorSheet wb, "Birchwood Controls", 6
    ExpectVendorSheet wb, "Juniper Bellows", 6
    ExpectVendorSheet wb, "Cedar Flow Systems", 5
    ExpectVendorSheet wb, "Hollis Thermal", 10
    ExpectVendorSheet wb, "Northgate Fittings", 10
    ExpectVendorSheet wb, "Dunmore Pumps", 10
    ExpectVendorSheet wb, "Fernhill Instruments", 7

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Split")
Public Sub Split_OnVendor_AwkwardVendorNames()
    Dim wb As Workbook
    Dim created As Long
    Dim failures As Collection
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    RunSplitOnVendor wb, created, failures

    ' "Alder Valve Co" and " alder valve co " are one vendor: 9 rows between them.
    ExpectVendorSheet wb, "Alder Valve Co", 9
    ' Characters a sheet name cannot hold become "_", then the name is cut at 31.
    ExpectVendorSheet wb, "Acme Valve_Fittings _West__ Ltd", 2
    ' Two vendors whose names are the same for the first 31 characters (#57).
    ExpectVendorSheet wb, "International Combustion & Cont", 1
    ExpectVendorSheet wb, "International Combustion &  (2)", 1

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Split")
Public Sub Split_OnVendor_EveryVendorRowLandsOnce_BlankVendorsNowhere()
    Dim wb As Workbook
    Dim created As Long
    Dim failures As Collection
    Dim i As Long
    Dim total As Long
    Dim sourceName As String
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    sourceName = wb.Worksheets(1).Name
    RunSplitOnVendor wb, created, failures

    ' Every sheet but the source is a vendor sheet.
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name <> sourceName Then total = total + DataRowCount(wb.Worksheets(i), 1)
    Next i

    Assert.AreEqual ROWS_WITH_A_VENDOR, total, _
                    "rows across all vendor sheets = rows that have a vendor; rows with none are on no sheet"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Split")
Public Sub Split_OnVendor_SheetsCarryTheChosenColumns_HeadingBold()
    Dim wb As Workbook
    Dim created As Long
    Dim failures As Collection
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    RunSplitOnVendor wb, created, failures
    Set ws = SheetNamed(wb, "Hollis Thermal")

    If ws Is Nothing Then
        Assert.Fail "no sheet named 'Hollis Thermal'"
    Else
        Assert.AreEqual CsvAsHeadings(DefaultSplitColumns()), HeadingsJoined(ws, 1), _
                        "only the chosen columns, in the chosen order"
        Assert.IsTrue ws.Cells(1, 1).Font.Bold = True, "the heading row is bold"
        Assert.IsFalse HasColouredFill(ws.Cells(2, 1)), "split copies values only, so no source fill comes across"
    End If

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
