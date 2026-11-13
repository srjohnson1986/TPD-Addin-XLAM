Attribute VB_Name = "modFlowTest_EQList"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Flow tests for the Customer EQ List: the whole of
'  modMain_CustEQList.CreateCustEQList_DoWork, run against
'  tests/fixtures/EQ_List_Sample.xlsx.
'
'  What the fixture holds below its heading row (rows 2-88):
'  16 PARENT rows, 69 item rows, 1 INCLUDED row and 1 blank
'  row (source row 85) - 87 rows in all. Source row 8 (T102)
'  is bold and italic throughout; the SIZE cell of source
'  row 31 (S106) is bold and italic; source row 57 is a
'  yellow "For Approval" item with no INTERNAL ID. Change the
'  fixture generator and these numbers change with it.
'
'  Telling PARENT rows from item rows: 7 of the 16 PARENT
'  rows have an INTERNAL ID, so "no ID" does not mean PARENT.
'  Those 7 are a lighter blue and not bold in the source;
'  only the 9 without an ID are bold.
'  Tests that care which rows are PARENT rows ask for the
'  Purchased column as well as the shipped defaults (it comes
'  out last) and read it. Rows are never found by row number.
'
'  Not checked here: column widths, row heights, where the
'  logo sits, frozen panes - nothing about how it looks.
'
'  Not part of /src. Headless only: tools\Test-Unit.ps1 -Suite flows
'===========================================================

Option Explicit
Option Private Module

Private Const FIXTURE As String = "EQ_List_Sample.xlsx"
Private Const PARENT_ROWS As Long = 16
Private Const BOLD_PARENT_ROWS As Long = 9    ' the 9 with no INTERNAL ID; the other 7 are lighter blue and not bold
Private Const ITEM_ROWS As Long = 70          ' 69 items + the INCLUDED row
Private Const NUMBERED_ROWS As Long = 69      ' EQ Count skips PARENT and INCLUDED

Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub

' The shipped default columns, then Purchased.
Private Function ColumnsWithPurchased() As String
    ColumnsWithPurchased = DefaultEqListColumns() & ", Purchased"
End Function

' Runs the flow on the fixture's sheet and returns the sheet it made.
Private Function RunEqList(ByVal wb As Workbook, ByVal columnsCsv As String, ByVal toggles As Object) As Worksheet
    Dim sheetsBefore As Long

    sheetsBefore = wb.Worksheets.Count
    modMain_CustEQList.CreateCustEQList_DoWork wb.Worksheets(1), ColumnsFromCsv(columnsCsv), toggles

    Assert.AreEqual sheetsBefore + 1, wb.Worksheets.Count, "the flow should add exactly one sheet"
    Set RunEqList = wb.Worksheets(wb.Worksheets.Count)
End Function

Private Function HeadingRow() As Long
    HeadingRow = EQ_HEADER_ROW_COUNT + 1
End Function

' Counts the data rows by kind, reading the Purchased column: how many are
' PARENT rows and how many of those have a coloured fill / are bold, and the
' same for every other row. The first cell of the row is the one looked at.
Private Sub TallyByKind(ByVal ws As Worksheet, _
                        ByRef parents As Long, ByRef parentsColoured As Long, ByRef parentsBold As Long, _
                        ByRef others As Long, ByRef othersColoured As Long, ByRef othersBold As Long)
    Dim r As Long
    Dim purchasedCol As Long
    Dim isParent As Boolean

    parents = 0: parentsColoured = 0: parentsBold = 0
    others = 0: othersColoured = 0: othersBold = 0

    purchasedCol = ColumnOf(ws, HeadingRow(), "Purchased")
    If purchasedCol = 0 Then
        Assert.Fail "this test needs the Purchased column in the output, and it is not there. " & _
                    DescribeSheet(ws, HeadingRow())
        Exit Sub
    End If

    For r = HeadingRow() + 1 To GetLastRow(ws)
        isParent = (UCase$(Trim$(CStr(ws.Cells(r, purchasedCol).Value))) = "PARENT")
        If isParent Then
            parents = parents + 1
            If HasColouredFill(ws.Cells(r, 1)) Then parentsColoured = parentsColoured + 1
            If ws.Cells(r, 1).Font.Bold = True Then parentsBold = parentsBold + 1
        Else
            others = others + 1
            If HasColouredFill(ws.Cells(r, 1)) Then othersColoured = othersColoured + 1
            If ws.Cells(r, 1).Font.Bold = True Then othersBold = othersBold + 1
        End If
    Next r
End Sub

' ---- as shipped ------------------------------------------------------------

'@TestMethod("EQ List")
Public Sub EqList_AsShipped_SheetColumnsAndRowCount()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith())

    Assert.AreEqual "Customer EQ List", ws.Name
    Assert.AreEqual CsvAsHeadings(DefaultEqListColumns()), HeadingsJoined(ws, HeadingRow()), _
                    "only the chosen columns, in the chosen order, on the heading row under the title block"
    Assert.AreEqual PARENT_ROWS + ITEM_ROWS, DataRowCount(ws, HeadingRow()), "PARENT rows + item rows, blank row gone"
    Assert.AreEqual 0, BlankDataRowCount(ws, HeadingRow()), "Remove blank rows ships on (#171)"
    Assert.AreEqual 1, ws.Shapes.Count, "one logo and nothing else"
    Assert.IsFalse ws.AutoFilterMode, "Add column filters ships off"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List")
Public Sub EqList_AsShipped_ItemRowsLoseCarriedFormatting()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim idCol As Long, sizeCol As Long
    Dim r As Long
    Dim parents As Long, parentsColoured As Long, parentsBold As Long
    Dim others As Long, othersColoured As Long, othersBold As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, ColumnsWithPurchased(), TogglesWith())
    idCol = ColumnOf(ws, HeadingRow(), "INTERNAL ID")
    sizeCol = ColumnOf(ws, HeadingRow(), "SIZE")

    ' Source row 8: bold and italic across the whole row (#169).
    r = FindDataRow(ws, HeadingRow(), idCol, "T102")
    Assert.IsTrue r > 0, "row T102 should be in the output"
    Assert.IsFalse ws.Cells(r, idCol).Font.Bold = True, "T102 should not be bold"
    Assert.IsFalse ws.Cells(r, idCol).Font.Italic = True, "T102 should not be italic"

    ' Source row 31: only the SIZE cell was bold and italic.
    r = FindDataRow(ws, HeadingRow(), idCol, "S106")
    Assert.IsTrue r > 0, "row S106 should be in the output"
    Assert.IsFalse ws.Cells(r, sizeCol).Font.Bold = True, "S106's SIZE cell should not be bold"
    Assert.IsFalse ws.Cells(r, sizeCol).Font.Italic = True, "S106's SIZE cell should not be italic"

    TallyByKind ws, parents, parentsColoured, parentsBold, others, othersColoured, othersBold
    Assert.AreEqual ITEM_ROWS, others, "item rows"
    Assert.AreEqual 0, othersColoured, "no item row keeps its source fill - including the ID-less yellow one (#170)"
    Assert.AreEqual 0, othersBold, "no item row is bold"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List")
Public Sub EqList_AsShipped_ParentRowsKeepTheirLook()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parents As Long, parentsColoured As Long, parentsBold As Long
    Dim others As Long, othersColoured As Long, othersBold As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, ColumnsWithPurchased(), TogglesWith())
    TallyByKind ws, parents, parentsColoured, parentsBold, others, othersColoured, othersBold

    Assert.AreEqual PARENT_ROWS, parents, "PARENT rows"
    Assert.AreEqual PARENT_ROWS, parentsColoured, "every PARENT row keeps its fill"
    Assert.AreEqual BOLD_PARENT_ROWS, parentsBold, "the PARENT rows that are bold in the source stay bold, and no others become bold"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List")
Public Sub EqList_AsShipped_DataCellsHaveBorders()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith())

    Assert.AreEqual CLng(xlContinuous), CLng(ws.Cells(HeadingRow() + 1, 1).Borders(xlEdgeBottom).LineStyle), _
                    "Add cell borders ships on"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' ---- one option changed at a time ------------------------------------------

'@TestMethod("EQ List options")
Public Sub EqList_RemoveBlankRowsOff_KeepsTheBlankRow()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith(PREF_EQLIST_REMOVE_BLANK_ROWS, False))

    Assert.AreEqual PARENT_ROWS + ITEM_ROWS + 1, DataRowCount(ws, HeadingRow())
    Assert.AreEqual 1, BlankDataRowCount(ws, HeadingRow())

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_RemoveParentRows_LeavesOnlyItems()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    ' The shipped columns, which do not include Purchased.
    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith(PREF_EQLIST_REMOVE_PARENT_ROWS, True))

    Assert.AreEqual ITEM_ROWS, DataRowCount(ws, HeadingRow()), _
                    "the 16 PARENT rows go and the 70 item rows stay. " & DescribeSheet(ws, HeadingRow())

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_RemoveParentRows_PurchasedKept_NoParentRowLeft()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parents As Long, parentsColoured As Long, parentsBold As Long
    Dim others As Long, othersColoured As Long, othersBold As Long
    On Error GoTo TestFail

    ' The same, with Purchased among the chosen columns so the rows can be told apart.
    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, ColumnsWithPurchased(), TogglesWith(PREF_EQLIST_REMOVE_PARENT_ROWS, True))

    Assert.AreEqual ITEM_ROWS, DataRowCount(ws, HeadingRow()), _
                    "the 16 PARENT rows go and the 70 item rows stay. " & DescribeSheet(ws, HeadingRow())

    If DataRowCount(ws, HeadingRow()) > 0 Then
        TallyByKind ws, parents, parentsColoured, parentsBold, others, othersColoured, othersBold
        Assert.AreEqual 0, parents, "no PARENT row left"
        Assert.AreEqual ITEM_ROWS, others, "every item row is still there"
    End If

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_AddEqCount_NumbersItemsOnly()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim numbered As Long
    Dim highest As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith(PREF_EQLIST_ADD_COUNT_COLUMN, True))

    Assert.AreEqual "EQ COUNT|" & CsvAsHeadings(DefaultEqListColumns()), HeadingsJoined(ws, HeadingRow()), _
                    "EQ COUNT goes in as the leftmost column"

    For r = HeadingRow() + 1 To GetLastRow(ws)
        If Len(CStr(ws.Cells(r, 1).Value)) > 0 Then
            numbered = numbered + 1
            If CLng(ws.Cells(r, 1).Value) > highest Then highest = CLng(ws.Cells(r, 1).Value)
        End If
    Next r

    Assert.AreEqual NUMBERED_ROWS, numbered, "every item is numbered; PARENT and INCLUDED rows are not"
    Assert.AreEqual NUMBERED_ROWS, highest, "numbering runs 1 to the item count with no gaps"
    Assert.AreEqual "00", ws.Cells(HeadingRow() + 1, 1).NumberFormat, "leading zeros sized to the highest number"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_PlainParentRows_NoFillAndNotBold()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parents As Long, parentsColoured As Long, parentsBold As Long
    Dim others As Long, othersColoured As Long, othersBold As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, ColumnsWithPurchased(), TogglesWith(PREF_EQLIST_PLAIN_PARENT_ROWS, True))
    TallyByKind ws, parents, parentsColoured, parentsBold, others, othersColoured, othersBold

    Assert.AreEqual PARENT_ROWS, parents, "the PARENT rows are still there"
    Assert.AreEqual 0, parentsColoured, "PARENT rows are white (#121)"
    Assert.AreEqual 0, parentsBold, "and no longer bold"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_PlainAndBoldParentRows_WhiteRowsWithBoldText()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parents As Long, parentsColoured As Long, parentsBold As Long
    Dim others As Long, othersColoured As Long, othersBold As Long
    On Error GoTo TestFail

    ' Bold PARENT rows on its own cannot be told from the source, whose PARENT
    ' rows are already bold. With Plain on as well, bold has to be put back.
    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, ColumnsWithPurchased(), _
                       TogglesWith(PREF_EQLIST_PLAIN_PARENT_ROWS, True, PREF_EQLIST_BOLD_PARENT_ROWS, True))
    TallyByKind ws, parents, parentsColoured, parentsBold, others, othersColoured, othersBold

    Assert.AreEqual 0, parentsColoured, "PARENT rows are white"
    Assert.AreEqual PARENT_ROWS, parentsBold, "and bold again (#174)"
    Assert.AreEqual 0, othersBold, "item rows are not made bold"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_AddCellBordersOff_NoBorders()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith(PREF_EQLIST_ADD_CELL_BORDERS, False))

    Assert.AreEqual CLng(xlLineStyleNone), CLng(ws.Cells(HeadingRow() + 1, 1).Borders(xlEdgeBottom).LineStyle), _
                    "the fixture has no borders of its own, so none should appear (#122)"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("EQ List options")
Public Sub EqList_AddColumnFilters_FilterSitsOnTheHeadingRow()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunEqList(wb, DefaultEqListColumns(), TogglesWith(PREF_EQLIST_ADD_COLUMN_FILTERS, True))

    Assert.IsTrue ws.AutoFilterMode, "a filter should be on"
    If ws.AutoFilterMode Then
        Assert.AreEqual HeadingRow(), ws.AutoFilter.Range.Row, _
                        "the title block pushes the filter down to the heading row (#147)"
    End If

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
