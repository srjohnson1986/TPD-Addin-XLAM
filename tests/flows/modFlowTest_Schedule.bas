Attribute VB_Name = "modFlowTest_Schedule"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Flow tests for the Customer Schedule: the whole of
'  modMain_CustSchedule.CreateCustSchedule_DoWork, run
'  against tests/fixtures/Schedule_Sample.xlsx.
'
'  The fixture has 135 rows below its heading row: 27 bold
'  summary rows and 108 task rows, no blank row among them.
'  The flow deletes no rows, so source row N is always
'  output row N + SCHEDULE_HEADER_ROW_COUNT, and a source
'  cell can be compared with its output cell directly.
'
'  Source rows used below: 2 "Boiler Island" (bold, blue
'  fill, top level), 4 "Mechanical" (bold, grey fill), 5 a
'  plain task, 32 a task at the deepest outline level.
'
'  Not checked here: column widths, row heights, the logo's
'  position, frozen panes, or whether a column that was
'  hidden in the source comes out hidden.
'
'  Not part of /src. Headless only: tools\Test-Unit.ps1 -Suite flows
'===========================================================

Option Explicit
Option Private Module

Private Const FIXTURE As String = "Schedule_Sample.xlsx"
Private Const DATA_ROWS As Long = 135

Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub

Private Function RunSchedule(ByVal wb As Workbook, ByVal columnsCsv As String) As Worksheet
    Dim sheetsBefore As Long

    sheetsBefore = wb.Worksheets.Count
    modMain_CustSchedule.CreateCustSchedule_DoWork wb.Worksheets(1), ColumnsFromCsv(columnsCsv)

    Assert.AreEqual sheetsBefore + 1, wb.Worksheets.Count, "the flow should add exactly one sheet"
    Set RunSchedule = wb.Worksheets(wb.Worksheets.Count)
End Function

Private Function HeadingRow() As Long
    HeadingRow = SCHEDULE_HEADER_ROW_COUNT + 1
End Function

' The output row that source row sourceRow ends up on.
Private Function OutRow(ByVal sourceRow As Long) As Long
    OutRow = sourceRow + SCHEDULE_HEADER_ROW_COUNT
End Function

'@TestMethod("Schedule")
Public Sub Schedule_DefaultColumns_SheetColumnsAndRowCount()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunSchedule(wb, DefaultScheduleColumns())

    Assert.AreEqual "Customer Schedule", ws.Name
    Assert.AreEqual CsvAsHeadings(DefaultScheduleColumns()), HeadingsJoined(ws, HeadingRow()), _
                    "only the chosen columns, in the chosen order, on the heading row under the title block"
    Assert.AreEqual DATA_ROWS, DataRowCount(ws, HeadingRow()), "every source row comes across"
    Assert.AreEqual 1, ws.Shapes.Count, "one logo and nothing else"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Schedule")
Public Sub Schedule_ChosenOrder_IsHonoured()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo TestFail

    ' Three columns that are visible in the source, in an order that is not
    ' the source's (there it is Status, Tasks, Start Date) - #127.
    Set wb = OpenFixture(FIXTURE)
    Set ws = RunSchedule(wb, "Tasks, Start Date, Status")

    Assert.AreEqual "Tasks|Start Date|Status", HeadingsJoined(ws, HeadingRow())

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Schedule")
Public Sub Schedule_HiddenSourceColumn_ComesAcrossWithItsValues()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim src As Worksheet
    Dim srcCol As Long
    Dim outCol As Long
    On Error GoTo TestFail

    ' Predecessors is hidden in the source. Asked for second of four.
    Set wb = OpenFixture(FIXTURE)
    Set src = wb.Worksheets(1)
    Set ws = RunSchedule(wb, "Tasks, Predecessors, Start Date, Status")

    srcCol = ColumnOf(src, 1, "Predecessors")
    outCol = ColumnOf(ws, HeadingRow(), "Predecessors")
    Assert.IsTrue srcCol > 0, "the fixture should have a Predecessors column"
    Assert.IsTrue outCol > 0, "the chosen hidden column should be in the output"

    ' Source row 14 has "13FF" in Predecessors.
    Assert.AreEqual "13FF", CStr(src.Cells(14, srcCol).Value), "the fixture row this test relies on"
    If outCol > 0 Then
        Assert.AreEqual "13FF", CStr(ws.Cells(OutRow(14), outCol).Value), "its values come across"

        ' Where it lands, and whether it is still hidden, is not asserted:
        ' tests/README.md leaves both open. Reported instead.
        RecordTestNote "asked for 'Tasks, Predecessors, Start Date, Status'; got [" & _
                       HeadingsJoined(ws, HeadingRow()) & "]; Predecessors is in column " & outCol & _
                       ", hidden = " & CBool(ws.Columns(outCol).Hidden)
    End If

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Schedule")
Public Sub Schedule_FillsCleared_BoldKeptOnSummaryRows()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tasksCol As Long
    Dim r As Long
    Dim coloured As Long
    Dim boldRows As Long
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set ws = RunSchedule(wb, DefaultScheduleColumns())
    tasksCol = ColumnOf(ws, HeadingRow(), "Tasks")

    Assert.AreEqual "Boiler Island", CStr(ws.Cells(OutRow(2), tasksCol).Value), "source row 2 should be the first data row"

    For r = HeadingRow() + 1 To GetLastRow(ws)
        If HasColouredFill(ws.Cells(r, tasksCol)) Then coloured = coloured + 1
        If ws.Cells(r, tasksCol).Font.Bold = True Then boldRows = boldRows + 1
    Next r

    Assert.AreEqual 0, coloured, "every source fill is cleared"
    Assert.AreEqual 27, boldRows, "bold is what still marks the 27 summary rows"
    Assert.IsTrue ws.Cells(OutRow(4), tasksCol).Font.Bold = True, "'Mechanical' (grey in the source) stays bold"
    Assert.IsFalse ws.Cells(OutRow(5), tasksCol).Font.Bold = True, "a plain task stays plain"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Schedule")
Public Sub Schedule_NumberFormatsAndGroupingSurvive()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim src As Worksheet
    On Error GoTo TestFail

    Set wb = OpenFixture(FIXTURE)
    Set src = wb.Worksheets(1)
    Set ws = RunSchedule(wb, DefaultScheduleColumns())

    Assert.AreEqual src.Cells(5, ColumnOf(src, 1, "Start Date")).NumberFormat, _
                    ws.Cells(OutRow(5), ColumnOf(ws, HeadingRow(), "Start Date")).NumberFormat, "date format"
    Assert.AreEqual src.Cells(5, ColumnOf(src, 1, "% Complete")).NumberFormat, _
                    ws.Cells(OutRow(5), ColumnOf(ws, HeadingRow(), "% Complete")).NumberFormat, "percent format"
    Assert.IsTrue IsDate(ws.Cells(OutRow(5), ColumnOf(ws, HeadingRow(), "Start Date")).Value), _
                  "a start date is still a date, not text"

    ' Source row 32 is at the deepest outline level; row 2 at the top.
    Assert.AreEqual src.Rows(32).OutlineLevel, ws.Rows(OutRow(32)).OutlineLevel, "deepest grouping level"
    Assert.AreEqual src.Rows(2).OutlineLevel, ws.Rows(OutRow(2)).OutlineLevel, "top grouping level"
    Assert.IsTrue ws.Rows(OutRow(32)).OutlineLevel > ws.Rows(OutRow(2)).OutlineLevel, "and the two differ"

TestExit:
    CloseFixture wb
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
