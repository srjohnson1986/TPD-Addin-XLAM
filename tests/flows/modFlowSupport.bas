Attribute VB_Name = "modFlowSupport"
'@Folder("TPD_Addin.Tests")

'===========================================================
'  Shared by the flow tests in this folder: opening a
'  fixture workbook, building the inputs a flow's _DoWork
'  routine takes, and reading the sheet it produces.
'
'  Flow tests run a whole flow against a copy of a fixture
'  in tests/fixtures and check the result sheet. They call
'  the flows' _DoWork routines directly, so no picker dialog
'  is shown, and they pass every EQ List option explicitly,
'  so a run never depends on what is saved in Set Defaults.
'
'  They run headlessly only - tools\Test-Unit.ps1 -Suite
'  flows - which copies the fixtures to a temp folder and
'  tells this module where they are. Nothing here saves a
'  workbook: fixtures are opened read-only and closed
'  without saving.
'
'  Not part of /src - never imported by BuildAddin, never
'  shipped in the .xlam.
'===========================================================

Option Explicit
Option Private Module

Private mFixtureFolder As String

' Called by the generated runner before any test.
Public Sub SetFixtureFolder(ByVal folderPath As String)
    mFixtureFolder = folderPath
    If Right$(mFixtureFolder, 1) = "\" Then mFixtureFolder = Left$(mFixtureFolder, Len(mFixtureFolder) - 1)
End Sub

' Opens a fixture read-only and makes it the active workbook, which is where
' the flows put their output.
Public Function OpenFixture(ByVal fileName As String) As Workbook
    Dim wb As Workbook

    If Len(mFixtureFolder) = 0 Then
        Err.Raise vbObjectError + 513, "modFlowSupport", _
                  "No fixture folder has been set. Run the flow tests with tools\Test-Unit.ps1 -Suite flows."
    End If

    Set wb = Workbooks.Open(mFixtureFolder & "\" & fileName, 0, True)
    wb.Activate
    Set OpenFixture = wb
End Function

' Never raises: tests call it from their error handler's exit path, where a
' second error would send them round again.
Public Sub CloseFixture(ByVal wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
End Sub

' "A, B, C" -> a Collection of "A", "B", "C", the shape _DoWork takes.
Public Function ColumnsFromCsv(ByVal csv As String) As Collection
    Dim parts() As String
    Dim i As Long
    Dim result As Collection

    Set result = New Collection
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then result.Add Trim$(parts(i))
    Next i

    Set ColumnsFromCsv = result
End Function

' All seven EQ List options as the add-in ships them - Remove blank rows and
' Add cell borders on, the rest off - with any number of them changed:
'     TogglesWith()                                     as shipped
'     TogglesWith(PREF_EQLIST_ADD_COUNT_COLUMN, True)   as shipped, plus EQ Count
' Every option is always present, so a run never falls back to the registry.
Public Function TogglesWith(ParamArray keyValuePairs() As Variant) As Object
    Dim toggles As Object
    Dim i As Long

    Set toggles = CreateObject("Scripting.Dictionary")
    toggles(PREF_EQLIST_REMOVE_BLANK_ROWS) = True
    toggles(PREF_EQLIST_REMOVE_PARENT_ROWS) = False
    toggles(PREF_EQLIST_ADD_COUNT_COLUMN) = False
    toggles(PREF_EQLIST_PLAIN_PARENT_ROWS) = False
    toggles(PREF_EQLIST_BOLD_PARENT_ROWS) = False
    toggles(PREF_EQLIST_ADD_CELL_BORDERS) = True
    toggles(PREF_EQLIST_ADD_COLUMN_FILTERS) = False

    For i = LBound(keyValuePairs) To UBound(keyValuePairs) - 1 Step 2
        toggles(CStr(keyValuePairs(i))) = CBool(keyValuePairs(i + 1))
    Next i

    Set TogglesWith = toggles
End Function

' ---- reading a result sheet ------------------------------------------------

' The heading row's text, "A|B|C".
Public Function HeadingsJoined(ByVal ws As Worksheet, ByVal headingRow As Long) As String
    Dim c As Long
    Dim lastCol As Long
    Dim joined As String

    lastCol = GetLastCol(ws, headingRow)
    For c = 1 To lastCol
        If c > 1 Then joined = joined & "|"
        joined = joined & CStr(ws.Cells(headingRow, c).Value)
    Next c

    HeadingsJoined = joined
End Function

' "A, B, C" -> "A|B|C", to compare with HeadingsJoined.
Public Function CsvAsHeadings(ByVal csv As String) As String
    Dim item As Variant
    Dim joined As String

    For Each item In ColumnsFromCsv(csv)
        If Len(joined) > 0 Then joined = joined & "|"
        joined = joined & CStr(item)
    Next item

    CsvAsHeadings = joined
End Function

' Rows below the heading row, down to the last row with anything in it.
Public Function DataRowCount(ByVal ws As Worksheet, ByVal headingRow As Long) As Long
    DataRowCount = GetLastRow(ws) - headingRow
End Function

' The column a heading is in, or 0.
Public Function ColumnOf(ByVal ws As Worksheet, ByVal headingRow As Long, ByVal heading As String) As Long
    Dim c As Long

    For c = 1 To GetLastCol(ws, headingRow)
        If StrComp(Trim$(CStr(ws.Cells(headingRow, c).Value)), heading, vbTextCompare) = 0 Then
            ColumnOf = c
            Exit Function
        End If
    Next c
End Function

' The first data row whose cell in col reads wanted, or 0.
Public Function FindDataRow(ByVal ws As Worksheet, ByVal headingRow As Long, _
                            ByVal col As Long, ByVal wanted As String) As Long
    Dim r As Long

    For r = headingRow + 1 To GetLastRow(ws)
        If Trim$(CStr(ws.Cells(r, col).Value)) = wanted Then
            FindDataRow = r
            Exit Function
        End If
    Next r
End Function

' Data rows with nothing in any column of the table.
Public Function BlankDataRowCount(ByVal ws As Worksheet, ByVal headingRow As Long) As Long
    Dim r As Long
    Dim lastCol As Long
    Dim found As Long

    lastCol = GetLastCol(ws, headingRow)
    For r = headingRow + 1 To GetLastRow(ws)
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) = 0 Then found = found + 1
    Next r

    BlankDataRowCount = found
End Function

' True when the cell has a background colour a reader would see: any fill
' other than none or plain white.
Public Function HasColouredFill(ByVal cell As Range) As Boolean
    If cell.Interior.ColorIndex = xlColorIndexNone Then Exit Function
    HasColouredFill = (cell.Interior.Color <> vbWhite)
End Function

' A one-line account of what is on a result sheet, for a failure message or a
' note: enough to tell an empty sheet from a table that is in the wrong place.
Public Function DescribeSheet(ByVal ws As Worksheet, ByVal headingRow As Long) As String
    Dim r As Long
    Dim hiddenRows As Long
    Dim hiddenCols As Long
    Dim c As Long
    Dim lastRow As Long

    lastRow = GetLastRow(ws)
    For r = 1 To lastRow
        If ws.Rows(r).Hidden Then hiddenRows = hiddenRows + 1
    Next r
    For c = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If ws.Columns(c).Hidden Then hiddenCols = hiddenCols + 1
    Next c

    DescribeSheet = "sheet '" & ws.Name & "': used range " & ws.UsedRange.Address(False, False) & _
                    ", last row with content " & lastRow & _
                    ", cells with content " & Application.WorksheetFunction.CountA(ws.Cells) & _
                    ", hidden rows " & hiddenRows & ", hidden columns " & hiddenCols & _
                    ", row " & headingRow & " reads [" & HeadingsJoined(ws, headingRow) & "]" & _
                    ", row " & headingRow + 1 & " starts [" & CStr(ws.Cells(headingRow + 1, 1).Value) & "]"
End Function

' The worksheet called sheetName in wb, or Nothing.
Public Function SheetNamed(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetNamed = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
