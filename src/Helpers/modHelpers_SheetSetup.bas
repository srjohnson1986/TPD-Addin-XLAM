Attribute VB_Name = "modHelpers_SheetSetup"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Builds the TPD title block (header) at the top of a sheet
'  and orchestrates the default EQ List / Schedule header
'  setup (rows, date text, logo, freeze panes). "header" here
'  is the title block, not the column heading row.
'===========================================================

Option Explicit

' Number of rows InsertEQHeaderBlock inserts and populates (A1:B6). Anything
' that needs the row where the data headings land after the block is inserted
' should use EQ_HEADER_ROW_COUNT + 1 rather than a bare literal. If the block
' grows, update this constant AND the A1:B6 layout in InsertEQHeaderBlock.
Public Const EQ_HEADER_ROW_COUNT As Long = 6

Public Sub InsertEQHeaderBlock(ws As Worksheet)

    InsertRows ws, EQ_HEADER_ROW_COUNT

    ws.Range("A1").value = "EQUIPMENT LIST"
    ws.Range("A2").value = "Customer"
    ws.Range("A3").value = "Project"
    ws.Range("A4").value = "PM"
    ws.Range("A5").value = "Date"
    ws.Range("B5").value = GetTodaysDate
    ws.Range("A6").value = "Rev"

    With ws.Range("A1")
        .Font.name = "Aptos Narrow"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("A2:B6")
        .Font.name = "Aptos Narrow"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub InsertRows(ws As Worksheet, numberOfRows As Long)
    If numberOfRows <= 0 Then Exit Sub
    ws.Rows("1:" & numberOfRows).Insert Shift:=xlDown
End Sub


Public Function GetTodaysDate() As String
    GetTodaysDate = Format(Date, "mm-dd-yyyy")
End Function

Public Sub InsertDefaultCustScheduleHeader(ws As Worksheet)

    ' on the active worksheet, get the last column that is populated in row 1
    Dim lastColumn As Long
    lastColumn = GetLastCol(ws, 1)

    ' number of rows to insert
    Dim numberOfRows As Long
    numberOfRows = 5

    ' on the active worksheet, insert numberOfRows
    InsertRows ws, numberOfRows

    ' headingsRow is one more than the number of inserted rows
    Dim headingsRow As Long
    headingsRow = numberOfRows + 1

    ' date text 2 rows above headingsRow
    Dim dateTextRow As Long
    dateTextRow = headingsRow - 2

    ' Left date text cell (column before last)
    With ws.Cells(dateTextRow, lastColumn - 1)
        .value = "Date"
        .Font.name = "Arial"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With

    ' Right date text cell (last column)
    With ws.Cells(dateTextRow, lastColumn)
        .value = GetTodaysDate()
        .Font.name = "Arial"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlRight
    End With

    ' Insert logo on this worksheet on row 1, left-aligned, scaled to the number of rows we inserted
    InsertDefaultLogo ws, 1, "left", numberOfRows

End Sub

Public Sub InsertDefaultCustEQHeader(ws As Worksheet)

    Dim headingsRow As Long
    ' Assume the column headings are on row 1 to start.
    headingsRow = 1

    ' On the active worksheet, get the last column that is populated in the row of headings
    Dim lastColumn As Long
    lastColumn = GetLastCol(ws, headingsRow)

    ' Formats the sheet for equipment lists using assuming that row 1 is the header.
    FormatEQSheet ws, headingsRow
    AutoFitUsedColumns ws

    ' Inserts the EQ header block (EQ_HEADER_ROW_COUNT rows) and its text
    InsertEQHeaderBlock ws

    ' headingsRow moves down by the number of rows the block inserted
    headingsRow = EQ_HEADER_ROW_COUNT + headingsRow

    ' Insert logo on this worksheet on the headingsRow, right-top aligned, scaled to the number of rows we inserted
    InsertDefaultLogo ws, headingsRow, "right-top", EQ_HEADER_ROW_COUNT

    SafeFreezePanes ws, headingsRow

End Sub
