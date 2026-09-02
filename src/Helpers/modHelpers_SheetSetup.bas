Attribute VB_Name = "modHelpers_SheetSetup"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Sub InsertEQHeaderBlock(ws As Worksheet)

    InsertRows ws, 6

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

    ' Inserts 6 rows (currently) and adds the EQ header text
    InsertEQHeaderBlock ws

    ' number of rows inserted (modify as EQ header block changes)
    Dim numberOfRows As Long
    numberOfRows = 6

    ' headingsRow is after the number of inserted rows
    headingsRow = numberOfRows + headingsRow

    ' Insert logo on this worksheet on the headingsRow, right-top aligned, scaled to the number of rows we inserted
    InsertDefaultLogo ws, headingsRow, "right-top", numberOfRows

    SafeFreezePanes ws, headingsRow

End Sub
