Attribute VB_Name = "modHelpers_SheetFormatting"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Visual formatting for generated EQ / split sheets: fonts,
'  alignment, borders, autofit, and freeze panes. Pure
'  presentation - no data is moved here.
'===========================================================

Option Explicit

Public Sub FormatEQSheet(ws As Worksheet, headingsRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim descCol As Variant

    lastRow = GetLastRow(ws)
    lastCol = GetLastCol(ws, headingsRow)

    Set dataRange = ws.Range(ws.Cells(headingsRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With

    descCol = Application.Match("Description", ws.Rows(headingsRow), 0)

    If Not IsError(descCol) Then
        If descCol > 1 Then
            ws.Range(ws.Cells(headingsRow, 1), ws.Cells(lastRow, descCol - 1)).HorizontalAlignment = xlCenter
        End If
        If descCol < lastCol Then
            ws.Range(ws.Cells(headingsRow, descCol + 1), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
        End If
    Else
        dataRange.HorizontalAlignment = xlCenter
    End If

    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    Dim r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = headingsRow To lastRow
        If Trim(CStr(ws.Cells(r, 1).value)) <> "" Then
            ws.Rows(r).Interior.ColorIndex = xlNone
        End If
    Next r


End Sub

Public Sub FormatSplitSheet(ws As Worksheet, headingsRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = GetLastCol(ws, headingsRow)

    Set dataRange = ws.Range(ws.Cells(headingsRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With

    ws.Rows(headingsRow).Font.Bold = True

    AutoFitUsedColumns ws

End Sub


Public Sub SafeFreezePanes(ws As Worksheet, headingsRow As Long)
    On Error Resume Next
    ws.Activate
    ws.Range("A" & headingsRow + 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Public Sub AutoFitUsedColumns(ws As Worksheet)
    Dim usedCols As Range

    ' If the sheet has no used range, exit safely
    If ws.UsedRange Is Nothing Then Exit Sub

    ' Get only the used columns
    Set usedCols = ws.UsedRange.Columns

    ' Autofit only the populated columns
    usedCols.AutoFit
End Sub
