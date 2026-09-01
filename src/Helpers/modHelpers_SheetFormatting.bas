Attribute VB_Name = "modHelpers_SheetFormatting"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Sub FormatEQSheet(ws As Worksheet, dataHeaderRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim descCol As Variant

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells(dataHeaderRow, ws.Columns.Count).End(xlToLeft).Column

    Set dataRange = ws.Range(ws.Cells(dataHeaderRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With

    descCol = Application.Match("Description", ws.Rows(dataHeaderRow), 0)

    If Not IsError(descCol) Then
        If descCol > 1 Then
            ws.Range(ws.Cells(dataHeaderRow, 1), ws.Cells(lastRow, descCol - 1)).HorizontalAlignment = xlCenter
        End If
        If descCol < lastCol Then
            ws.Range(ws.Cells(dataHeaderRow, descCol + 1), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
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
    For r = dataHeaderRow To lastRow
        If Trim(CStr(ws.Cells(r, 1).value)) <> "" Then
            ws.Rows(r).Interior.ColorIndex = xlNone
        End If
    Next r


End Sub

Public Sub FormatSplitSheet(ws As Worksheet, headerRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With

    AutoFitUsedColumns (ws)

End Sub


Public Sub SafeFreezePanes(ws As Worksheet, headerRow As Long)
    On Error Resume Next
    ws.Activate
    ws.Range("A" & headerRow + 1).Select
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
