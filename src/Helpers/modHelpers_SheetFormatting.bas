Attribute VB_Name = "modHelpers_SheetFormatting"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Visual formatting for generated EQ / schedule / split
'  sheets: fonts, alignment, borders, autofit, and freeze
'  panes. Pure presentation - no data is moved here.
'===========================================================

Option Explicit

' Formats a generated data table: Arial 10, thin borders, and every column
' centre-aligned EXCEPT leftAlignHeading (the sheet's one long-text column),
' which keeps its natural left alignment. The EQ List flow leaves the default
' "Description"; the Customer Schedule flow passes "Tasks". If that heading
' isn't on the row, every column is centred.
Public Sub FormatEQSheet(ws As Worksheet, headingsRow As Long, _
                         Optional ByVal leftAlignHeading As String = "Description")
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

    descCol = Application.Match(leftAlignHeading, ws.Rows(headingsRow), 0)

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

End Sub

' Strips the background fill carried over from the source sheet by a
' row-for-row copy. Call BEFORE any columns are dropped, so the heading
' lookups below still see the full source column set (the picker and the
' one-click flows would otherwise disagree - the old column-A guard here
' kept fill wherever the leftmost *kept* column was blank, and the two
' flows order their columns differently, #130).
'
' preserveStructuralRows:=True (the Customer EQ List) keeps the fill on
' rows that act as a visual grouping cue:
'   - PURCHASED = "PARENT", or
'   - INTERNAL ID and CUSTOMER ID both blank.
' All three headings are looked up case-insensitively and are optional: no
' PURCHASED column falls back to the both-IDs-blank test, and with none of
' the three present every data row is cleared. The Customer Schedule flow
' takes the default (False) and clears every row.
Public Sub StripDataRowFill(ws As Worksheet, headingsRow As Long, _
                            Optional ByVal preserveStructuralRows As Boolean = False)
    Dim headings As Variant
    Dim purchasedCol As Long, internalIdCol As Long, customerIdCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim keepFill As Boolean

    If preserveStructuralRows Then
        headings = modHelpers_Headers.GetHeadingList(ws, headingsRow)
        purchasedCol = modHelpers_Headers.FindHeadingIndex(headings, "PURCHASED", preferRightmost:=False)
        internalIdCol = modHelpers_Headers.FindHeadingIndex(headings, "INTERNAL ID", preferRightmost:=False)
        customerIdCol = modHelpers_Headers.FindHeadingIndex(headings, "CUSTOMER ID", preferRightmost:=False)
    End If

    lastRow = GetLastRow(ws)

    For r = headingsRow To lastRow
        keepFill = False

        If preserveStructuralRows And r > headingsRow Then
            If purchasedCol > 0 Then
                If UCase$(Trim$(CStr(ws.Cells(r, purchasedCol).value))) = "PARENT" Then keepFill = True
            End If
            If Not keepFill Then
                If internalIdCol > 0 And customerIdCol > 0 Then
                    If Trim$(CStr(ws.Cells(r, internalIdCol).value)) = "" _
                       And Trim$(CStr(ws.Cells(r, customerIdCol).value)) = "" Then keepFill = True
                End If
            End If
        End If

        If Not keepFill Then ws.Rows(r).Interior.ColorIndex = xlNone
    Next r

End Sub

Public Sub FormatSplitSheet(ws As Worksheet, headingsRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    lastRow = GetLastRow(ws)
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
