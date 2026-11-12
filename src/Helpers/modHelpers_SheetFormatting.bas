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
'
' Named FormatEQSheet until the Schedule flow started calling it too - it has
' never been EQ-specific, only EQ-first.
'
' addBorders defaults True (the historic behaviour, kept for the Schedule
' caller and every existing path). The EQ List flow passes the user's
' "Add cell borders" toggle so the border can be turned off (#122).
Public Sub FormatDataTable(ws As Worksheet, headingsRow As Long, _
                         Optional ByVal leftAlignHeading As String = "Description", _
                         Optional ByVal addBorders As Boolean = True)
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

    If addBorders Then
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If

End Sub

' Strips the formatting a row-for-row copy carries over from the source
' sheet: the background fill, and on data rows the font emphasis too (bold,
' italic, underline, strikethrough, colour back to automatic - #169). A
' whole-row Rows.Copy brings every cell's look across; Split Sheets copies
' values only, so its output never had this. Font colour goes with the fill
' because a source rule like white-on-dark would otherwise leave invisible
' text once the fill is gone. The heading row loses its fill but keeps its
' font.
'
' Call BEFORE any columns are dropped, so the heading lookups below still see
' the full source column set (the picker and the one-click flows would
' otherwise disagree - the old column-A guard here kept fill wherever the
' leftmost *kept* column was blank, and the two flows order their columns
' differently, #130).
'
' preserveStructuralRows:=True (the Customer EQ List) leaves rows that act as
' a visual grouping cue untouched - fill and font both, so a bold PARENT row
' stays bold ("Plain PARENT rows", #121, is the switch for that). Which rows
' count is IsStructuralRow's call. The Customer Schedule flow takes the
' default (False) and clears every row.
'
' clearFontEmphasis:=False keeps the font and strips fill only - the Customer
' Schedule passes it, because once every fill is gone bold is the only thing
' still marking its summary rows.
Public Sub StripCarriedRowFormatting(ws As Worksheet, headingsRow As Long, _
                                     Optional ByVal preserveStructuralRows As Boolean = False, _
                                     Optional ByVal clearFontEmphasis As Boolean = True)
    Dim headings As Variant
    Dim purchasedCol As Long, internalIdCol As Long, customerIdCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim keepRow As Boolean

    If preserveStructuralRows Then
        headings = modHelpers_Headers.GetHeadingList(ws, headingsRow)
        purchasedCol = modHelpers_Headers.FindHeadingIndex(headings, "PURCHASED", preferRightmost:=False)
        internalIdCol = modHelpers_Headers.FindHeadingIndex(headings, "INTERNAL ID", preferRightmost:=False)
        customerIdCol = modHelpers_Headers.FindHeadingIndex(headings, "CUSTOMER ID", preferRightmost:=False)
    End If

    lastRow = GetLastRow(ws)

    ws.Rows(headingsRow).Interior.ColorIndex = xlNone

    For r = headingsRow + 1 To lastRow
        keepRow = False
        If preserveStructuralRows Then
            keepRow = IsStructuralRow(purchasedCol > 0, CellText(ws, r, purchasedCol), _
                                      internalIdCol > 0 And customerIdCol > 0, _
                                      CellText(ws, r, internalIdCol), CellText(ws, r, customerIdCol))
        End If

        If Not keepRow Then
            ws.Rows(r).Interior.ColorIndex = xlNone
            If clearFontEmphasis Then
                With ws.Rows(r).Font
                    .Bold = False
                    .Italic = False
                    .Underline = xlUnderlineStyleNone
                    .Strikethrough = False
                    .ColorIndex = xlColorIndexAutomatic
                End With
            End If
        End If
    Next r

End Sub

' True when an EQ List data row is a grouping cue whose source look is kept
' (#130). Pure - the caller reads the cells - so it can be unit-tested.
'
'   - With a PURCHASED column, that column decides alone: PARENT rows only.
'     An item row is an item even when both its IDs are blank (#170 - a real
'     "For Approval" row with no IDs used to keep its yellow).
'   - With no PURCHASED column, fall back to INTERNAL ID and CUSTOMER ID
'     both blank (hasIdCols = both headings present).
'   - With neither, nothing is structural.
Public Function IsStructuralRow(ByVal hasPurchasedCol As Boolean, ByVal purchasedText As String, _
                                ByVal hasIdCols As Boolean, ByVal internalIdText As String, _
                                ByVal customerIdText As String) As Boolean
    If hasPurchasedCol Then
        IsStructuralRow = (UCase$(Trim$(purchasedText)) = "PARENT")
    ElseIf hasIdCols Then
        IsStructuralRow = (Trim$(internalIdText) = vbNullString And Trim$(customerIdText) = vbNullString)
    End If
End Function

' The cell's text for IsStructuralRow, or "" when col is 0 (heading absent).
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long) As String
    If col > 0 Then CellText = CStr(ws.Cells(r, col).value)
End Function

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
    AutoFitDataRowHeights ws, headingsRow

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

' Autofits the row height of the data table (headingsRow through the last row
' of data) so wrapped/multi-line cell text isn't clipped (#162). Scoped to
' headingsRow..lastRow the same way FormatDataTable's dataRange is, so it
' never touches rows above headingsRow.
'
' Call this BEFORE inserting any header/logo block (InsertEQHeaderBlock /
' InsertDefaultCustScheduleHeader) - those just insert blank rows above and
' shift everything else down, so heights set here move down with their rows.
' Calling it after would mean re-deriving where the block landed instead.
'
' Also call it AFTER AutoFitUsedColumns - row autofit measures wrapped text
' against the CURRENT column widths, so columns need to be final first.
Public Sub AutoFitDataRowHeights(ByVal ws As Worksheet, ByVal headingsRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = GetLastRow(ws)
    lastCol = GetLastCol(ws, headingsRow)

    ws.Range(ws.Cells(headingsRow, 1), ws.Cells(lastRow, lastCol)).Rows.AutoFit
End Sub
