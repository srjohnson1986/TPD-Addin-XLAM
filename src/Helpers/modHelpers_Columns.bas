Attribute VB_Name = "modHelpers_Columns"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Column-oriented helpers: case-insensitive collection
'  membership, deleting columns whose heading isn't in a
'  selected list, reordering the kept columns to a chosen
'  order, unique values in a column, and copying rows
'  between sheets (all rows, or only rows matching a value in
'  the group column) projected onto a chosen set of columns.
'===========================================================

Option Explicit

Public Function CollectionContainsText(col As Collection, value As String) As Boolean
    Dim v As Variant
    For Each v In col
        If StrComp(CStr(v), value, vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next v
End Function

Public Sub DeleteUnselectedColumnsByHeading(ws As Worksheet, selectedHeadings As Collection, headingsRow As Long)
    Dim lastCol As Long
    Dim colHeading As String
    Dim c As Long

    lastCol = GetLastCol(ws, headingsRow)

    For c = lastCol To 1 Step -1
        colHeading = CStr(ws.Cells(headingsRow, c).value)
        If Not CollectionContainsText(selectedHeadings, colHeading) Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub

' Reorders ws's columns so their headings on headingsRow run left-to-right in
' orderedHeadings order, moving each wanted column into place with cut/insert.
' Headings not found at/after the current target position are skipped (already
' placed, or absent); any columns whose heading isn't in orderedHeadings are
' left untouched to the right of the ordered block.
'
' Pairs with DeleteUnselectedColumnsByHeading: that drops the unselected
' columns but leaves the survivors in source order, so the EQ List / Schedule
' flows call this straight after to honour the order set in Set TPD Defaults
' (#127). Cheap no-op when orderedHeadings is already in source order (e.g. the
' checkbox pickers, which can't express a custom order).
Public Sub ReorderColumnsByHeading(ByVal ws As Worksheet, ByVal orderedHeadings As Collection, ByVal headingsRow As Long)
    Dim targetCol As Long
    Dim h As Variant
    Dim j As Long
    Dim lastCol As Long
    Dim curIdx As Long

    targetCol = 1
    For Each h In orderedHeadings
        lastCol = GetLastCol(ws, headingsRow)
        curIdx = 0
        For j = targetCol To lastCol
            If StrComp(CStr(ws.Cells(headingsRow, j).value), CStr(h), vbTextCompare) = 0 Then
                curIdx = j
                Exit For
            End If
        Next j

        If curIdx = targetCol Then
            targetCol = targetCol + 1
        ElseIf curIdx > targetCol Then
            ws.Columns(curIdx).Cut
            ws.Columns(targetCol).Insert Shift:=xlToRight
            targetCol = targetCol + 1
        End If
    Next h
End Sub

' Returns the subset of `wanted` heading names that actually appear on
' headingsRow of ws, preserving `wanted` order and its (possibly dirty)
' spelling. The one-click flows use this to avoid handing
' DeleteUnselectedColumnsByHeading a selection that matches nothing on
' the source sheet - which would delete every column (#96).
Public Function HeadingsPresentOnSheet( _
        ByVal ws As Worksheet, _
        ByVal headingsRow As Long, _
        ByVal wanted As Collection) As Collection

    Dim present As New Collection
    Dim headings As Variant
    Dim onSheet As New Collection
    Dim i As Long
    Dim w As Variant

    headings = modHelpers_Headers.GetHeadingList(ws, headingsRow)
    For i = LBound(headings) To UBound(headings)
        onSheet.Add CStr(headings(i))
    Next i

    For Each w In wanted
        If CollectionContainsText(onSheet, CStr(w)) Then present.Add CStr(w)
    Next w

    Set HeadingsPresentOnSheet = present
End Function

Public Function GetUniqueValuesInColumn(ws As Worksheet, colIndex As Long) As Collection

    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    For i = 2 To lastRow   ' skip heading row
        cellValue = NormalizeCellText(ws.Cells(i, colIndex).value)

        If Len(cellValue) > 0 Then
            If Not CollectionContainsText(result, cellValue) Then
                result.Add cellValue
            End If
        End If
    Next i

    Set GetUniqueValuesInColumn = result
End Function


Public Sub CopyFilteredRowsByColumns( _
    ByVal wsSource As Worksheet, _
    ByVal wsDest As Worksheet, _
    ByVal groupColIndex As Long, _
    ByVal selectedHeadings As Collection, _
    ByVal matchValue As String, _
    ByVal headingsRow As Long)

    Dim lastRow As Long
    Dim headings As Variant
    Dim colMap As Object
    Dim h As Variant
    Dim srcIdx As Long
    Dim srcColIndex As Variant
    Dim destCol As Long
    Dim destRow As Long
    Dim cell As Range

    lastRow = GetLastRow(wsSource)

    ' Build the column map by walking selectedHeadings in order (not the source
    ' columns) so the output columns come out in the chosen order rather than
    ' the source-sheet order (#127). Scripting.Dictionary preserves insertion
    ' order, so colMap.Keys below is already in selectedHeadings order.
    headings = modHelpers_Headers.GetHeadingList(wsSource, headingsRow)
    Set colMap = CreateObject("Scripting.Dictionary")
    For Each h In selectedHeadings
        srcIdx = modHelpers_Headers.FindHeadingIndex(headings, CStr(h), preferRightmost:=True)
        If srcIdx > 0 Then
            If Not colMap.exists(srcIdx) Then colMap.Add srcIdx, headings(srcIdx)
        End If
    Next h

    ' Copy headings row
    destCol = 1
    For Each srcColIndex In colMap.Keys
        wsDest.Cells(headingsRow, destCol).value = wsSource.Cells(headingsRow, CLng(srcColIndex)).value
        destCol = destCol + 1
    Next srcColIndex

    ' Copy rows. Formatting / bold headings / freeze panes / autofit are the
    ' caller's job (FormatSplitSheet + SafeFreezePanes) - this routine only
    ' moves data.
    destRow = headingsRow + 1

    For Each cell In wsSource.Range(wsSource.Cells(headingsRow + 1, groupColIndex), wsSource.Cells(lastRow, groupColIndex))

        If StrComp(NormalizeCellText(cell.value), NormalizeCellText(matchValue), vbTextCompare) = 0 Then

            destCol = 1
            For Each srcColIndex In colMap.Keys
                wsDest.Cells(destRow, destCol).value = wsSource.Cells(cell.Row, CLng(srcColIndex)).value
                destCol = destCol + 1
            Next srcColIndex

            destRow = destRow + 1
        End If

    Next cell

End Sub

' Copies every row of wsSource onto wsDest verbatim. A whole-row copy (rather
' than a range copy) carries Excel's row grouping / outline levels across too,
' which the Customer EQ List needs.
Public Sub CopyEntireSheetRows(wsSource As Worksheet, wsDest As Worksheet)
    wsSource.Rows("1:" & wsSource.Rows.Count).Copy wsDest.Rows(1)
End Sub


