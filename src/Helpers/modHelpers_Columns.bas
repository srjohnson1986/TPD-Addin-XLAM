Attribute VB_Name = "modHelpers_Columns"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Column-oriented helpers: case-insensitive collection
'  membership, deleting columns whose heading isn't in a
'  selected list, unique values in a column, and copying rows
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

Public Function GetUniqueValuesInColumn(ws As Worksheet, colIndex As Long) As Collection

    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim exists As Boolean
    Dim v As Variant

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    For i = 2 To lastRow   ' skip heading row
        cellValue = NormalizeCellText(ws.Cells(i, colIndex).value)


        If Len(cellValue) > 0 Then
            exists = False

            For Each v In result
                If StrComp(v, cellValue, vbTextCompare) = 0 Then
                    exists = True
                    Exit For
                End If
            Next v

            If Not exists Then
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
    Dim lastCol As Long
    Dim colMap As Object
    Dim i As Long
    Dim idx As Variant
    Dim destCol As Long
    Dim destRow As Long
    Dim cell As Range

    lastRow = GetLastRow(wsSource)

    lastCol = GetLastCol(wsSource, headingsRow)

    ' Build column map
    Set colMap = CreateObject("Scripting.Dictionary")
    For i = 1 To lastCol
        If CollectionContainsText(selectedHeadings, CStr(wsSource.Cells(headingsRow, i).value)) Then
            colMap.Add i, wsSource.Cells(headingsRow, i).value
        End If
    Next i

    ' Copy headings row
    destCol = 1
    For Each idx In colMap.Keys
        wsDest.Cells(headingsRow, destCol).value = wsSource.Cells(headingsRow, CLng(idx)).value
        destCol = destCol + 1
    Next idx

    ' Copy rows. Formatting / bold headings / freeze panes / autofit are the
    ' caller's job (FormatSplitSheet + SafeFreezePanes) - this routine only
    ' moves data.
    destRow = headingsRow + 1

    For Each cell In wsSource.Range(wsSource.Cells(headingsRow + 1, groupColIndex), wsSource.Cells(lastRow, groupColIndex))

        If StrComp(NormalizeCellText(cell.value), NormalizeCellText(matchValue), vbTextCompare) = 0 Then
    
            destCol = 1
            For Each idx In colMap.Keys
                wsDest.Cells(destRow, destCol).value = wsSource.Cells(cell.Row, CLng(idx)).value
                destCol = destCol + 1
            Next idx
    
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


