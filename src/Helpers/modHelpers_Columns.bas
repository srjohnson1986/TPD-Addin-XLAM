Attribute VB_Name = "modHelpers_Columns"
'@Folder("TPD_Addin.Helpers")

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

Public Sub DeleteUnselectedColumnsByHeader(ws As Worksheet, selectedHeaders As Collection, headerRow As Long)
    Dim lastCol As Long
    Dim colHeader As String
    Dim c As Long

    lastCol = GetLastCol(ws, headerRow)

    For c = lastCol To 1 Step -1
        colHeader = CStr(ws.Cells(headerRow, c).value)
        If Not CollectionContainsText(selectedHeaders, colHeader) Then
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

    For i = 2 To lastRow   ' skip header row
        cellValue = CleanValue(ws.Cells(i, colIndex).value)


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
    ByVal selectedHeaders As Collection, _
    ByVal matchValue As String, _
    ByVal headerRow As Long)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim colMap As Object
    Dim i As Long
    Dim idx As Variant
    Dim destCol As Long
    Dim destRow As Long
    Dim cell As Range

    lastRow = GetLastRow(wsSource)

    lastCol = GetLastCol(wsSource, headerRow)

    ' Build column map
    Set colMap = CreateObject("Scripting.Dictionary")
    For i = 1 To lastCol
        If CollectionContainsText(selectedHeaders, CStr(wsSource.Cells(headerRow, i).value)) Then
            colMap.Add i, wsSource.Cells(headerRow, i).value
        End If
    Next i

    ' Copy header
    destCol = 1
    For Each idx In colMap.Keys
        wsDest.Cells(headerRow, destCol).value = wsSource.Cells(headerRow, CLng(idx)).value
        destCol = destCol + 1
    Next idx

    wsDest.Rows(headerRow).Font.Bold = True

    ' Freeze header
    wsDest.Activate
    wsDest.Range("A" & headerRow + 1).Select
    ActiveWindow.FreezePanes = True

    ' Copy rows
    destRow = headerRow + 1
    
    For Each cell In wsSource.Range(wsSource.Cells(headerRow + 1, groupColIndex), wsSource.Cells(lastRow, groupColIndex))

        If StrComp(CleanValue(cell.value), CleanValue(matchValue), vbTextCompare) = 0 Then
    
            destCol = 1
            For Each idx In colMap.Keys
                wsDest.Cells(destRow, destCol).value = wsSource.Cells(cell.Row, CLng(idx)).value
                destCol = destCol + 1
            Next idx
    
            destRow = destRow + 1
        End If
    
    Next cell

    
'    For Each cell In wsSource.Range(wsSource.Cells(headerRow + 1, groupColIndex), wsSource.Cells(lastRow, groupColIndex))
'        If CStr(cell.value) = matchValue Then
'            destCol = 1
'            For Each idx In colMap.Keys
'                wsDest.Cells(destRow, destCol).value = wsSource.Cells(cell.Row, CLng(idx)).value
'                destCol = destCol + 1
'            Next idx
'            destRow = destRow + 1
'        End If
'    Next cell

    wsDest.Columns.AutoFit

End Sub

Public Sub CopyAllRowsPreserveGroups(wsSource As Worksheet, wsDest As Worksheet)
    wsSource.Rows("1:" & wsSource.Rows.Count).Copy wsDest.Rows(1)
End Sub


