Attribute VB_Name = "modHelpers_Headers"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Function GetHeaderList(ws As Worksheet, headerRow As Long) As Variant
    Dim lastCol As Long
    Dim arr() As String
    Dim c As Long

    lastCol = GetLastCol(ws, headerRow)
    ReDim arr(1 To lastCol)

    For c = 1 To lastCol
        arr(c) = CStr(ws.Cells(headerRow, c).value)
    Next c

    GetHeaderList = arr
End Function

'Public Function FindHeaderIndex(headers As Variant, headerName As String) As Long
'    Dim i As Long
'    For i = LBound(headers) To UBound(headers)
'        If StrComp(headers(i), headerName, vbTextCompare) = 0 Then
'            FindHeaderIndex = i
'            Exit Function
'        End If
'    Next i
'End Function

Public Function FindHeaderIndex(headers As Variant, headerName As String) As Long
    Dim i As Long
    FindHeaderIndex = 0

    ' Search from RIGHT to LEFT so the real Vendor column is found
    For i = UBound(headers) To LBound(headers) Step -1
        If StrComp(CStr(headers(i)), headerName, vbTextCompare) = 0 Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
End Function

