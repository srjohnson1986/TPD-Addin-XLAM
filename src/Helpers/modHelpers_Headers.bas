Attribute VB_Name = "modHelpers_Headers"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Function GetHeadingList(ws As Worksheet, headingsRow As Long) As Variant
    Dim lastCol As Long
    Dim arr() As String
    Dim c As Long

    lastCol = GetLastCol(ws, headingsRow)
    ReDim arr(1 To lastCol)

    For c = 1 To lastCol
        arr(c) = CStr(ws.Cells(headingsRow, c).value)
    Next c

    GetHeadingList = arr
End Function


Public Function FindHeadingIndex(headings As Variant, headingName As String) As Long
    Dim i As Long
    FindHeadingIndex = 0

    ' Search from RIGHT to LEFT so the real Vendor column is found
    For i = UBound(headings) To LBound(headings) Step -1
        If StrComp(CStr(headings(i)), headingName, vbTextCompare) = 0 Then
            FindHeadingIndex = i
            Exit Function
        End If
    Next i
End Function

