Attribute VB_Name = "modHelpers_Headers"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Column-heading-row helpers: read a sheet's heading row into
'  a 1-based array (GetHeadingList) and find a heading's column
'  index (FindHeadingIndex). "heading" = the data table's
'  column-title row, not the TPD title block.
'===========================================================

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


' Returns the 1-based position of headingName in the headings array, or 0 if it
' isn't there. Case-insensitive.
'
' Defaults to a RIGHT-TO-LEFT scan: some TPD source EQ lists carry more than one
' column with the same label (e.g. a working "Vendor" column sits to the right
' of an imported/decoy one) and the rightmost is the authoritative one. Pass
' preferRightmost:=False for a plain left-to-right "first match wins" lookup.
Public Function FindHeadingIndex(headings As Variant, headingName As String, _
                                 Optional ByVal preferRightmost As Boolean = True) As Long
    Dim i As Long
    Dim firstIdx As Long, lastIdx As Long, stepDir As Long

    FindHeadingIndex = 0

    If preferRightmost Then
        firstIdx = UBound(headings): lastIdx = LBound(headings): stepDir = -1
    Else
        firstIdx = LBound(headings): lastIdx = UBound(headings): stepDir = 1
    End If

    For i = firstIdx To lastIdx Step stepDir
        If StrComp(CStr(headings(i)), headingName, vbTextCompare) = 0 Then
            FindHeadingIndex = i
            Exit Function
        End If
    Next i
End Function

