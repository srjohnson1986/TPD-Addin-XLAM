Attribute VB_Name = "modHelpers_Diagnostics"
'@Folder("TPD_Addin.Core")

Option Explicit

' Returns True if a sheet with the given name exists in ThisWorkbook (the add-in).
' For checking an arbitrary workbook, use SheetExists(wb, name) in modHelpers_Workbook.
Public Function SheetExists2(name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExists2 = Not ws Is Nothing
End Function
