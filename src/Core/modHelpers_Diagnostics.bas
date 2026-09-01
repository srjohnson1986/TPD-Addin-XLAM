Attribute VB_Name = "modHelpers_Diagnostics"
'@Folder("TPD_Addin.Core")

Public Function SheetExists2(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(name) Is Nothing
End Function
