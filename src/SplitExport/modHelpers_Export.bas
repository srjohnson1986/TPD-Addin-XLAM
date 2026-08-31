Attribute VB_Name = "modHelpers_Export"
'@Folder("TPD_Addin.SplitExport")

Option Explicit

Public Function EnsureExportFolder(wb As Workbook) As String
    Dim basePath As String
    Dim baseName As String
    Dim exportFolder As String

    basePath = wb.Path
    If basePath = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Function
    End If

    baseName = SanitizeSheetName(wb.Worksheets(1).name)
    exportFolder = basePath & "\" & baseName

    If Dir(exportFolder, vbDirectory) = "" Then
        MkDir exportFolder
    End If

    EnsureExportFolder = exportFolder
End Function

Public Sub ExportSheetToXLSX(ws As Worksheet, folder As String, suffix As String)
    Dim safeName As String
    Dim newWb As Workbook
    Dim fullName As String

    safeName = ws.name
    If Len(suffix) > 0 Then
        safeName = safeName & " - " & suffix
    End If
    safeName = SanitizeFileText(safeName)

    ws.Copy
    Set newWb = ActiveWorkbook

    fullName = folder & "\" & safeName & ".xlsx"
    newWb.SaveAs Filename:=fullName, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
End Sub

