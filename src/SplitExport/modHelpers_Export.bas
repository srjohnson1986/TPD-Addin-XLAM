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

' Copies one worksheet out to its own .xlsx file.
' Returns "" on success, or a "'SheetName': reason" string on failure.
' On any failure the temporary ws.Copy workbook is closed without saving,
' so it is never left open and visible.
Public Function ExportSheetToXLSX(ws As Worksheet, folder As String, suffix As String) As String
    Dim safeName As String
    Dim newWb As Workbook
    Dim fullName As String

    safeName = ws.name
    If Len(suffix) > 0 Then
        safeName = safeName & " - " & suffix
    End If
    safeName = SanitizeFileText(safeName)

    ' Excel refuses SaveAs when another workbook with the same file name is
    ' already open (even in a different folder). This happens when a sheet is
    ' named like the source workbook's file. Disambiguate rather than fail.
    If IsWorkbookOpen(safeName & ".xlsx") Then
        safeName = safeName & " (export)"
    End If

    fullName = folder & "\" & safeName & ".xlsx"

    On Error GoTo Failed

    ws.Copy
    Set newWb = ActiveWorkbook
    newWb.SaveAs Filename:=fullName, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Exit Function

Failed:
    Dim reason As String
    reason = Err.Description

    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0

    ExportSheetToXLSX = "'" & ws.name & "': " & reason
End Function
