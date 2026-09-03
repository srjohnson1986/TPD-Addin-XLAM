Attribute VB_Name = "modHelpers_Export"
'@Folder("TPD_Addin.SplitExport")

'===========================================================
'  Helpers behind "Save Each Sheet to XLSX": locate/create the
'  per-workbook export folder (EnsureExportFolder) and write a
'  single worksheet out to its own .xlsx (ExportSheetToXLSX,
'  which returns "" on success or a reason string on failure).
'===========================================================

Option Explicit

' Returns the path of the per-workbook export folder (created next to the
' workbook if missing), or "" - after a message - if the workbook is unsaved
' or the folder can't be created.
Public Function EnsureExportFolder(wb As Workbook) As String
    Dim basePath As String
    Dim baseName As String
    Dim exportFolder As String
    Dim dotPos As Long

    basePath = wb.Path
    If basePath = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Function
    End If

    ' Name the folder after the workbook file (minus extension), not whatever
    ' sheet happens to be leftmost in tab order.
    baseName = wb.name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
    baseName = SanitizeFileText(baseName)
    If Len(baseName) = 0 Then baseName = "Export"

    exportFolder = basePath & "\" & baseName

    On Error GoTo MkDirFailed
    If Dir(exportFolder, vbDirectory) = "" Then
        MkDir exportFolder
    End If
    On Error GoTo 0

    EnsureExportFolder = exportFolder
    Exit Function

MkDirFailed:
    MsgBox "Couldn't create the export folder:" & vbCrLf & exportFolder & _
           vbCrLf & vbCrLf & Err.Description, vbExclamation, "TPD Add-in"
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
