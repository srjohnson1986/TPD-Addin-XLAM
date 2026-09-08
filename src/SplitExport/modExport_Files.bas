Attribute VB_Name = "modExport_Files"
'@Folder("TPD_Addin.SplitExport")

'===========================================================
'  The file end of "Save Each Sheet to XLSX": locate/create the
'  per-workbook export folder (EnsureExportFolder), write a
'  single worksheet out to its own .xlsx (ExportSheetToXLSX,
'  which returns "" on success or a reason string on failure),
'  and build the filename suffix the dialog previews and the
'  export uses (BuildExportSuffix / PreviewExportFileName).
'
'  Was modHelpers_Export - the only modHelpers_* outside the
'  Helpers folder, and it wasn't shared: everything here is
'  export-only. The one genuinely shared routine it did hold,
'  ReportBatchOutcome, moved to modHelpers_Reporting.
'===========================================================

Option Explicit

' The filename suffix "Save Each Sheet to XLSX" appends to every sheet name:
' the user's text, then " - " + today's date when the date box is ticked, then
' just the date when the text is blank. One place so the frmFilenameOptions
' live preview and the real export can't disagree (#118).
Public Function BuildExportSuffix(ByVal userAppend As String, _
                                  ByVal includeDate As Boolean) As String
    Dim s As String

    s = Trim$(userAppend)
    If includeDate Then
        If Len(s) > 0 Then s = s & " - " & GetTodaysDate() Else s = GetTodaysDate()
    End If

    BuildExportSuffix = s
End Function

' The file name one sheet would be exported as, for the frmFilenameOptions
' preview. Mirrors ExportSheetToXLSX: "<sheet> - <suffix>.xlsx", the whole
' thing run through SanitizeFileText.
Public Function PreviewExportFileName(ByVal sampleSheetName As String, _
                                     ByVal userAppend As String, _
                                     ByVal includeDate As Boolean) As String
    Dim base As String
    Dim suffix As String

    base = sampleSheetName
    If Len(Trim$(base)) = 0 Then base = "Sheet1"

    suffix = BuildExportSuffix(userAppend, includeDate)
    If Len(suffix) > 0 Then base = base & " - " & suffix

    PreviewExportFileName = SanitizeFileText(base) & ".xlsx"
End Function


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
' so it is never left open and visible. An already-existing target file is
' overwritten without prompting (re-exporting is a deliberate refresh).
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

    ' Without this, SaveAs onto an existing file prompts "replace?" - once per
    ' sheet on a re-export. Restored right after (and in Failed).
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=fullName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    newWb.Close SaveChanges:=False

    Exit Function

Failed:
    Dim reason As String
    reason = Err.Description

    Application.DisplayAlerts = True
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0

    ExportSheetToXLSX = "'" & ws.name & "': " & reason
End Function
