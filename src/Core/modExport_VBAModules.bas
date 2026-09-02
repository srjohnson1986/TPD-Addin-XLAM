Attribute VB_Name = "modExport_VBAModules"
'@Folder("TPD_Addin.Core")

Sub ExportAllVBAModules()
    Dim vbProj As Object, vbComp As Object
    Dim exportPath As String, ext As String

    exportPath = "C:\VBA_Export\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    Set vbProj = ThisWorkbook.VBProject

    For Each vbComp In vbProj.VBComponents
        If vbComp.Type <> 100 Then  ' skip document objects (ThisWorkbook, sheets)
            Select Case vbComp.Type
                Case 1: ext = ".bas"
                Case 2: ext = ".cls"
                Case 3: ext = ".frm"
            End Select
            vbComp.Export exportPath & vbComp.name & ext
        End If
    Next vbComp

    MsgBox "Exported " & vbProj.VBComponents.Count & " components to " & exportPath
End Sub

'============================================================
' ExportAllVBAModules2 - replacement for ExportAllVBAModules
' above (kept alongside it for reference; safe to delete once
' you've verified this one works).
'
' Fixes vs. the original:
'   - No longer skips document modules (ThisWorkbook, Sheet1-3)
'   - Destination is a constant you set once, not a hardcoded
'     C:\ path unrelated to your repo
'   - One failing component logs an error and keeps going,
'     instead of aborting the whole export
'   - Groups output into subfolders read from each module's
'     '@Folder("TPD_Addin.X") annotation, matching the Code
'     Explorer organization
'   - Writes export_log.txt listing exactly what happened
'
' Requires: Trust Center > Macro Settings > "Trust access to
' the VBA project object model" must be enabled.
'============================================================

Public Sub ExportAllVBAModules2()

    ' Point this at your local clone's /src folder.
    Const DEST_ROOT As String = "C:\Users\srjoh\OneDrive\Documents\GitHub\TPD-Addin-XLAM\src\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(DEST_ROOT) Then
        MsgBox "Destination folder does not exist:" & vbNewLine & DEST_ROOT & vbNewLine & _
               "Update DEST_ROOT in this macro to point at your cloned repo's /src folder, then try again.", _
               vbExclamation, "Export cancelled"
        Exit Sub
    End If

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim logLines As New Collection
    Dim exportedCount As Long, skippedCount As Long
    Dim vbComp As Object
    Dim result As String

    For Each vbComp In vbProj.VBComponents
        result = ExportOneComponent(vbComp, DEST_ROOT, fso)
        logLines.Add result
        If Left(result, 2) = "OK" Then
            exportedCount = exportedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next vbComp

    WriteExportLog DEST_ROOT, logLines, exportedCount, skippedCount

    MsgBox "Export complete." & vbNewLine & _
           exportedCount & " component(s) exported" & vbNewLine & _
           skippedCount & " skipped/failed" & vbNewLine & vbNewLine & _
           "See export_log.txt in:" & vbNewLine & DEST_ROOT, vbInformation, "Export finished"

End Sub

Private Function ExportOneComponent(vbComp As Object, destRoot As String, fso As Object) As String
    On Error GoTo Failed

    Dim ext As String
    Select Case vbComp.Type
        Case 1: ext = ".bas"    ' Standard module
        Case 2: ext = ".cls"    ' Class module
        Case 3: ext = ".frm"    ' UserForm (.frx is written alongside automatically)
        Case 100: ext = ".cls"  ' Document module (ThisWorkbook, Sheet1, Sheet2, Sheet3)
        Case Else
            ExportOneComponent = "SKIPPED (unrecognized component type " & vbComp.Type & "): " & vbComp.name
            Exit Function
    End Select

    Dim folderTag As String
    folderTag = GetFolderTag(vbComp)

    Dim destFolder As String
    destFolder = destRoot
    If Len(folderTag) > 0 Then
        destFolder = destRoot & Replace(folderTag, "TPD_Addin.", "") & "\"
    End If
    If Not fso.FolderExists(destFolder) Then fso.CreateFolder destFolder

    Dim destPath As String
    destPath = destFolder & vbComp.name & ext
    vbComp.Export destPath

    ExportOneComponent = "OK: " & vbComp.name & ext & "  ->  " & destPath
    Exit Function

Failed:
    ExportOneComponent = "FAILED: " & vbComp.name & "  (" & Err.Description & ")"
End Function

Private Function GetFolderTag(vbComp As Object) As String
    ' Reads a '@Folder("TPD_Addin.X") annotation from a module's
    ' first few lines, if present, so export can mirror the
    ' Code Explorer grouping on disk.
    Dim cm As Object
    Dim i As Long, lineText As String
    Dim startPos As Long, endPos As Long
    Dim linesToScan As Long

    Set cm = vbComp.CodeModule
    linesToScan = cm.CountOfLines
    If linesToScan > 10 Then linesToScan = 10

    For i = 1 To linesToScan
        lineText = cm.lines(i, 1)
        If InStr(1, lineText, "@Folder(", vbTextCompare) > 0 Then
            startPos = InStr(lineText, """") + 1
            endPos = InStr(startPos, lineText, """")
            If startPos > 1 And endPos > startPos Then
                GetFolderTag = Mid(lineText, startPos, endPos - startPos)
            End If
            Exit Function
        End If
    Next i
End Function

Private Sub WriteExportLog(destRoot As String, logLines As Collection, exportedCount As Long, skippedCount As Long)
    Dim fnum As Integer
    Dim entry As Variant

    fnum = FreeFile
    Open destRoot & "export_log.txt" For Output As #fnum
        Print #fnum, "TPD_Addin export - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #fnum, exportedCount & " exported, " & skippedCount & " skipped/failed"
        Print #fnum, String(60, "-")
        For Each entry In logLines
            Print #fnum, entry
        Next entry
    Close #fnum
End Sub
