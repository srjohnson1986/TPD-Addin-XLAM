Attribute VB_Name = "modExport_VBAModules"
'@Folder("TPD_Addin.Core")

'============================================================
' ExportAllVBAModules - exports every VBA component to /src.
'
'   - Includes document modules (ThisWorkbook, Sheet1-3)
'   - Destination is found automatically: it walks up from the folder
'     this workbook is saved in until it reaches a repo root (a folder
'     holding both /src and /customUI), and falls back to a folder
'     picker if it can't (e.g. the add-in is installed elsewhere)
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

Option Explicit

Public Sub ExportAllVBAModules()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim srcRoot As String
    srcRoot = ResolveSrcRoot(fso)

    If Len(srcRoot) = 0 Then
        MsgBox "No destination folder was chosen, so nothing was exported.", _
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
        result = ExportOneComponent(vbComp, srcRoot, fso)
        logLines.Add result
        If Left(result, 2) = "OK" Then
            exportedCount = exportedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next vbComp

    WriteExportLog srcRoot, logLines, exportedCount, skippedCount

    MsgBox "Export complete." & vbNewLine & _
           exportedCount & " component(s) exported" & vbNewLine & _
           skippedCount & " skipped/failed" & vbNewLine & vbNewLine & _
           "See export_log.txt in:" & vbNewLine & srcRoot, vbInformation, "Export finished"

End Sub

Private Function ResolveSrcRoot(fso As Object) As String
    ' Finds the clone's /src folder without any hardcoded path.
    '   1. Walk up from this workbook's folder (a build in <repo>\build\ or
    '      a dev copy anywhere under the clone) to the first folder that holds
    '      both "src" and "customUI" - that is the repo root.
    '   2. Otherwise ask the user to pick the /src folder.
    ' Returns "" if neither works. The result always ends in a backslash.
    Const MAX_PARENT_LEVELS As Long = 6

    Dim probe As String
    Dim level As Long

    probe = ThisWorkbook.Path
    For level = 1 To MAX_PARENT_LEVELS
        If Len(probe) = 0 Then Exit For
        If fso.FolderExists(fso.BuildPath(probe, "src")) And _
           fso.FolderExists(fso.BuildPath(probe, "customUI")) Then
            ResolveSrcRoot = fso.BuildPath(probe, "src") & "\"
            Exit Function
        End If
        probe = fso.GetParentFolderName(probe)
    Next level

    Dim picked As String
    With Application.FileDialog(4)          ' msoFileDialogFolderPicker
        .Title = "Select the src folder of your TPD-Addin-XLAM clone"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function   ' cancelled
        picked = .SelectedItems(1)
    End With

    If Right$(picked, 1) <> "\" Then picked = picked & "\"
    If fso.FolderExists(picked) Then ResolveSrcRoot = picked
End Function

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
