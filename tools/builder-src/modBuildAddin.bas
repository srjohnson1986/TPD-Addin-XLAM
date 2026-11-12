Attribute VB_Name = "modBuildAddin"
'@Folder("TPD_Addin.Core")

Option Explicit

'===========================================================
' modBuildAddin
'
' Rebuilds TPD_Addin.xlam from the /src source tree.
'
' WHERE THIS RUNS:
'   In a SEPARATE builder workbook, tools\TPD_Builder.xlsm - NOT
'   in TPD_Addin.xlam itself. A running macro cannot safely
'   strip/reimport the module it is currently executing from,
'   and document modules (ThisWorkbook, Sheet1-3) can never be
'   removed and reimported at all - only code-injected in place.
'
'   This file, tools\builder-src\modBuildAddin.bas, is the
'   readable copy of that workbook's code. Keep the two the
'   same: tools\Import-BuilderSource.ps1 puts this text into the
'   workbook, tools\Export-BuilderSource.ps1 goes the other way,
'   and tools\Test-All.ps1 fails if they differ.
'
' REQUIREMENTS:
'   - File > Options > Trust Center > Macro Settings >
'     "Trust access to the VBA project object model" (checked)
'   - A base file at build\_base\TPD_Addin_base.xlam: a .xlam
'     with NO standard modules or forms, but with
'     Sheet1/Sheet2/Sheet3 already present (names must match
'     /src/Document exactly) and _Resources already containing
'     the embedded "DefaultLogo" shape.
'     tools\New-BaseFile.ps1 makes one from a released .xlam.
'   - The ribbon is NOT managed by this macro (VBA has no
'     object model for it). tools\Build-TPDAddin.ps1 writes
'     customUI\customUI14.xml into the built file after this
'     macro has finished. Run BuildAddin straight from the VBE
'     and the build keeps whatever ribbon the base file has.
'
' HOW TO RUN:
'   powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1
'   which calls BuildAddin(silent:=True, repoRoot:=<your clone>).
'   From the VBE, press F5 in BuildAddin: with no repoRoot it
'   uses the folder above the one this workbook is saved in.
'===========================================================

' Where everything is, worked out by SetPaths at the start of each build.
' Nothing here is tied to one machine.
Private mSrcRoot As String
Private mBasePath As String
Private mOutputPath As String
Private mLogPath As String

' Pass silent:=True when calling this via Application.Run from an
' external automation script (e.g. PowerShell/COM). In that mode there
' is no user to click a MsgBox - a blocking MsgBox would hang the
' headless Excel process forever - so results go to the log instead.
'
' repoRoot is the folder holding src\, build\ and tools\. Leave it
' empty to use the folder above this workbook's own.
Public Sub BuildAddin(Optional ByVal silent As Boolean = False, _
                      Optional ByVal repoRoot As String = vbNullString)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not SetPaths(repoRoot, silent) Then Exit Sub

    ' ---- Sanity checks ----
    If Not fso.FileExists(mBasePath) Then
        ReportResult "BUILD CANCELLED: Base file not found: " & mBasePath, silent, isError:=True
        Exit Sub
    End If
    If Not fso.FolderExists(mSrcRoot) Then
        ReportResult "BUILD CANCELLED: Source folder not found: " & mSrcRoot, silent, isError:=True
        Exit Sub
    End If

    ' ---- Start from a fresh copy of the base file ----
    If fso.FileExists(mOutputPath) Then
        On Error Resume Next
        fso.DeleteFile mOutputPath, True
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            ReportResult "BUILD CANCELLED: Can't overwrite " & mOutputPath & _
                         " - it is most likely still loaded as an Excel add-in, or open " & _
                         "in another Excel window. Unload it (File > Options > Add-ins), " & _
                         "close any stray EXCEL.EXE, then run BuildAddin again.", _
                         silent, isError:=True
            Exit Sub
        End If
        On Error GoTo 0
    End If
    fso.CopyFile mBasePath, mOutputPath, True


    Dim targetWb As Workbook
    Set targetWb = Workbooks.Open(mOutputPath)

    Dim vbProj As Object
    Set vbProj = targetWb.VBProject

    Dim importLog As New Collection
    Dim skipLog As New Collection

    ' ---- Import every component from /src ----
    ImportComponentsFromFolder vbProj, mSrcRoot, fso, importLog, skipLog

    ' ---- Save and close ----
    targetWb.Save
    targetWb.Close SaveChanges:=False

    Dim msg As String
    msg = "BUILD COMPLETE: " & mOutputPath & " | " & _
          importLog.Count & " component(s) imported/injected, " & _
          skipLog.Count & " skipped"

    If skipLog.Count > 0 Then
        Dim s As Variant
        msg = msg & " | Skipped:"
        For Each s In skipLog
            msg = msg & " [" & s & "]"
        Next s
    End If

    ReportResult msg, silent, isError:=(skipLog.Count > 0)

End Sub

' Works out where src\, the base file, the output and the log are, from the
' repository folder. Returns False, having said why, when it cannot.
Private Function SetPaths(ByVal repoRoot As String, ByVal silent As Boolean) As Boolean
    Dim root As String

    root = Trim$(repoRoot)
    If Len(root) = 0 Then
        ' This workbook is saved in <repo>\tools.
        root = ThisWorkbook.Path
        If InStrRev(root, "\") > 0 Then root = Left$(root, InStrRev(root, "\") - 1)
    End If
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)

    ' A workbook in a OneDrive-synced folder can report a web address as its
    ' own path, which is no use for finding files on disk.
    If Len(root) = 0 Or LCase$(Left$(root, 4)) = "http" Then
        Debug.Print "BUILD CANCELLED: cannot tell where the repository is (got '" & root & "')."
        If Not silent Then
            MsgBox "Can't tell where the repository folder is from this workbook's path:" & vbCrLf & _
                   root & vbCrLf & vbCrLf & _
                   "Run tools\Build-TPDAddin.ps1 instead - it passes the folder in.", _
                   vbCritical, "Build cancelled"
        End If
        Exit Function
    End If

    mSrcRoot = root & "\src\"
    mBasePath = root & "\build\_base\TPD_Addin_base.xlam"
    mOutputPath = root & "\build\TPD_Addin.xlam"
    mLogPath = root & "\build\build.log"
    SetPaths = True
End Function

' Routes a result message either to an interactive MsgBox (normal VBE
' use) or appends it to mLogPath with a timestamp (headless/automated
' use). Also mirrors to the Immediate window either way for convenience.
Private Sub ReportResult(msg As String, silent As Boolean, isError As Boolean)
    Debug.Print msg

    If silent Then
        Dim fnum As Integer
        fnum = FreeFile
        Open mLogPath For Append As #fnum
            Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & msg
        Close #fnum
    Else
        MsgBox msg, IIf(isError, vbCritical, vbInformation), IIf(isError, "Build cancelled", "Build finished")
    End If
End Sub

Private Sub ImportComponentsFromFolder(vbProj As Object, folderPath As String, fso As Object, _
                                        importLog As Collection, skipLog As Collection)

    Dim subFolder As Object
    Dim fileItem As Object
    Dim ext As String
    Dim baseName As String
    Dim folder As Object

    Set folder = fso.GetFolder(folderPath)

    For Each fileItem In folder.Files
        ext = LCase(fso.GetExtensionName(fileItem.Path))
        baseName = fso.GetBaseName(fileItem.Path)

        Select Case ext
            Case "bas", "frm"
                On Error Resume Next
                vbProj.VBComponents.Import fileItem.Path
                If Err.Number <> 0 Then
                    skipLog.Add baseName & " (" & Err.Description & ")"
                    Err.Clear
                Else
                    importLog.Add baseName
                End If
                On Error GoTo 0

            Case "cls"
                If IsDocumentModule(baseName) Then
                    On Error Resume Next
                    InjectDocumentModuleCode vbProj, baseName, fileItem.Path
                    If Err.Number <> 0 Then
                        skipLog.Add baseName & " (" & Err.Description & ")"
                        Err.Clear
                    Else
                        importLog.Add baseName & " (injected)"
                    End If
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    vbProj.VBComponents.Import fileItem.Path
                    If Err.Number <> 0 Then
                        skipLog.Add baseName & " (" & Err.Description & ")"
                        Err.Clear
                    Else
                        importLog.Add baseName
                    End If
                    On Error GoTo 0
                End If

            Case "frx", "txt"
                ' .frx is pulled in automatically alongside its .frm; skip explicitly.
                ' export_log.txt and similar are not source components.

            Case Else
                ' Unrecognized file type in /src - ignore.
        End Select
    Next fileItem

    ' Recurse into subfolders (Helpers, EQList, SplitExport, etc.)
    For Each subFolder In folder.SubFolders
        ImportComponentsFromFolder vbProj, subFolder.Path & "\", fso, importLog, skipLog
    Next subFolder

End Sub

Private Function IsDocumentModule(componentName As String) As Boolean
    Select Case componentName
        Case "ThisWorkbook", "Sheet1", "Sheet2", "Sheet3"
            IsDocumentModule = True
        Case Else
            IsDocumentModule = False
    End Select
End Function

' Document modules (ThisWorkbook, Sheet1-3) already exist in the base
' file and cannot be removed/reimported. Instead we clear the existing
' code and inject the body of the exported .cls, skipping the
' VERSION/BEGIN...END/Attribute header block that the component
' already carries as metadata.
Private Sub InjectDocumentModuleCode(vbProj As Object, componentName As String, filePath As String)

    Dim comp As Object
    Dim cm As Object
    Dim fnum As Integer
    Dim raw As String
    Dim bodyCode As String

    On Error Resume Next
    Set comp = vbProj.VBComponents(componentName)
    On Error GoTo 0

    If comp Is Nothing Then
        Err.Raise vbObjectError + 1, , "No matching document module '" & componentName & "' in base file"
    End If

    Set cm = comp.CodeModule

    If cm.CountOfLines > 0 Then
        cm.DeleteLines 1, cm.CountOfLines
    End If

    fnum = FreeFile
    Open filePath For Input As #fnum
        raw = Input$(LOF(fnum), fnum)
    Close #fnum

    bodyCode = StripClassHeader(raw)

    If Len(Trim(bodyCode)) > 0 Then
        cm.AddFromString bodyCode
    End If

End Sub

' Strips the VERSION/BEGIN...END block and the run of "Attribute VB_..."
' lines from an exported .cls file, returning just the code body
' (declarations, procedures, and any trailing comments like @Folder).
Private Function StripClassHeader(raw As String) As String
    Dim normalized As String
    Dim lines() As String
    Dim i As Long
    Dim pastHeader As Boolean
    Dim inBeginEnd As Boolean
    Dim result As String
    Dim trimmed As String

    normalized = Replace(raw, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)

    pastHeader = False
    inBeginEnd = False

    For i = LBound(lines) To UBound(lines)
        trimmed = Trim(lines(i))

        If Not pastHeader Then
            If inBeginEnd Then
                ' Inside the BEGIN...END object-description block - skip
                ' every line unconditionally (property lines vary in
                ' format, so don't try to pattern-match them) until the
                ' matching END is found.
                If StrComp(trimmed, "END", vbTextCompare) = 0 Then
                    inBeginEnd = False
                End If
            ElseIf StrComp(trimmed, "BEGIN", vbTextCompare) = 0 _
                   Or Left(trimmed, 6) = "Begin " Then
                ' e.g. "BEGIN" (class) or "Begin {GUID} frmName" (form) - skip
                inBeginEnd = True
            ElseIf Left(UCase(trimmed), 7) = "VERSION" Then
                ' skip
            ElseIf Left(trimmed, 10) = "Attribute " Then
                ' skip
            ElseIf trimmed = "" Then
                ' skip blank lines while still in the header region
            Else
                pastHeader = True
            End If
        End If

        If pastHeader Then
            result = result & lines(i) & vbNewLine
        End If
    Next i

    StripClassHeader = result
End Function
