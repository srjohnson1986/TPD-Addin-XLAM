Attribute VB_Name = "modHelpers_Logo"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Places and sizes the header logo on a generated sheet.
'  InsertDefaultLogo takes horizontal/vertical alignment
'  keywords and an optional max row count to scale into;
'  anchorRow is the row the logo is positioned against - a
'  title-block row or a heading row depending on the caller.
'  ResizeImageToMaxRows is the scaling helper.
'
'  The logo is the user's chosen image when one has been set
'  in Set TPD Defaults (EffectiveLogoPath -> a file under
'  %APPDATA%\TPD_Addin\, [#95](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/95)),
'  otherwise the built-in logo. The built-in lives as the
'  _Resources / DefaultLogo shape in the base .xlam
'  (DefaultLogoShape); BuiltInLogoFile extracts it once to a
'  PNG next to the user copy so it too can be previewed and
'  AddPicture'd. SetUserLogo / ClearUserLogo manage the user
'  copy; LogoPreviewPicture backs the dialog's preview ([#109]).
'===========================================================

Option Explicit

' The hidden sheet + shape that hold the add-in's embedded default logo,
' shipped in the base .xlam. Everything resolves them through
' DefaultLogoShape so the names live in exactly one place.
Private Const RESOURCE_SHEET_NAME As String = "_Resources"
Private Const LOGO_SHAPE_NAME As String = "DefaultLogo"

' Where a user-chosen logo is copied to, and the name it's stored under
' (the extension varies). Per Windows user, alongside the registry prefs.
Private Const USER_LOGO_BASENAME As String = "user_logo"

' A file copy of the embedded DefaultLogo shape, kept in the same folder so
' the dialog can preview the built-in logo and callers can AddPicture it
' (see BuiltInLogoFile).
Private Const BUILTIN_LOGO_FILE As String = "builtin_logo.png"

' The embedded default-logo shape, or Nothing when the _Resources sheet
' or the shape is missing (a base-file build mistake - see CONTRIBUTING.md).
Public Function DefaultLogoShape() As Shape
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET_NAME)
    If Not ws Is Nothing Then Set DefaultLogoShape = ws.Shapes(LOGO_SHAPE_NAME)
    On Error GoTo 0
End Function


'===========================================================
'  User-chosen logo (Set TPD Defaults - Logo tab)
'===========================================================

' %APPDATA%\TPD_Addin - created on demand. Same per-user home the registry
' prefs live under. Returns "" only if APPDATA is somehow unset.
Public Function UserLogoDir() As String
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then Exit Function

    base = base & "\TPD_Addin"
    If Len(Dir$(base, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir base
        On Error GoTo 0
    End If

    UserLogoDir = base
End Function

' A PNG copy of the embedded DefaultLogo shape in %APPDATA%\TPD_Addin\, so
' the dialog can preview the built-in logo and InsertDefaultLogo can
' AddPicture it. The only way to get a Shape to a file is a throwaway chart,
' so this is extracted once and refreshed whenever the add-in is newer than
' the copy. Returns "" when the extraction isn't possible - a non-interactive
' Excel (build / automation, where Chart.Export hangs), or an export failure
' - and the callers fall back to copying the shape / a blank preview.
Public Function BuiltInLogoFile() As String
    Dim folder As String
    Dim dest As String

    folder = UserLogoDir()
    If Len(folder) = 0 Then Exit Function
    dest = folder & "\" & BUILTIN_LOGO_FILE

    On Error Resume Next
    If Len(Dir$(dest)) > 0 Then
        If FileDateTime(dest) >= FileDateTime(ThisWorkbook.FullName) Then
            BuiltInLogoFile = dest
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' Chart.Export needs a real display context - never attempt it headless.
    If Not Application.Visible Then Exit Function

    If ExportDefaultLogoShape(dest) Then BuiltInLogoFile = dest
End Function

' Pastes the DefaultLogo shape onto a throwaway chart and exports it to dest
' as PNG. Best-effort - the chart is always removed. Returns True only if a
' file landed.
Private Function ExportDefaultLogoShape(ByVal dest As String) As Boolean
    Dim src As Shape
    Dim host As Worksheet
    Dim co As ChartObject

    Set src = DefaultLogoShape()
    If src Is Nothing Then Exit Function
    Set host = src.Parent

    On Error GoTo CleanUp
    On Error Resume Next: Kill dest: On Error GoTo CleanUp

    Set co = host.ChartObjects.Add(0, 0, src.Width + 4, src.Height + 4)
    src.Copy
    co.Chart.Paste
    co.Chart.Export dest, "PNG"

CleanUp:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    On Error GoTo 0
    ExportDefaultLogoShape = (Len(Dir$(dest)) > 0)
End Function

' The user's chosen logo file, or "" when none is set (or the saved copy
' has gone missing - in which case the embedded shape takes over).
Public Function EffectiveLogoPath() As String
    Dim p As String
    p = LoadPref(PREF_LOGO_PATH, "")
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then EffectiveLogoPath = p
    End If
End Function

' Copies sourcePath into %APPDATA%\TPD_Addin\user_logo.<ext> and records it
' as PREF_LOGO_PATH. Returns "" on success, or a reason string the dialog
' shows and keeps its form open on. Validates that GDI+ can actually read
' the image before committing, so a broken file can't become the logo.
Public Function SetUserLogo(ByVal sourcePath As String) As String
    Dim ext As String
    Dim folder As String
    Dim dest As String

    If Len(Dir$(sourcePath)) = 0 Then
        SetUserLogo = "That file could not be found."
        Exit Function
    End If

    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".")))
    If InStr(1, ".png.jpg.jpeg.gif.bmp.", ext & ".", vbTextCompare) = 0 Then
        SetUserLogo = "Choose a PNG, JPG, GIF or BMP image."
        Exit Function
    End If

    If modGdiPlus.LoadPictureGDIP(sourcePath) Is Nothing Then
        SetUserLogo = "That image couldn't be read. Try a different file."
        Exit Function
    End If

    folder = UserLogoDir()
    If Len(folder) = 0 Then
        SetUserLogo = "Couldn't find a place to store the logo (APPDATA is unset)."
        Exit Function
    End If

    ClearUserLogo                         ' drop any previous copy (any ext)
    dest = folder & "\" & USER_LOGO_BASENAME & ext

    On Error Resume Next
    FileCopy sourcePath, dest
    If Err.Number <> 0 Then
        SetUserLogo = "Couldn't copy the image: " & Err.Description
        Exit Function
    End If
    On Error GoTo 0

    SavePref PREF_LOGO_PATH, dest
End Function

' Removes the user's logo copy and clears the preference, so the embedded
' shape wins again. Safe to call when nothing is set. All best-effort.
Public Sub ClearUserLogo()
    Dim savedPath As String
    Dim folder As String
    Dim strays() As String
    Dim count As Long
    Dim i As Long
    Dim f As String

    On Error Resume Next

    savedPath = LoadPref(PREF_LOGO_PATH, "")
    If Len(savedPath) > 0 Then Kill savedPath

    ' Any stray user_logo.* (e.g. the pref was lost but a copy remains).
    ' Collect every name first - Kill mid-Dir-enumeration invalidates the
    ' iterator.
    folder = UserLogoDir()
    If Len(folder) > 0 Then
        ReDim strays(0 To 15)
        f = Dir$(folder & "\" & USER_LOGO_BASENAME & ".*")
        Do While Len(f) > 0 And count <= UBound(strays)
            strays(count) = folder & "\" & f
            count = count + 1
            f = Dir$
        Loop
        For i = 0 To count - 1
            Kill strays(i)
        Next i
    End If

    DeletePref PREF_LOGO_PATH

    On Error GoTo 0
End Sub

' Picture for the Set TPD Defaults logo preview, loaded via GDI+: the staged
' file the user just picked (pendingPath), else the saved user logo, else the
' built-in logo (a file copy - BuiltInLogoFile). Returns Nothing only when
' even the built-in can't be rendered (headless, or the extraction failed) -
' the dialog then just clears the preview.
Public Function LogoPreviewPicture(ByVal pendingPath As String, _
                                   ByVal pendingRestore As Boolean) As stdole.IPictureDisp
    Dim imgPath As String

    If pendingRestore Then
        imgPath = BuiltInLogoFile()
    ElseIf Len(pendingPath) > 0 Then
        imgPath = pendingPath
    Else
        imgPath = EffectiveLogoPath()
        If Len(imgPath) = 0 Then imgPath = BuiltInLogoFile()
    End If

    If Len(imgPath) = 0 Then Exit Function

    Set LogoPreviewPicture = modGdiPlus.LoadPictureGDIP(imgPath)
End Function

Public Sub ResizeImageToMaxRows(pic As Shape, ws As Worksheet, anchorRow As Long, maxRows As Long)
    Dim maxHeight As Double
    Dim scaleFactor As Double

    ' Calculate maximum allowed height based on the header row height
    maxHeight = ws.Rows(anchorRow).Height * maxRows

    If pic.Height <= maxHeight Then Exit Sub

    scaleFactor = maxHeight / pic.Height

    pic.LockAspectRatio = msoTrue
    pic.Height = pic.Height * scaleFactor
End Sub

Public Sub InsertDefaultLogo(ws As Worksheet, anchorRow As Long, _
                             Optional alignment As String = "left", _
                             Optional maxRows As Long = 0)

    Dim srcPic As Shape
    Dim newPic As Shape
    Dim userPath As String
    Dim naturalWidth As Single, naturalHeight As Single
    Dim targetCell As Range
    Dim lastCol As Long
    Dim horiz As String, vert As String

    ' Split alignment into horizontal + vertical parts
    alignment = LCase(alignment)
    If InStr(alignment, "-") > 0 Then
        horiz = Split(alignment, "-")(0)
        vert = Split(alignment, "-")(1)
    Else
        horiz = alignment
        vert = "anchor"   ' default vertical alignment
    End If

    ' The logo to stamp on ws, in preference order: the user's chosen image,
    ' the extracted PNG copy of the built-in logo, then a straight copy of the
    ' embedded shape. Either way newPic is a picture shape on ws for the
    ' sizing/alignment below; fail loud only if none are available (the caller
    ' is building a sheet around it).
    userPath = EffectiveLogoPath()
    If Len(userPath) = 0 Then userPath = BuiltInLogoFile()

    If Len(userPath) > 0 Then
        Set newPic = ws.Shapes.AddPicture(userPath, msoFalse, msoTrue, 0, 0, -1, -1)
    Else
        Set srcPic = DefaultLogoShape()
        If srcPic Is Nothing Then
            Err.Raise 5, "InsertDefaultLogo", _
                "The add-in's _Resources sheet or its DefaultLogo shape is missing."
        End If
        srcPic.Copy
        ws.Paste
        Set newPic = ws.Shapes(ws.Shapes.Count)
    End If

    ' Capture natural size
    naturalWidth = newPic.Width
    naturalHeight = newPic.Height

    ' Optional: resize to fit max rows
    If maxRows > 0 Then
        ResizeImageToMaxRows newPic, ws, anchorRow, maxRows
    Else
        newPic.Width = naturalWidth
        newPic.Height = naturalHeight
    End If

    ' -----------------------------
    ' Horizontal alignment
    ' -----------------------------
    Select Case horiz
        Case "left"
            newPic.Left = ws.Cells(anchorRow, 1).Left

        Case "right"
            lastCol = GetLastCol(ws, anchorRow)
            Set targetCell = ws.Cells(anchorRow, lastCol)
            newPic.Left = targetCell.Left + (targetCell.Width - newPic.Width)

        Case "center"
            lastCol = GetLastCol(ws, anchorRow)
            Set targetCell = ws.Range(ws.Cells(anchorRow, 1), ws.Cells(anchorRow, lastCol))
            newPic.Left = targetCell.Left + (targetCell.Width - newPic.Width) / 2

        Case Else
            newPic.Left = ws.Cells(anchorRow, 1).Left
    End Select

    ' -----------------------------
    ' Vertical alignment
    ' -----------------------------
    Select Case vert
        Case "top"
            newPic.Top = ws.Cells(1, 1).Top   ' align to row 1

        Case "anchor"
            newPic.Top = ws.Cells(anchorRow, 1).Top   ' default behavior

        Case Else
            newPic.Top = ws.Cells(anchorRow, 1).Top
    End Select

    newPic.LockAspectRatio = msoTrue
    newPic.Placement = xlFreeFloating
End Sub
