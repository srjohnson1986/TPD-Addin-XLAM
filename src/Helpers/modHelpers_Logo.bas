Attribute VB_Name = "modHelpers_Logo"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Places and sizes the embedded default logo (from the
'  _Resources sheet) on a generated sheet: InsertDefaultLogo
'  takes horizontal/vertical alignment keywords and an optional
'  max row count to scale into. anchorRow is the row the logo
'  is positioned against - a title-block row or a heading row
'  depending on the caller. ResizeImageToMaxRows is the scaling
'  helper. DefaultLogoShape is the one place the _Resources /
'  DefaultLogo names are resolved.
'===========================================================

Option Explicit

' The hidden sheet + shape that hold the add-in's embedded default logo,
' shipped in the base .xlam. Everything resolves them through
' DefaultLogoShape so the names live in exactly one place.
Private Const RESOURCE_SHEET_NAME As String = "_Resources"
Private Const LOGO_SHAPE_NAME As String = "DefaultLogo"

' The embedded default-logo shape, or Nothing when the _Resources sheet
' or the shape is missing (a base-file build mistake - see CONTRIBUTING.md).
Public Function DefaultLogoShape() As Shape
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET_NAME)
    If Not ws Is Nothing Then Set DefaultLogoShape = ws.Shapes(LOGO_SHAPE_NAME)
    On Error GoTo 0
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

    ' Get embedded logo - fail loud, the caller is building a sheet around it
    Set srcPic = DefaultLogoShape()
    If srcPic Is Nothing Then
        Err.Raise 5, "InsertDefaultLogo", _
            "The add-in's _Resources sheet or its DefaultLogo shape is missing."
    End If

    ' Copy/paste into target sheet
    srcPic.Copy
    ws.Paste
    Set newPic = ws.Shapes(ws.Shapes.Count)

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
