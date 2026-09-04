Attribute VB_Name = "modHelpers_Logo"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Places and sizes the embedded default logo (from the
'  _Resources sheet) on a generated sheet: InsertDefaultLogo
'  takes horizontal/vertical alignment keywords and an optional
'  max row count to scale into. anchorRow is the row the logo
'  is positioned against - a title-block row or a heading row
'  depending on the caller. ResizeImageToMaxRows is the scaling
'  helper. PastePicture pulls an image off the clipboard as a
'  StdPicture (used by frmSetTPDDefaults to preview the logo).
'===========================================================

Option Explicit

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

    Dim srcSheet As Worksheet
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

    ' Get embedded logo
    Set srcSheet = ThisWorkbook.Worksheets("_Resources")
    Set srcPic = srcSheet.Shapes("DefaultLogo")

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

Public Function PastePicture() As StdPicture
    ' Returns a picture object from the clipboard
    Dim IData As Object
    Set IData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    IData.GetData 1
    Set PastePicture = IData.GetData(1)
End Function
