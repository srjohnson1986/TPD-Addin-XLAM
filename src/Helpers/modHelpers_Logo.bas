Attribute VB_Name = "modHelpers_Logo"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Places and sizes the logo shape on a generated sheet:
'  the embedded default logo from _Resources (InsertDefaultLogo,
'  with horizontal/vertical alignment keywords) and resizing a
'  picture to fit a maximum row count. anchorRow is the row the
'  logo is positioned against - a title-block row or a heading
'  row depending on the caller.
'  PastePicture pulls an image off the clipboard as a StdPicture
'  (used by frmSetTPDDefaults to preview the embedded logo).
'  InsertLogoAtRight / SafeInsertLogoAtRight are legacy of the
'  file-path logo approach - slated for removal with the
'  one-click EQ List refactor (#25).
'===========================================================

Option Explicit

Public Sub InsertLogoAtRight(ws As Worksheet, imgPath As String, headerRow As Long, dataHeaderRow As Long)
    Dim pic As Shape
    Dim lastUsedCol As Long
    Dim naturalWidth As Single, naturalHeight As Single

    If Len(Dir(imgPath)) = 0 Then Exit Sub

    Set pic = ws.Shapes.AddPicture( _
        Filename:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=-1, Height:=-1)

    naturalWidth = pic.Width
    naturalHeight = pic.Height

    lastUsedCol = ws.Cells(dataHeaderRow, ws.Columns.Count).End(xlToLeft).Column

    pic.Left = ws.Cells(headerRow, lastUsedCol).Left + _
               (ws.Cells(headerRow, lastUsedCol).Width - naturalWidth)
    pic.Top = ws.Cells(headerRow, lastUsedCol).Top

    pic.Width = naturalWidth
    pic.Height = naturalHeight
    pic.LockAspectRatio = msoTrue
    pic.Placement = xlFreeFloating
End Sub

Public Sub SafeInsertLogoAtRight(ws As Worksheet, imgPath As String, headerRow As Long, dataHeaderRow As Long)
    On Error GoTo LogoErr

    If Len(Dir(imgPath)) = 0 Then Exit Sub

    InsertLogoAtRight ws, imgPath, headerRow, dataHeaderRow
    Exit Sub

LogoErr:
    MsgBox "Unable to insert logo image: " & imgPath, vbExclamation
End Sub

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
            lastCol = ws.Cells(anchorRow, ws.Columns.Count).End(xlToLeft).Column
            Set targetCell = ws.Cells(anchorRow, lastCol)
            newPic.Left = targetCell.Left + (targetCell.Width - newPic.Width)

        Case "center"
            lastCol = ws.Cells(anchorRow, ws.Columns.Count).End(xlToLeft).Column
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
