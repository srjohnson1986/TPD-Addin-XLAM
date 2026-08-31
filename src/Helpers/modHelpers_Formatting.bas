Attribute VB_Name = "modHelpers_Formatting"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Sub InsertEQHeaderBlock(ws As Worksheet)
    
    InsertRows ws, 6

    ws.Range("A1").value = "EQUIPMENT LIST"
    ws.Range("A2").value = "Customer"
    ws.Range("A3").value = "Project"
    ws.Range("A4").value = "PM"
    ws.Range("A5").value = "Date"
    ws.Range("B5").value = GetTodaysDate
    ws.Range("A6").value = "Rev"

    With ws.Range("A1")
        .Font.name = "Aptos Narrow"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("A2:B6")
        .Font.name = "Aptos Narrow"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub InsertLogoAtRight(ws As Worksheet, imgPath As String, headerRow As Long, dataHeaderRow As Long)
    Dim pic As Shape
    Dim lastUsedCol As Long
    Dim origW As Single, origH As Single

    If Len(Dir(imgPath)) = 0 Then Exit Sub

    Set pic = ws.Shapes.AddPicture( _
        Filename:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=-1, Height:=-1)

    origW = pic.Width
    origH = pic.Height

    lastUsedCol = ws.Cells(dataHeaderRow, ws.Columns.Count).End(xlToLeft).Column

    pic.Left = ws.Cells(headerRow, lastUsedCol).Left + _
               (ws.Cells(headerRow, lastUsedCol).Width - origW)
    pic.Top = ws.Cells(headerRow, lastUsedCol).Top

    pic.Width = origW
    pic.Height = origH
    pic.LockAspectRatio = msoTrue
    pic.Placement = xlFreeFloating
End Sub

Public Sub FormatEQSheet(ws As Worksheet, dataHeaderRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim descCol As Variant

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells(dataHeaderRow, ws.Columns.Count).End(xlToLeft).Column

    Set dataRange = ws.Range(ws.Cells(dataHeaderRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With

    descCol = Application.Match("Description", ws.Rows(dataHeaderRow), 0)

    If Not IsError(descCol) Then
        If descCol > 1 Then
            ws.Range(ws.Cells(dataHeaderRow, 1), ws.Cells(lastRow, descCol - 1)).HorizontalAlignment = xlCenter
        End If
        If descCol < lastCol Then
            ws.Range(ws.Cells(dataHeaderRow, descCol + 1), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
        End If
    Else
        dataRange.HorizontalAlignment = xlCenter
    End If

    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    Dim r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = dataHeaderRow To lastRow
        If Trim(CStr(ws.Cells(r, 1).value)) <> "" Then
            ws.Rows(r).Interior.ColorIndex = xlNone
        End If
    Next r
    
    
End Sub

Public Sub FormatSplitSheet(ws As Worksheet, headerRow As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    With dataRange
        .Font.name = "Arial"
        .Font.Size = 10
    End With
    
    AutoFitUsedColumns (ws)

End Sub


Public Sub SafeFreezePanes(ws As Worksheet, headerRow As Long)
    On Error Resume Next
    ws.Activate
    ws.Range("A" & headerRow + 1).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Public Sub ResizeImageToMaxRows(pic As Shape, ws As Worksheet, headerRow As Long, maxRows As Long)
    Dim maxHeight As Double
    Dim scaleFactor As Double

    ' Calculate maximum allowed height based on the header row height
    maxHeight = ws.Rows(headerRow).Height * maxRows

    If pic.Height <= maxHeight Then Exit Sub

    scaleFactor = maxHeight / pic.Height

    pic.LockAspectRatio = msoTrue
    pic.Height = pic.Height * scaleFactor
End Sub

Public Sub InsertRows(ws As Worksheet, numberOfRows As Long)
    If numberOfRows <= 0 Then Exit Sub
    ws.Rows("1:" & numberOfRows).Insert Shift:=xlDown
End Sub


Public Function GetTodaysDate() As String
    GetTodaysDate = Format(Date, "mm-dd-yyyy")
End Function

Public Sub InsertDefaultCustScheduleHeader(ws As Worksheet)
    
    ' on the active worksheet, get the last column that is populated in row 1
    Dim lastColumn As Long
    lastColumn = GetLastCol(ws, 1)
    
    ' number of rows to insert
    Dim numberOfRows As Long
    numberOfRows = 5
    
    ' on the active worksheet, insert numberOfRows
    InsertRows ws, numberOfRows
    
    ' headerrow is one more than the number of inserted rows
    Dim headerRow As Long
    headerRow = numberOfRows + 1
    
    ' date text 2 rows above header
    Dim dateTextRow As Long
    dateTextRow = headerRow - 2
    
    ' Left date text cell (column before last)
    With ws.Cells(dateTextRow, lastColumn - 1)
        .value = "Date"
        .Font.name = "Arial"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
    
    ' Right date text cell (last column)
    With ws.Cells(dateTextRow, lastColumn)
        .value = GetTodaysDate()
        .Font.name = "Arial"
        .Font.Size = 11
        .Font.Bold = False
        .HorizontalAlignment = xlRight
    End With
    
    ' Insert logo on this worksheet on row 1, left-aligned, scaled to the number of rows we inserted
    InsertDefaultLogo ws, 1, "left", numberOfRows
     
End Sub

Public Sub AutoFitUsedColumns(ws As Worksheet)
    Dim usedCols As Range

    ' If the sheet has no used range, exit safely
    If ws.UsedRange Is Nothing Then Exit Sub

    ' Get only the used columns
    Set usedCols = ws.UsedRange.Columns

    ' Autofit only the populated columns
    usedCols.AutoFit
End Sub

Public Sub InsertDefaultCustEQHeader(ws As Worksheet)
    
    Dim headingsRow As Long
    ' Assume the column headings are on row 1 to start.
    headingsRow = 1
    
    ' On the active worksheet, get the last column that is populated in the row of headings
    Dim lastColumn As Long
    lastColumn = GetLastCol(ws, headingsRow)
    
    ' Formats the sheet for equipment lists using assuming that row 1 is the header.
    FormatEQSheet ws, headingsRow
    AutoFitUsedColumns ws
      
    ' Inserts 6 rows (currently) and adds the EQ header text
    InsertEQHeaderBlock ws
    
    ' number of rows inserted (modify as EQ header block changes)
    Dim numberOfRows As Long
    numberOfRows = 6
          
    ' headingsRow is after the number of inserted rows
    headingsRow = numberOfRows + headingsRow
          
    ' Insert logo on this worksheet on the headingsRow, right-top aligned, scaled to the number of rows we inserted
    InsertDefaultLogo ws, headingsRow, "right-top", numberOfRows
    
    SafeFreezePanes ws, headingsRow
     
End Sub

Public Sub InsertDefaultLogo(ws As Worksheet, headerRow As Long, _
                             Optional alignment As String = "left", _
                             Optional maxRows As Long = 0)

    Dim srcSheet As Worksheet
    Dim srcPic As Shape
    Dim newPic As Shape
    Dim origW As Single, origH As Single
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
        vert = "header"   ' default vertical alignment
    End If

    ' Get embedded logo
    Set srcSheet = ThisWorkbook.Worksheets("_Resources")
    Set srcPic = srcSheet.Shapes("DefaultLogo")

    ' Copy/paste into target sheet
    srcPic.Copy
    ws.Paste
    Set newPic = ws.Shapes(ws.Shapes.Count)

    ' Capture natural size
    origW = newPic.Width
    origH = newPic.Height

    ' Optional: resize to fit max rows
    If maxRows > 0 Then
        ResizeImageToMaxRows newPic, ws, headerRow, maxRows
    Else
        newPic.Width = origW
        newPic.Height = origH
    End If

    ' -----------------------------
    ' Horizontal alignment
    ' -----------------------------
    Select Case horiz
        Case "left"
            newPic.Left = ws.Cells(headerRow, 1).Left

        Case "right"
            lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
            Set targetCell = ws.Cells(headerRow, lastCol)
            newPic.Left = targetCell.Left + (targetCell.Width - newPic.Width)

        Case "center"
            lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
            Set targetCell = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
            newPic.Left = targetCell.Left + (targetCell.Width - newPic.Width) / 2

        Case Else
            newPic.Left = ws.Cells(headerRow, 1).Left
    End Select

    ' -----------------------------
    ' Vertical alignment
    ' -----------------------------
    Select Case vert
        Case "top"
            newPic.Top = ws.Cells(1, 1).Top   ' align to row 1

        Case "header"
            newPic.Top = ws.Cells(headerRow, 1).Top   ' default behavior

        Case Else
            newPic.Top = ws.Cells(headerRow, 1).Top
    End Select

    newPic.LockAspectRatio = msoTrue
    newPic.Placement = xlFreeFloating
End Sub
