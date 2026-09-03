Attribute VB_Name = "modHelpers_CheckboxLayout"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Builds the checkbox grid inside a column-picker frame: one
'  checkbox per heading, laid out in 3 columns, filling each
'  column top-to-bottom before starting the next.
'
'  rowsPerColumn is the preferred column height. If there are
'  more headings than fit in 3 columns of that height, the
'  columns grow downward instead of spilling into a 4th column
'  off the right edge, and the frame gets a vertical scrollbar
'  so every checkbox stays reachable (#47).
'===========================================================

Option Explicit

Private Const COLUMN_COUNT As Long = 3
Private Const ROW_PITCH As Single = 18
Private Const MARGIN As Single = 6

Public Sub LayoutCheckboxes( _
        ByVal fra As MSForms.Frame, _
        ByVal headingList As Variant, _
        ByVal rowsPerColumn As Long, _
        ByVal baseName As String, _
        Optional ByVal preferredDefaults As Variant = Empty, _
        Optional ByVal applyDefaults As Boolean = False)

    Dim i As Long
    Dim idx As Long
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim chk As MSForms.CheckBox
    Dim chkWidth As Single
    Dim heading As String
    Dim headingCount As Long
    Dim gridHeight As Single

    ' Clear existing checkboxes
    For idx = fra.Controls.Count - 1 To 0 Step -1
        If TypeName(fra.Controls(idx)) = "CheckBox" Then
            fra.Controls.Remove fra.Controls(idx).name
        End If
    Next idx

    headingCount = UBound(headingList) - LBound(headingList) + 1
    If headingCount <= 0 Then Exit Sub

    ' Keep it to COLUMN_COUNT columns: if the preferred height isn't enough,
    ' make the columns as tall as they need to be.
    If rowsPerColumn < 1 Then rowsPerColumn = 1
    If headingCount > rowsPerColumn * COLUMN_COUNT Then
        rowsPerColumn = -Int(-headingCount / COLUMN_COUNT)   ' ceil(count / COLUMN_COUNT)
    End If

    chkWidth = (fra.Width - 12) / COLUMN_COUNT

    colIndex = 0
    rowIndex = 0

    For i = LBound(headingList) To UBound(headingList)
        heading = CStr(headingList(i))

        Set chk = fra.Controls.Add("Forms.CheckBox.1", _
                    baseName & "_" & i & "_" & Replace(heading, " ", "_"))
        chk.Caption = heading
        chk.Left = MARGIN + (colIndex * chkWidth)
        chk.Top = MARGIN + (rowIndex * ROW_PITCH)
        chk.Width = chkWidth - MARGIN

        rowIndex = rowIndex + 1
        If rowIndex = rowsPerColumn Then
            rowIndex = 0
            colIndex = colIndex + 1
        End If
    Next i

    ' Scroll vertically when the grid is taller than the frame's interior.
    gridHeight = (MARGIN * 2) + (rowsPerColumn * ROW_PITCH)
    If gridHeight > fra.InsideHeight Then
        fra.ScrollBars = fmScrollBarsVertical
        fra.ScrollHeight = gridHeight
        fra.KeepScrollBarsVisible = fmScrollBarsVertical
    Else
        fra.ScrollBars = fmScrollBarsNone
    End If
End Sub
