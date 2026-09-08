Attribute VB_Name = "modHelpers_CheckboxLayout"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Builds the checkbox grid inside a column-picker frame: one
'  checkbox per heading, laid out in AT MOST 3 columns, filling
'  each column top-to-bottom before starting the next
'  (column-major - the reading order stays down-then-across).
'
'  rowsPerColumn is the preferred column height. If there are
'  more headings than fit in 3 columns of that height, the
'  columns grow taller rather than spilling into a 4th column
'  off the right edge (#47).
'
'  The frame keeps a vertical scrollbar ALWAYS (#150): a short
'  heading set shows a present-but-idle bar, a long one scrolls.
'  Never a horizontal bar - ScrollBars is fmScrollBarsVertical,
'  and the column width leaves room for the bar so the third
'  column isn't clipped under it. The frame is a fixed viewport;
'  this never resizes it or the form.
'===========================================================

Option Explicit

Private Const COLUMN_COUNT As Long = 3
Private Const ROW_PITCH As Single = 18
Private Const MARGIN As Single = 6

' Width the always-on vertical scrollbar takes out of the frame interior.
' Generous by a point or two - the cost of over-estimating is a little right
' padding on the third column; under-estimating clips it under the bar.
Private Const SCROLLBAR_ALLOWANCE As Single = 16

Public Sub LayoutCheckboxes( _
        ByVal fra As MSForms.Frame, _
        ByVal headingList As Variant, _
        ByVal rowsPerColumn As Long, _
        ByVal baseName As String)

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
    If headingCount <= 0 Then
        NormalizeFrameScroll fra, MARGIN * 2
        Exit Sub
    End If

    ' Keep it to COLUMN_COUNT columns: if the preferred height isn't enough,
    ' make the columns as tall as they need to be (scroll to reach them).
    If rowsPerColumn < 1 Then rowsPerColumn = 1
    If headingCount > rowsPerColumn * COLUMN_COUNT Then
        rowsPerColumn = -Int(-headingCount / COLUMN_COUNT)   ' ceil(count / COLUMN_COUNT)
    End If

    ' Leave room for the always-on scrollbar so the third column clears it.
    chkWidth = (fra.InsideWidth - SCROLLBAR_ALLOWANCE - MARGIN) / COLUMN_COUNT

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

    gridHeight = (MARGIN * 2) + (rowsPerColumn * ROW_PITCH)
    NormalizeFrameScroll fra, gridHeight
End Sub

' Always-on vertical scrollbar, never horizontal (#150). ScrollHeight is the
' laid-out content, or the interior height when the grid is shorter - so a
' short list still shows the bar (idle, thumb filling the track) and the grid
' never scrolls sideways. Wrapped defensively: a scrollbar quirk must never
' break the picker.
Private Sub NormalizeFrameScroll(ByVal fra As MSForms.Frame, ByVal contentHeight As Single)
    On Error Resume Next
    fra.ScrollBars = fmScrollBarsVertical
    fra.KeepScrollBarsVisible = fmScrollBarsVertical
    If contentHeight > fra.InsideHeight Then
        fra.ScrollHeight = contentHeight
    Else
        fra.ScrollHeight = fra.InsideHeight
    End If
    On Error GoTo 0
End Sub
