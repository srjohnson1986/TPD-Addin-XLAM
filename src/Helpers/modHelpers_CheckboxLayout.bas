Attribute VB_Name = "modHelpers_CheckboxLayout"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Sub LayoutCheckboxes( _
        ByVal fra As MSForms.Frame, _
        ByVal headerList As Variant, _
        ByVal rowsPerColumn As Long, _
        ByVal baseName As String, _
        Optional ByVal preferredDefaults As Variant = Empty, _
        Optional ByVal applyDefaults As Boolean = False)

    Dim i As Long
    Dim idx As Long
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim chk As MSForms.CheckBox
    Dim chkLeft As Single
    Dim chkTop As Single
    Dim chkWidth As Single
    Dim hdr As String

    ' Clear existing checkboxes
    For idx = fra.Controls.Count - 1 To 0 Step -1
        If TypeName(fra.Controls(idx)) = "CheckBox" Then
            fra.Controls.Remove fra.Controls(idx).name
        End If
    Next idx

    chkWidth = (fra.Width - 12) / 3

    colIndex = 0
    rowIndex = 0

    For i = LBound(headerList) To UBound(headerList)
        hdr = CStr(headerList(i))

        chkLeft = 6 + (colIndex * chkWidth)
        chkTop = 6 + (rowIndex * 18)

        Set chk = fra.Controls.Add("Forms.CheckBox.1", _
                    baseName & "_" & i & "_" & Replace(hdr, " ", "_"))
        chk.Caption = hdr
        chk.Left = chkLeft
        chk.Top = chkTop
        chk.Width = chkWidth - 6

        rowIndex = rowIndex + 1
        If rowIndex = rowsPerColumn Then
            rowIndex = 0
            colIndex = colIndex + 1
        End If
    Next i
End Sub



