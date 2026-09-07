VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitSheetOptions 
   Caption         =   "Split Sheet By Column"
   ClientHeight    =   6684
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12264
   OleObjectBlob   =   "frmSplitSheetOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitSheetOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("TPD_Addin.SplitExport")

Option Explicit

Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Load columns into checkboxes + ComboBox
'===========================================================
Public Sub LoadColumns(headingList As Variant)
    Dim i As Long

    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkSplit"

    ' Pre-fill: this picker's own last-used selection, else the Set TPD
    ' Defaults value, else the shipped default list (#96, #98).
    ApplyColumnSelection fraColumns, _
        ResolveColumnList(Array(PREF_SPLIT_PICKER_COLUMNS, PREF_SPLIT_COLUMNS), _
                          DefaultSplitColumns())

    cboGroupColumn.Clear
    For i = LBound(headingList) To UBound(headingList)
        cboGroupColumn.AddItem headingList(i)
    Next i

    SelectGroupColumn ResolveGroupColumn( _
        Array(PREF_SPLIT_PICKER_GROUPCOL, PREF_SPLIT_GROUPCOL), _
        DefaultSplitGroupColumn())
End Sub

' Picks the group-by column after the dropdown is populated: wantGroup
' (already resolved through the LastUsed* -> DefaultUser* -> shipped
' "Vendor" chain) if it's one of the columns on this sheet, otherwise
' the first column.
Private Sub SelectGroupColumn(ByVal wantGroup As String)
    Dim i As Long
    Dim want As String

    ' Normalize both sides - a saved preference can carry stray whitespace
    ' (older builds, a paste with a leading tab) that a raw StrComp against
    ' the real header would miss.
    want = NormalizeCellText(wantGroup)
    If Len(want) > 0 Then
        For i = 0 To cboGroupColumn.ListCount - 1
            If StrComp(NormalizeCellText(cboGroupColumn.List(i)), want, vbTextCompare) = 0 Then
                cboGroupColumn.ListIndex = i
                Exit Sub
            End If
        Next i
    End If

    If cboGroupColumn.ListCount > 0 Then cboGroupColumn.ListIndex = 0
End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()

    If cboGroupColumn.ListIndex = -1 Then
        MsgBox "Please select a group column.", vbExclamation
        Exit Sub
    End If

    If Not HasColumnSelection(fraColumns) Then
        MsgBox "Please select at least one column to keep.", vbExclamation
        Exit Sub
    End If

    SavePref PREF_SPLIT_PICKER_GROUPCOL, cboGroupColumn.value
    SaveColumnList PREF_SPLIT_PICKER_COLUMNS, GetSelectedColumns(fraColumns)

    CancelPressed = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    CancelPressed = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelPressed = True
        Me.Hide
    End If
End Sub
