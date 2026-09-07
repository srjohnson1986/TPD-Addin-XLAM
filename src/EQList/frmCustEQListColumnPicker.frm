VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustEQListColumnPicker 
   Caption         =   "Customer EQ List Generator"
   ClientHeight    =   7260
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12552
   OleObjectBlob   =   "frmCustEQListColumnPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustEQListColumnPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("TPD_Addin.EQList")

Option Explicit


Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean


Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Load columns into checkboxes
'===========================================================
Public Sub LoadColumns(headingList As Variant)

    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkCustEQ"

    ' Pre-fill: this picker's own last-used selection, else the Set TPD
    ' Defaults value, else the shipped default list (#96, #98).
    ApplyColumnSelection fraColumns, _
        ResolveColumnList(Array(PREF_EQLIST_PICKER_COLUMNS, PREF_EQLIST_COLUMNS), _
                          DefaultEqListColumns())

End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()
    If Not HasColumnSelection(fraColumns) Then
        MsgBox "Please select at least one column to keep.", vbExclamation
        Exit Sub
    End If

    SaveColumnList PREF_EQLIST_PICKER_COLUMNS, GetSelectedColumns(fraColumns)

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
