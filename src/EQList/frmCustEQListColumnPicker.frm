VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustEQListColumnPicker 
   Caption         =   "Customer EQ List"
   ClientHeight    =   6936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12984
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
' Chrome - brand band, help line, About links (#118). The
' styling that lives in code rather than the .frx (so it shows
' up in review) is shared with the other dialogs in
' modHelpers_DialogChrome, exactly as frmSetTPDDefaults does it.
'===========================================================
Private Sub UserForm_Initialize()
    ApplyDialogChrome lblBrandBar, lblHelp
    InitAboutLinks lblVersion, lblUserGuide
End Sub

Private Sub lblVersion_Click()
    modAbout.OpenReleasePage
End Sub

Private Sub lblUserGuide_Click()
    modAbout.OpenUserGuide
End Sub

'===========================================================
' Column grid actions (#149). No persistence - cmdOK_Click
' still writes PREF_EQLIST_PICKER_COLUMNS. Restore defaults
' resolves DefaultUser* -> shipped, skipping LastUsed*, so it
' means the same as it does on the Set Defaults EQ List tab.
'===========================================================
Private Sub cmdSelectAll_Click()
    SetAllColumns fraColumns, True
End Sub

Private Sub cmdSelectNone_Click()
    SetAllColumns fraColumns, False
End Sub

Private Sub cmdRestoreColumns_Click()
    RestoreColumnDefaults fraColumns, PREF_EQLIST_COLUMNS, DefaultEqListColumns()
End Sub

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
