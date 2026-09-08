VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustScheduleColumnPicker 
   Caption         =   "Customer Schedule"
   ClientHeight    =   6936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12984
   OleObjectBlob   =   "frmCustScheduleColumnPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustScheduleColumnPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'@Folder("TPD_Addin.Forms")

Option Explicit

'===========================================================
'  Column picker for the "Create Customer Schedule" flow -
'  the schedule twin of frmCustEQListColumnPicker. LoadColumns
'  builds the checkbox grid then applies the resolved selection
'  (this picker's own last-used, else the Set TPD Defaults
'  value, else the shipped default list); OK saves the ticked
'  set to the picker-only PREF_SCHEDULE_PICKER_COLUMNS key after
'  checking at least one column is ticked.
'
'  The .frx started as a byte-clone of frmCustEQListColumnPicker.frx
'  (#113). The shared lblTitle now carries "Customer Schedule" as a
'  design-time caption (#118 folded in #114) - the old run-time
'  retext hack is gone.
'===========================================================

Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean


Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Chrome - brand band, help line, About links (#118). Kept
' byte-identical to frmCustEQListColumnPicker: the two pickers
' are visual twins.
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
' Column grid actions (#149). Kept identical to
' frmCustEQListColumnPicker - only the pref key and shipped
' list differ. Restore defaults resolves DefaultUser* ->
' shipped, skipping LastUsed*.
'===========================================================
Private Sub cmdSelectAll_Click()
    SetAllColumns fraColumns, True
End Sub

Private Sub cmdSelectNone_Click()
    SetAllColumns fraColumns, False
End Sub

Private Sub cmdRestoreColumns_Click()
    RestoreColumnDefaults fraColumns, PREF_SCHEDULE_COLUMNS, DefaultScheduleColumns()
End Sub

'===========================================================
' Load columns into checkboxes
'===========================================================
' Grid + pre-fill are modHelpers_ColumnPicker's job, shared with the other
' two pickers: last-used selection, else the Set Defaults value, else the
' shipped default list (#96, #112).
Public Sub LoadColumns(headingList As Variant)
    InitColumnPicker fraColumns, headingList, ROWS_PER_COLUMN, "chkCustSched", _
                     PREF_SCHEDULE_PICKER_COLUMNS, PREF_SCHEDULE_COLUMNS, _
                     DefaultScheduleColumns()
End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()
    ' Guards "at least one column" and writes the picker-only LastUsed* key.
    If Not CommitColumnPicker(fraColumns, PREF_SCHEDULE_PICKER_COLUMNS) Then Exit Sub

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

