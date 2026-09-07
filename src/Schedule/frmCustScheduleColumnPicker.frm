VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustScheduleColumnPicker 
   Caption         =   "Customer Schedule Generator"
   ClientHeight    =   7260
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12552
   OleObjectBlob   =   "frmCustScheduleColumnPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustScheduleColumnPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("TPD_Addin.Schedule")

Option Explicit

'===========================================================
'  Column picker for the "Create Customer Schedule" flow -
'  the schedule twin of CustEQListColumnPickerForm. LoadColumns
'  builds the checkbox grid then applies the resolved selection
'  (this picker's own last-used, else the Set TPD Defaults
'  value, else the shipped default list); OK saves the ticked
'  set to the picker-only PREF_SCHEDULE_PICKER_COLUMNS key after
'  checking at least one column is ticked.
'
'  The .frx started as a byte-clone of CustEQListColumnPickerForm.frx
'  (#113); lblCustScheduleTitle was renamed from the EQ control
'  name, and its "Customer EQ List Generator" design-time caption
'  is overridden in UserForm_Initialize (#114).
'===========================================================

Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean


Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

Private Sub UserForm_Initialize()
    ' The .frx still holds the cloned "Customer EQ List Generator" caption (#114).
    lblCustScheduleTitle.Caption = "Customer Schedule Generator"
End Sub

'===========================================================
' Load columns into checkboxes
'===========================================================
Public Sub LoadColumns(headingList As Variant)

    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkCustSched"

    ' Pre-fill: this picker's own last-used selection, else the Set TPD
    ' Defaults value, else the shipped default list (#96, #112).
    ApplyColumnSelection fraColumns, _
        ResolveColumnList(Array(PREF_SCHEDULE_PICKER_COLUMNS, PREF_SCHEDULE_COLUMNS), _
                          DefaultScheduleColumns())

End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()
    If Not HasColumnSelection(fraColumns) Then
        MsgBox "Please select at least one column to keep.", vbExclamation
        Exit Sub
    End If

    SaveColumnList PREF_SCHEDULE_PICKER_COLUMNS, GetSelectedColumns(fraColumns)

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

