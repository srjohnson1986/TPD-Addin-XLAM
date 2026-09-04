VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustEQListColumnPickerForm 
   Caption         =   "Customer EQ List Generator"
   ClientHeight    =   7260
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12552
   OleObjectBlob   =   "CustEQListColumnPickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustEQListColumnPickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("TPD_Addin.EQList")

Option Explicit


Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean
Private PreferredDefaultColumns As Variant


Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Load columns into checkboxes
'===========================================================
Public Sub LoadColumns(headingList As Variant)

    PreferredDefaultColumns = Array( _
        "INTERNAL ID", _
        "CUSTOMER ID", _
        "Location", _
        "TYPE1", _
        "TYPE2", _
        "SIZE", _
        "RATING", _
        "CONNECTION TYPE", _
        "Description", _
        "Manufacturer", _
        "Vendor", _
        "Model", _
        "Details" _
    )

    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkCustEQ"
    ApplyColumnSelection fraColumns, PREF_EQLIST_COLUMNS, PreferredDefaultColumns

End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()
    SaveColumnList PREF_EQLIST_COLUMNS, GetSelectedColumns(fraColumns)

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


