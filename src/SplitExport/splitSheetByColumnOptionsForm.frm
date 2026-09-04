VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} splitSheetByColumnOptionsForm 
   Caption         =   "Split Sheet By Column"
   ClientHeight    =   6684
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12264
   OleObjectBlob   =   "splitSheetByColumnOptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "splitSheetByColumnOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("TPD_Addin.SplitExport")

Option Explicit

Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean
Private PreferredDefaultColumns As Variant

Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Load columns into checkboxes + ComboBox
'===========================================================
Public Sub LoadColumns(headingList As Variant)
    Dim i As Long

    PreferredDefaultColumns = Array( _
        "INTERNAL ID", _
        "CUSTOMER ID", _
        "SIZE", _
        "RATING", _
        "CONNECTION TYPE", _
        "Description", _
        "Manufacturer", _
        "Vendor", _
        "Model", _
        "Details" _
    )

    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkSplit"
    ApplyColumnSelection fraColumns, PREF_SPLIT_COLUMNS, PreferredDefaultColumns

    cboGroupColumn.Clear
    For i = LBound(headingList) To UBound(headingList)
        cboGroupColumn.AddItem headingList(i)
    Next i

    SelectGroupColumn LoadPref(PREF_SPLIT_GROUPCOL, "")
End Sub

' Picks the group-by column after the dropdown is populated: the saved
' preference if it's still one of the columns on this sheet, otherwise
' "Vendor" if present, otherwise the first column.
Private Sub SelectGroupColumn(ByVal savedGroup As String)
    Dim i As Long

    If Len(savedGroup) > 0 Then
        For i = 0 To cboGroupColumn.ListCount - 1
            If StrComp(cboGroupColumn.List(i), savedGroup, vbTextCompare) = 0 Then
                cboGroupColumn.ListIndex = i
                Exit Sub
            End If
        Next i
    End If

    For i = 0 To cboGroupColumn.ListCount - 1
        If StrComp(cboGroupColumn.List(i), "Vendor", vbTextCompare) = 0 Then
            cboGroupColumn.ListIndex = i
            Exit Sub
        End If
    Next i

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

    SavePref PREF_SPLIT_GROUPCOL, cboGroupColumn.value
    SaveColumnList PREF_SPLIT_COLUMNS, GetSelectedColumns(fraColumns)

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


