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
Private SelectedImgPath As String
Private PreferredDefaultColumns As Variant


Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

Public Property Get imgPath() As String
    imgPath = SelectedImgPath
End Property

'===========================================================
' Load columns into checkboxes
'===========================================================
Public Sub LoadColumns(headingList As Variant)

    Dim savedCols As Collection
    Dim applyDefaults As Boolean
    Dim i As Long

    ' Define default columns
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

    ' Build checkboxes
    LayoutCheckboxes fraColumns, headingList, 10, "chkCustEQ"

    ' Load saved preferences
    Set savedCols = LoadColumnList(PREF_CUSTEQ_COLUMNS)
    applyDefaults = (savedCols Is Nothing) Or (savedCols.Count = 0)

    ' Apply defaults if no saved prefs
    If applyDefaults Then
        Dim ctrl As control
        For Each ctrl In fraColumns.Controls
            If TypeName(ctrl) = "CheckBox" Then
                If IsInArray(ctrl.Caption, PreferredDefaultColumns) Then
                    ctrl.value = True
                End If
            End If
        Next ctrl
    Else
        ' Apply saved preferences
        AutoSelectColumns fraColumns, savedCols
    End If

End Sub



'===========================================================
' Initialize form
'===========================================================
Private Sub UserForm_Initialize()
    Dim savedCols As Collection

    SelectedImgPath = LoadPref(PREF_CUSTEQ_IMAGE, "")
    txtImgPath.Text = SelectedImgPath

    Set savedCols = LoadColumnList(PREF_CUSTEQ_COLUMNS)
    AutoSelectColumns fraColumns, savedCols
End Sub

'===========================================================
' Browse for image
'===========================================================
Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Logo Image"
        .Filters.Clear
        .Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectedImgPath = .SelectedItems(1)
            txtImgPath.Text = SelectedImgPath
        End If
    End With
End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()
    SavePref PREF_CUSTEQ_IMAGE, SelectedImgPath
    SaveColumnList PREF_CUSTEQ_COLUMNS, GetSelectedColumns(fraColumns)

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


