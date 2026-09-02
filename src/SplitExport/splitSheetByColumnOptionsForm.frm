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
    Dim savedCols As Collection
    Dim i As Long
    Dim applyDefaults As Boolean

    ' Initialize defaults here
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

    ' Build checkboxes
    LayoutCheckboxes fraColumns, headingList, ROWS_PER_COLUMN, "chkSplit", PreferredDefaultColumns, applyDefaults

    ' Populate group column dropdown
    cboGroupColumn.Clear
    For i = LBound(headingList) To UBound(headingList)
        cboGroupColumn.AddItem headingList(i)
     
    Next i
    
    '-----------------------------------------------------------
    ' Auto-select "Vendor" in cboGroupColumn if available
    '-----------------------------------------------------------
    For i = 0 To cboGroupColumn.ListCount - 1
        If StrComp(cboGroupColumn.List(i), "Vendor", vbTextCompare) = 0 Then
            cboGroupColumn.ListIndex = i
            Exit For
        End If
    Next i

    ' Load saved preferences
    Set savedCols = LoadColumnList(PREF_SPLIT_COLUMNS)
    applyDefaults = (savedCols Is Nothing) Or (savedCols.Count = 0)

    ' If no saved prefs ? apply defaults by caption
    If applyDefaults Then
        Dim ctrl As control
        For Each ctrl In fraColumns.Controls
            If TypeName(ctrl) = "CheckBox" Then
                If IsInArray(ctrl.Caption, PreferredDefaultColumns) Then
                    Debug.Print "MATCHED DEFAULT: " & ctrl.Caption
                    ctrl.value = True
                End If
            End If
        Next ctrl
    End If

    ' If saved prefs exist ? override defaults
    If Not (savedCols Is Nothing) Then
        If savedCols.Count > 0 Then
            AutoSelectColumns fraColumns, savedCols
        End If
    End If
End Sub

'===========================================================
' Initialize form
'===========================================================
Private Sub UserForm_Initialize()

    Dim savedGroup As String
    Dim savedCols As Collection
    Dim i As Long
    Dim vendorIndex As Long

    savedGroup = LoadPref(PREF_SPLIT_GROUPCOL, "")
    Set savedCols = LoadColumnList(PREF_SPLIT_COLUMNS)

    AutoSelectColumns fraColumns, savedCols

    ' 1. Try saved preference
    If Len(savedGroup) > 0 Then
        cboGroupColumn.value = savedGroup
    End If
    If cboGroupColumn.ListIndex <> -1 Then Exit Sub

    ' 2. Try Vendor
    vendorIndex = -1
    For i = 0 To cboGroupColumn.ListCount - 1
        If StrComp(cboGroupColumn.List(i), "Vendor", vbTextCompare) = 0 Then
            vendorIndex = i
            Exit For
        End If
    Next i

    If vendorIndex <> -1 Then
        cboGroupColumn.ListIndex = vendorIndex
        Exit Sub
    End If

    ' 3. Default to first column
    If cboGroupColumn.ListCount > 0 Then
        cboGroupColumn.ListIndex = 0
    End If

End Sub

'===========================================================
' OK / Cancel
'===========================================================
Private Sub cmdOK_Click()

    If cboGroupColumn.ListIndex = -1 Then
        MsgBox "Please select a group column.", vbExclamation
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


