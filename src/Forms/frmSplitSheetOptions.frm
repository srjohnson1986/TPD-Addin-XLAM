VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitSheetOptions 
   Caption         =   "Split Sheet by Column"
   ClientHeight    =   7836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12984
   OleObjectBlob   =   "frmSplitSheetOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitSheetOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'@Folder("TPD_Addin.Forms")

Option Explicit

Private Const ROWS_PER_COLUMN As Long = 10
Private CancelPressed As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = CancelPressed
End Property

'===========================================================
' Chrome - brand band, help line, About links (#118), shared
' with the other dialogs via modHelpers_DialogChrome.
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
' Column grid actions (#149). These act on fraColumns only -
' never cboGroupColumn. No persistence; cmdOK_Click still
' writes the PREF_SPLIT_PICKER_* keys. Restore defaults
' resolves DefaultUser* -> shipped, skipping LastUsed*.
'===========================================================
Private Sub cmdSelectAll_Click()
    SetAllColumns fraColumns, True
End Sub

Private Sub cmdSelectNone_Click()
    SetAllColumns fraColumns, False
End Sub

Private Sub cmdRestoreColumns_Click()
    RestoreColumnDefaults fraColumns, PREF_SPLIT_COLUMNS, DefaultSplitColumns()
End Sub

'===========================================================
' Load columns into checkboxes + ComboBox
'===========================================================
' The checkbox grid and its pre-fill are modHelpers_ColumnPicker's job,
' shared with the other two pickers (last-used, else Set Defaults, else the
' shipped list - #96, #98). The group-column dropdown below is this form's
' alone.
Public Sub LoadColumns(headingList As Variant)
    Dim i As Long

    InitColumnPicker fraColumns, headingList, ROWS_PER_COLUMN, "chkSplit", _
                     PREF_SPLIT_PICKER_COLUMNS, PREF_SPLIT_COLUMNS, _
                     DefaultSplitColumns()

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

    ' Guards "at least one column" and writes the picker-only column key;
    ' the group column is this form's own, saved alongside it.
    If Not CommitColumnPicker(fraColumns, PREF_SPLIT_PICKER_COLUMNS) Then Exit Sub

    SavePref PREF_SPLIT_PICKER_GROUPCOL, cboGroupColumn.value

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
