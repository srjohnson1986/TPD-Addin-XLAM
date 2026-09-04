VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetTPDDefaults 
   Caption         =   "Set TPD Defaults"
   ClientHeight    =   8232.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13188
   OleObjectBlob   =   "frmSetTPDDefaults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetTPDDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TPD_Addin.Preferences")

'===========================================================
'  "Set TPD Defaults" - the one place to configure the default
'  column lists the per-run pickers fall back to. Storage,
'  parsing and the hard-coded fallback lists live in
'  modPreferences / modPreferences_Defaults; the ribbon entry
'  point and status-bar reset live in modMain_SetTPDDefaults.
'
'  Behaviour spec: design_handoff_tpd_addin_defaults/design/
'  "TPD Addin Defaults - Behavior Spec.dc.html".
'===========================================================

Option Explicit

Private Const STATUS_MESSAGE As String = "TPD defaults saved"
Private Const STATUS_CLEAR_SECONDS As Long = 4

' Ctrl bit in a KeyUp Shift argument.
Private Const CTRL_MASK As Integer = 2


'--- Lifecycle -------------------------------------------------------------

Private Sub UserForm_Initialize()
    Me.Caption = "Set TPD Defaults"

    ' TPD grey #64665D - VBA BackColor is BGR, not RGB.
    lblBrandBar.BackColor = &H5D6664
    lblBrandBar.Caption = vbNullString

    txtEqListColumns.Multiline = True
    txtEqListColumns.ScrollBars = fmScrollBarsVertical
    txtScheduleColumns.Multiline = True
    txtScheduleColumns.ScrollBars = fmScrollBarsVertical
    txtSplitColumns.Multiline = True
    txtSplitColumns.ScrollBars = fmScrollBarsVertical

    imgLogoPreview.PictureSizeMode = fmPictureSizeModeZoom

    lblVersion.Caption = "v" & ADDIN_VERSION

    LoadValues
    mpgPages.value = 0
    lblFirstRunNotice.Visible = modPreferences.DefaultsNeverSaved()
End Sub

Private Sub UserForm_Activate()
    ShowEmbeddedLogo
End Sub

Private Sub LoadValues()
    txtEqListColumns.Text = LoadPref(PREF_EQLIST_COLUMNS, DefaultEqListColumns())
    txtScheduleColumns.Text = LoadPref(PREF_SCHEDULE_COLUMNS, DefaultScheduleColumns())
    txtSplitColumns.Text = LoadPref(PREF_SPLIT_COLUMNS, DefaultSplitColumns())
    txtSplitGroupColumn.Text = LoadPref(PREF_SPLIT_GROUPCOL, DefaultSplitGroupColumn())
End Sub


'--- Column boxes: normalize on paste and on losing focus (spec 2) --------

Private Sub txtEqListColumns_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtEqListColumns.Text = NormalizeColumnList(txtEqListColumns.Text)
End Sub

Private Sub txtScheduleColumns_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtScheduleColumns.Text = NormalizeColumnList(txtScheduleColumns.Text)
End Sub

Private Sub txtSplitColumns_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtSplitColumns.Text = NormalizeColumnList(txtSplitColumns.Text)
End Sub

Private Sub txtSplitGroupColumn_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtSplitGroupColumn.Text = CollapseWhitespace(txtSplitGroupColumn.Text)
End Sub

' MSForms has no paste event, so catch Ctrl+V and normalize just after the
' text lands - otherwise the user sees tab-separated text until they tab out.
Private Sub txtEqListColumns_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If IsPasteChord(KeyCode, Shift) Then
        txtEqListColumns.Text = NormalizeColumnList(txtEqListColumns.Text)
    End If
End Sub

Private Sub txtScheduleColumns_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If IsPasteChord(KeyCode, Shift) Then
        txtScheduleColumns.Text = NormalizeColumnList(txtScheduleColumns.Text)
    End If
End Sub

Private Sub txtSplitColumns_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If IsPasteChord(KeyCode, Shift) Then
        txtSplitColumns.Text = NormalizeColumnList(txtSplitColumns.Text)
    End If
End Sub

Private Function IsPasteChord(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) As Boolean
    IsPasteChord = (KeyCode = vbKeyV And (Shift And CTRL_MASK) = CTRL_MASK)
End Function


'--- Restore defaults, per page (spec 6) ---------------------------------

Private Sub cmdRestoreEqList_Click()
    txtEqListColumns.Text = DefaultEqListColumns()
End Sub

Private Sub cmdRestoreSchedule_Click()
    txtScheduleColumns.Text = DefaultScheduleColumns()
End Sub

Private Sub cmdRestoreSplitSheets_Click()
    txtSplitColumns.Text = DefaultSplitColumns()
    txtSplitGroupColumn.Text = DefaultSplitGroupColumn()
End Sub


'--- Logo preview: read-only view of the embedded _Resources logo --------

Private Sub ShowEmbeddedLogo()
    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Resources")
    If ws Is Nothing Then Exit Sub
    Set shp = ws.Shapes("DefaultLogo")
    If shp Is Nothing Then Exit Sub

    shp.Copy
    Set imgLogoPreview.Picture = PastePicture()
    On Error GoTo 0
End Sub


'--- OK / Cancel (spec 6, 7) -------------------------------------------------

Private Sub cmdOk_Click()
    Dim eqList As String, scheduleCols As String
    Dim splitCols As String, groupColumn As String

    eqList = NormalizeColumnList(txtEqListColumns.Text)
    scheduleCols = NormalizeColumnList(txtScheduleColumns.Text)
    splitCols = NormalizeColumnList(txtSplitColumns.Text)
    groupColumn = CollapseWhitespace(txtSplitGroupColumn.Text)

    If Not SaveAllDefaults(eqList, scheduleCols, splitCols, groupColumn) Then
        MsgBox "Couldn't save your defaults. Your changes are still open behind " & _
               "this message - choose OK to return to the form and try again.", _
               vbCritical, "TPD Add-in"
        Exit Sub
    End If

    Me.Hide
    ShowSavedInStatusBar
End Sub

Private Sub cmdCancel_Click()
    Me.Hide          ' discard, no warning (spec 6)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide      ' title-bar X behaves as Cancel
    End If
End Sub

' Confirmation goes to the status bar, so there is nothing to dismiss (spec 7).
Private Sub ShowSavedInStatusBar()
    Application.StatusBar = STATUS_MESSAGE
    Application.OnTime Now + TimeSerial(0, 0, STATUS_CLEAR_SECONDS), "ClearTPDDefaultsStatusBar"
End Sub

