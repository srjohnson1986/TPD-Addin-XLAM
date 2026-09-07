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

' Logo tab: which change is staged for OK (Cancel discards it).
Private Const LOGO_NONE As Long = 0
Private Const LOGO_CHOOSE As Long = 1
Private Const LOGO_RESTORE As Long = 2
Private mLogoAction As Long
Private mLogoChosenPath As String


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

    lblVersion.Caption = "v" & ADDIN_VERSION

    LoadValues
    mpgPages.value = 0
    lblFirstRunNotice.Visible = modPreferences.DefaultsNeverSaved()

    imgLogoPreview.PictureSizeMode = fmPictureSizeModeZoom
    RefreshLogoTab
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


'--- Logo tab (#95 / #109) -------------------------------------------------
'
' Choosing an image or restoring the built-in one is STAGED - it updates the
' preview but nothing is written until OK (Cancel discards it), like the
' column-list fields. ApplyPendingLogo runs in cmdOk_Click before the column
' save. A chosen image previews via GDI+ (modHelpers_Logo.LogoPreviewPicture);
' the built-in preview reuses imgBrandLogo's picture - the same logo already
' shown on this dialog's brand bar - since there's no reliable way to render
' the _Resources shape into an Image control (#109).

Private Sub cmdChooseLogo_Click()
    Dim fd As FileDialog
    Dim picked As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Choose a logo image"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.png; *.jpg; *.jpeg; *.gif; *.bmp"

    If fd.Show <> -1 Then Exit Sub
    picked = fd.SelectedItems(1)

    If modGdiPlus.LoadPictureGDIP(picked) Is Nothing Then
        MsgBox "That image couldn't be read. Try a PNG, JPG, GIF or BMP file.", _
               vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    mLogoAction = LOGO_CHOOSE
    mLogoChosenPath = picked
    RefreshLogoTab
End Sub

Private Sub cmdRestoreLogo_Click()
    mLogoAction = LOGO_RESTORE
    mLogoChosenPath = vbNullString
    RefreshLogoTab
End Sub

' Preview + status for what the tab currently shows: the staged change if
' there is one, else the saved state.
Private Sub RefreshLogoTab()
    Dim pic As stdole.IPictureDisp
    Dim pendingPath As String
    Dim savedPath As String
    Dim logoName As String

    If mLogoAction = LOGO_CHOOSE Then pendingPath = mLogoChosenPath
    Set pic = modHelpers_Logo.LogoPreviewPicture(pendingPath, mLogoAction = LOGO_RESTORE)

    On Error Resume Next
    If pic Is Nothing Then
        Set imgLogoPreview.Picture = imgBrandLogo.Picture   ' the built-in logo
    Else
        Set imgLogoPreview.Picture = pic
    End If
    On Error GoTo 0

    Select Case mLogoAction
        Case LOGO_CHOOSE
            logoName = FileNameOnly(mLogoChosenPath)
        Case LOGO_RESTORE
            logoName = "the built-in logo"
        Case Else
            savedPath = modHelpers_Logo.EffectiveLogoPath()
            If Len(savedPath) > 0 Then logoName = FileNameOnly(savedPath) _
            Else logoName = "the built-in logo"
    End Select

    If mLogoAction = LOGO_NONE Then
        lblLogoStatus.Caption = "Currently using " & logoName & "."
    Else
        lblLogoStatus.Caption = "On OK: use " & logoName & "."
    End If
End Sub

Private Function FileNameOnly(ByVal path As String) As String
    Dim slash As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then FileNameOnly = Mid$(path, slash + 1) Else FileNameOnly = path
End Function

' Applies the staged logo change. "" on success or nothing staged, else a
' reason (the caller keeps the form open). Runs before the column save so a
' bad image can't leave saved columns with no logo (#95).
Private Function ApplyPendingLogo() As String
    Select Case mLogoAction
        Case LOGO_CHOOSE
            ApplyPendingLogo = modHelpers_Logo.SetUserLogo(mLogoChosenPath)
        Case LOGO_RESTORE
            modHelpers_Logo.ClearUserLogo
    End Select
End Function


'--- OK / Cancel (spec 6, 7) -------------------------------------------------

Private Sub cmdOk_Click()
    Dim eqList As String, scheduleCols As String
    Dim splitCols As String, groupColumn As String
    Dim emptyFields As String
    Dim logoErr As String

    eqList = NormalizeColumnList(txtEqListColumns.Text)
    scheduleCols = NormalizeColumnList(txtScheduleColumns.Text)
    splitCols = NormalizeColumnList(txtSplitColumns.Text)
    groupColumn = CollapseWhitespace(txtSplitGroupColumn.Text)

    ' A column list is never a meaningful "empty" - every consumer reads an
    ' empty saved value as "fall back to the built-in list" - so an empty
    ' field here would just silently mean "defaults". Make the user choose.
    ' The Split group-column field is not checked: an empty group pref
    ' harmlessly falls back to "Vendor" / first column in the picker.
    emptyFields = EmptyColumnFields(eqList, scheduleCols, splitCols)
    If Len(emptyFields) > 0 Then
        MsgBox "These column lists can't be left empty: " & emptyFields & "." & vbCrLf & vbCrLf & _
               "Type the columns you want, or use Restore defaults on that tab for " & _
               "the built-in list.", vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    ' Apply a staged logo change first - a bad image must not leave columns
    ' saved with no logo (#95).
    logoErr = ApplyPendingLogo()
    If Len(logoErr) > 0 Then
        MsgBox logoErr, vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    If Not SaveAllDefaults(eqList, scheduleCols, splitCols, groupColumn) Then
        MsgBox "Couldn't save your defaults. Your changes are still open behind " & _
               "this message - choose OK to return to the form and try again.", _
               vbCritical, "TPD Add-in"
        Exit Sub
    End If

    Me.Hide
    ShowSavedInStatusBar
End Sub

' Names (comma-joined) of the column-list fields that normalized to empty,
' or "" when all three hold a value. Order matches the tabs.
Private Function EmptyColumnFields(ByVal eqList As String, _
                                   ByVal scheduleCols As String, _
                                   ByVal splitCols As String) As String
    Dim names As String

    If Len(eqList) = 0 Then names = names & ", EQ List"
    If Len(scheduleCols) = 0 Then names = names & ", Schedule"
    If Len(splitCols) = 0 Then names = names & ", Split Sheets"

    If Len(names) > 0 Then EmptyColumnFields = Mid$(names, 3)
End Function

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
