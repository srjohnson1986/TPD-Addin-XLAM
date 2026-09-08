VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetTPDDefaults 
   Caption         =   "Set Defaults"
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

'@Folder("TPD_Addin.Forms")

'===========================================================
'  "Set Defaults" - the one place to configure the default
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

' First-run notice (#99): gap in points between the brand band and the
' notice, and between the notice and the page control below it.
Private Const NOTICE_GAP As Single = 6

' True while UserForm_Initialize runs, so setting mpgPages.Value from the
' saved tab (#135) doesn't bounce straight back through mpgPages_Change.
Private mInitializing As Boolean

' Logo tab: which change is staged for OK (Cancel discards it).
Private Const LOGO_NONE As Long = 0
Private Const LOGO_CHOOSE As Long = 1
Private Const LOGO_RESTORE As Long = 2
Private mLogoAction As Long
Private mLogoChosenPath As String

' imgLogoPreview's design-time picture (set in the VBE from
' assets/tpdHeaderLogo.jpg) is the built-in header logo. Cached here on
' Initialize, before any override, so "Use built-in logo" can restore it.
Private mBuiltInLogoPic As stdole.IPictureDisp

'--- Lifecycle -------------------------------------------------------------

Private Sub UserForm_Initialize()
    mInitializing = True
    Me.Caption = "Set Defaults"

    ' Brand band, About links: the same shared chrome the picker / split /
    ' export dialogs use, so the five forms cannot drift apart (#118).
    ApplyDialogChrome lblBrandBar, Nothing

    txtEqListColumns.Multiline = True
    txtEqListColumns.ScrollBars = fmScrollBarsVertical
    txtScheduleColumns.Multiline = True
    txtScheduleColumns.ScrollBars = fmScrollBarsVertical
    txtSplitColumns.Multiline = True
    txtSplitColumns.ScrollBars = fmScrollBarsVertical

    modHelpers_DialogChrome.InitAboutLinks lblVersion, lblUserGuide

    LoadValues
    SetHelperText
    mpgPages.value = SavedTabIndex()
    SyncSaveRunVisibility          ' dialog reopens on the last-used tab (#135)
    ApplyFirstRunNoticeLayout

    Set mBuiltInLogoPic = imgLogoPreview.Picture   ' before RefreshLogoTab overrides it
    imgLogoPreview.PictureSizeMode = fmPictureSizeModeZoom
    RefreshLogoTab

    mInitializing = False
End Sub

'--- Last-used tab (#135) ---------------------------------------------
'
' Reopen on whichever tab was showing last. UI state only - stored
' outside the dialog's default keys so it can't affect the first-run
' notice. An unset or out-of-range value falls back to the first tab,
' which also keeps the #99 notice on EQ List where it's designed to be.

Private Function SavedTabIndex() As Long
    Dim idx As Long

    idx = val(LoadPref(PREF_SETDEFAULTS_LAST_TAB, "0"))
    If idx < 0 Or idx > mpgPages.Pages.count - 1 Then idx = 0
    SavedTabIndex = idx
End Function

Private Sub mpgPages_Change()
    If mInitializing Then Exit Sub
    SavePref PREF_SETDEFAULTS_LAST_TAB, CStr(mpgPages.value)
    SyncSaveRunVisibility
End Sub

'--- First-run notice layout (#99) -------------------------------------
'
' lblFirstRunNotice is a form-level control - a sibling of mpgPages, not
' a child of pgEqList - so it can't push one page's fields out of line
' with the others (issue #99, solution b). It only ever shows on a
' machine that has never saved defaults. When it does, drop it just
' below the brand band and push mpgPages down by the height it takes,
' shrinking the page control by the same amount so the OK / Cancel row
' and the About links (both below mpgPages) keep their place. When it's
' hidden - every run after the first save - the design-time layout is
' already correct and nothing moves.

Private Sub ApplyFirstRunNoticeLayout()
    Dim Shift As Single

    If Not modPreferences.DefaultsNeverSaved() Then
        lblFirstRunNotice.Visible = False
        Exit Sub
    End If

    lblFirstRunNotice.Top = lblBrandBar.Top + lblBrandBar.Height + NOTICE_GAP
    lblFirstRunNotice.Visible = True

    Shift = lblFirstRunNotice.Height + NOTICE_GAP
    mpgPages.Top = mpgPages.Top + Shift
    mpgPages.Height = mpgPages.Height - Shift
End Sub

'--- About links (#94) --------------------------------------------------
'
' lblVersion and lblUserGuide are plain labels in the bottom corner of the
' form (outside mpgPages, so they show on every tab). Captions, tooltips and
' the blue-underline link look come from modHelpers_DialogChrome.InitAboutLinks,
' called in UserForm_Initialize; the URLs live in modAbout. Only the two Click
' handlers stay here - a standard module can't wire an event.

Private Sub lblVersion_Click()
    modAbout.OpenReleasePage
End Sub

Private Sub lblUserGuide_Click()
    modAbout.OpenUserGuide
End Sub

'--- Per-page helper text (#143) --------------------------------------
'
' A muted hint under each tab's entry field, saying what a "column" is
' and that pasting a heading row straight out of Excel works. Captions
' live here rather than as design-time .frx properties so the copy stays
' in source and reviewable (same reason as InitAboutLinks). The labels
' themselves are placed in the VBE - lblEqListHelp / lblScheduleHelp /
' lblSplitColsHelp / lblSplitGroupHelp / lblLogoHelp, all WordWrap, no
' TabStop. Width is forced here too: the design-time boxes were narrow
' enough to wrap the one-line hints to three cramped lines.

Private Sub SetHelperText()
    Const PASTE_HINT As String = _
        "Header names from row 1, in output order. " & _
        "Paste a row from Excel - tabs become commas."

    lblEqListHelp.Caption = PASTE_HINT
    lblScheduleHelp.Caption = PASTE_HINT
    lblSplitColsHelp.Caption = "Header names from row 1, in output order."
    lblSplitGroupHelp.Caption = "One header name. A sheet is created per unique value."
    lblLogoHelp.Caption = "PNG, JPG, GIF or BMP. The file is copied into the " & _
                          "add-in, so it applies to every workbook."

    StyleHelperLabel lblEqListHelp
    StyleHelperLabel lblScheduleHelp
    StyleHelperLabel lblSplitColsHelp
    StyleHelperLabel lblSplitGroupHelp
    StyleHelperLabel lblLogoHelp
End Sub

Private Sub StyleHelperLabel(ByVal lbl As MSForms.Label)
    lbl.ForeColor = RGB(90, 90, 90)   ' grey #5A5A5A (RGB() handles BGR)
    lbl.Width = 500                   ' points - wide enough to keep each hint on one line
End Sub

Private Sub LoadValues()
    txtEqListColumns.Text = LoadPref(PREF_EQLIST_COLUMNS, DefaultEqListColumns())
    txtScheduleColumns.Text = LoadPref(PREF_SCHEDULE_COLUMNS, DefaultScheduleColumns())
    txtSplitColumns.Text = LoadPref(PREF_SPLIT_COLUMNS, DefaultSplitColumns())
    txtSplitGroupColumn.Text = LoadPref(PREF_SPLIT_GROUPCOL, DefaultSplitGroupColumn())
    LoadEqListToggles
End Sub

'--- EQ List behaviour toggles (#97 / #119-#123) ----------------------
'
' Five checkboxes on the EQ List page, "1"/"0" prefs, honoured by
' modMain_CustEQList.CreateCustEQList_DoWork. Saved by CommitDefaults after
' the column keys, reset by cmdRestoreEqList_Click. Which one ships on is
' not written here - LoadEqListToggle reads it from the toggle table in
' modPreferences_Defaults, the one place that knows.
'
' These two routines are the only place a checkbox is named out loud, and
' they stay explicit for that reason: a control referenced by name in a loop
' would be a run-time 438 when renamed, where this is a compile error.

Private Sub LoadEqListToggles()
    cbxRemoveParentRows.value = LoadEqListToggle(PREF_EQLIST_REMOVE_PARENT_ROWS)
    cbxAddItemCountColumn.value = LoadEqListToggle(PREF_EQLIST_ADD_COUNT_COLUMN)
    cbxPlainParentRows.value = LoadEqListToggle(PREF_EQLIST_PLAIN_PARENT_ROWS)
    cbxAddCellBorders.value = LoadEqListToggle(PREF_EQLIST_ADD_CELL_BORDERS)
    cbxAddColumnFilters.value = LoadEqListToggle(PREF_EQLIST_ADD_COLUMN_FILTERS)
End Sub

Private Sub SaveEqListToggles()
    SaveToggle PREF_EQLIST_REMOVE_PARENT_ROWS, CBool(cbxRemoveParentRows.value)
    SaveToggle PREF_EQLIST_ADD_COUNT_COLUMN, CBool(cbxAddItemCountColumn.value)
    SaveToggle PREF_EQLIST_PLAIN_PARENT_ROWS, CBool(cbxPlainParentRows.value)
    SaveToggle PREF_EQLIST_ADD_CELL_BORDERS, CBool(cbxAddCellBorders.value)
    SaveToggle PREF_EQLIST_ADD_COLUMN_FILTERS, CBool(cbxAddColumnFilters.value)
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
    ' The behaviour toggles live on this page, so Restore defaults resets them
    ' too - each back to what it ships as, read from the toggle table rather
    ' than repeated as literals here (#122 put "Add cell borders" on).
    cbxRemoveParentRows.value = EqListToggleShipsOn(PREF_EQLIST_REMOVE_PARENT_ROWS)
    cbxAddItemCountColumn.value = EqListToggleShipsOn(PREF_EQLIST_ADD_COUNT_COLUMN)
    cbxPlainParentRows.value = EqListToggleShipsOn(PREF_EQLIST_PLAIN_PARENT_ROWS)
    cbxAddCellBorders.value = EqListToggleShipsOn(PREF_EQLIST_ADD_CELL_BORDERS)
    cbxAddColumnFilters.value = EqListToggleShipsOn(PREF_EQLIST_ADD_COLUMN_FILTERS)
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
' the built-in preview is imgLogoPreview's own design-time picture, cached in
' mBuiltInLogoPic (there's no reliable way to render the _Resources shape into
' an Image control - #109).

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
        Set imgLogoPreview.Picture = mBuiltInLogoPic
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

'--- Save / Save & Run / Cancel (spec 6, 7; #100) --------------------------
'
' The footer is Save (cmdOK, relabelled) / Save & Run (cmdSaveRun) / Cancel.
' Both Save buttons run CommitDefaults - normalize, block empty column lists,
' apply the staged logo, write all four keys - and on success hide + confirm
' in the status bar. Save & Run additionally hands modMain_SetTPDDefaults the
' PERF_* flow for the open tab; RunSetTPDDefaults runs it after the form
' unloads, so the perf wrapper toggles screen updating with no modal form in
' memory. Save & Run is hidden on the Logo tab (nothing to run there).

Private Sub cmdOK_Click()               ' the "Save" button
    If Not CommitDefaults() Then Exit Sub
    Me.Hide
    ShowSavedInStatusBar
End Sub

Private Sub cmdSaveRun_Click()
    If Not CommitDefaults() Then Exit Sub
    modMain_SetTPDDefaults.gPendingDefaultsFlow = PerfConstForActivePage()
    Me.Hide
    ShowSavedInStatusBar
End Sub

' True only when every field normalized, no column list was left empty, the
' staged logo applied, and all four keys saved. On False it has already shown
' the reason and the caller leaves the form open (spec 7).
Private Function CommitDefaults() As Boolean
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
        Exit Function
    End If

    ' Apply a staged logo change first - a bad image must not leave columns
    ' saved with no logo (#95).
    logoErr = ApplyPendingLogo()
    If Len(logoErr) > 0 Then
        MsgBox logoErr, vbExclamation, "TPD Add-in"
        Exit Function
    End If

    If Not SaveAllDefaults(eqList, scheduleCols, splitCols, groupColumn) Then
        MsgBox "Couldn't save your defaults. Your changes are still open behind " & _
               "this message - choose OK to return to the form and try again.", _
               vbCritical, "TPD Add-in"
        Exit Function
    End If

    SaveEqListToggles          ' "1"/"0" - best-effort, like the last-tab pref
    CommitDefaults = True
End Function

' The one-click PERF_* flow for the tab that's open, or "" for the Logo tab
' (Save & Run is hidden there, so this is belt-and-braces).
Private Function PerfConstForActivePage() As String
    Select Case mpgPages.SelectedItem.name
        Case "pgEqList":      PerfConstForActivePage = PERF_CREATE_CUST_EQ_LIST_DEFAULTS
        Case "pgSchedule":    PerfConstForActivePage = PERF_CREATE_CUST_SCHEDULE_DEFAULTS
        Case "pgSplitSheets": PerfConstForActivePage = PERF_SPLIT_SHEET_BY_COLUMN_DEFAULTS
    End Select
End Function

' Save & Run has nothing to run on the Logo tab - hide it there entirely
' rather than disable it in place (#100). Called on load and on tab change.
Private Sub SyncSaveRunVisibility()
    cmdSaveRun.Visible = (mpgPages.SelectedItem.name <> "pgLogo")
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
