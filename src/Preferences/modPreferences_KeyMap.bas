Attribute VB_Name = "modPreferences_KeyMap"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
' Preference Key Map
' Centralized list of all preference keys used by the add-in
'
' The "DefaultUser*" column keys are the single set shared by
' the Set TPD Defaults dialog (frmSetTPDDefaults) and the
' per-run pickers (CustEQListColumnPickerForm,
' splitSheetByColumnOptionsForm). The dialog sets them; each
' picker pre-fills from them and writes the user's in-the-
' moment choice back to the same key. A missing key falls back
' to modPreferences_Defaults.
'===========================================================

' ---- Customer EQ List ----
Public Const PREF_EQLIST_COLUMNS As String = "DefaultUserEqListColumns"

' ---- Customer Schedule ----
Public Const PREF_SCHEDULE_COLUMNS As String = "DefaultUserScheduleColumns"

' ---- Split Sheet By Column ----
Public Const PREF_SPLIT_GROUPCOL As String = "DefaultUserSplitSheetsGroupColumn"
Public Const PREF_SPLIT_COLUMNS As String = "DefaultUserSplitSheetsColumns"

' ---- Versioning ----
Public Const PREF_VERSION As String = "AddIn_Version"

' "Save Each Sheet to XLSX" (frmFilenameOptions) does not persist its
' append-text / include-date inputs - it asks every run. If that ever
' becomes desired it should be workbook-scoped (per file), not a global
' user preference, so no PREF_* key lives here for it.
