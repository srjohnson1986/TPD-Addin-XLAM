Attribute VB_Name = "modPreferences_KeyMap"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
' Preference Key Map
' Centralized list of all preference keys used by the add-in
'
' Column defaults are three layers deep:
'
'   1. Shipped defaults - modPreferences_Defaults, no key. The
'      final fallback for both the pickers and the one-click
'      runs.
'   2. User one-click defaults - the DefaultUser* keys below.
'      Written ONLY by the Set TPD Defaults dialog
'      (frmSetTPDDefaults). The one-click ribbon buttons read
'      these, falling back to the shipped defaults.
'   3. User picker preferences - the LastUsed* keys below.
'      Written by the per-run picker forms on OK. The pickers
'      pre-fill from these, falling back to the DefaultUser*
'      key, then the shipped defaults.
'
' modPreferences.ResolveColumnList / ResolveGroupColumn walk a
' caller-supplied key list in that order. No migration: an
' existing DefaultUser* value still drives the one-click run
' and the picker falls through to it until its own LastUsed*
' key is written.
'===========================================================

' ---- Customer EQ List ----
Public Const PREF_EQLIST_COLUMNS As String = "DefaultUserEqListColumns"
Public Const PREF_EQLIST_PICKER_COLUMNS As String = "LastUsedEqListColumns"

' ---- Customer Schedule ----
Public Const PREF_SCHEDULE_COLUMNS As String = "DefaultUserScheduleColumns"

' ---- Split Sheet By Column ----
Public Const PREF_SPLIT_GROUPCOL As String = "DefaultUserSplitSheetsGroupColumn"
Public Const PREF_SPLIT_COLUMNS As String = "DefaultUserSplitSheetsColumns"
Public Const PREF_SPLIT_PICKER_GROUPCOL As String = "LastUsedSplitSheetsGroupColumn"
Public Const PREF_SPLIT_PICKER_COLUMNS As String = "LastUsedSplitSheetsColumns"

' ---- Versioning ----
Public Const PREF_VERSION As String = "AddIn_Version"

' "Save Each Sheet to XLSX" (frmFilenameOptions) does not persist its
' append-text / include-date inputs - it asks every run. If that ever
' becomes desired it should be workbook-scoped (per file), not a global
' user preference, so no PREF_* key lives here for it.
