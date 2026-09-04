Attribute VB_Name = "modPreferences_KeyMap"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
' Preference Key Map
' Centralized list of all preference keys used by the add-in
'===========================================================

' ---- Customer EQ List ----
Public Const PREF_CUSTEQ_COLUMNS As String = "CustEQ_SelectedColumns"

' ---- Split Sheet By Column ----
Public Const PREF_SPLIT_GROUPCOL As String = "Split_GroupColumn"
Public Const PREF_SPLIT_COLUMNS As String = "Split_SelectedColumns"

' ---- Versioning ----
Public Const PREF_VERSION As String = "AddIn_Version"

' "Save Each Sheet to XLSX" (frmFilenameOptions) does not persist its
' append-text / include-date inputs - it asks every run. If that ever
' becomes desired it should be workbook-scoped (per file), not a global
' user preference, so no PREF_* key lives here for it.

