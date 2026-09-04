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

' ---- Export Sheets ----
Public Const PREF_EXPORT_APPEND As String = "Export_AppendText"
Public Const PREF_EXPORT_INCLUDEDATE As String = "Export_IncludeDate"

' ---- Versioning (optional) ----
Public Const PREF_VERSION As String = "AddIn_Version"

