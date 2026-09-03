Attribute VB_Name = "modStartup"
'@Folder("TPD_Addin.Core")

'===========================================================
'  Add-in startup sequencing. InitializeAddIn is called from
'  modRibbonCallbacks.RibbonOnLoad and makes sure the hidden
'  _Preferences (saved settings) sheet exists before any ribbon
'  command runs. The _Resources sheet + embedded logo shape
'  come from the base .xlam, not from code.
'===========================================================

Option Explicit

' The add-in's own version. Bump this on every release (see CONTRIBUTING.md);
' modPreferences_Initializer stamps it into _Preferences as PREF_VERSION.
Public Const ADDIN_VERSION As String = "2.3.0"

Public Sub InitializeAddIn()
    ' GetPrefSheet creates _Preferences on first call
    GetPrefSheet
    InitializePreferences

    ' Any other startup tasks go here
End Sub

