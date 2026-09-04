Attribute VB_Name = "modStartup"
'@Folder("TPD_Addin.Core")

'===========================================================
'  Add-in startup sequencing. InitializeAddIn is called from
'  modRibbonCallbacks.RibbonOnLoad; it records the running
'  add-in version for the current user via modPreferences
'  (registry-backed key/value store). The _Resources sheet +
'  embedded logo shape come from the base .xlam, not from code.
'===========================================================

Option Explicit

' The add-in's own version. Bump this on every release (see CONTRIBUTING.md);
' modPreferences_Initializer stamps it into the registry as PREF_VERSION.
Public Const ADDIN_VERSION As String = "2.3.1"

Public Sub InitializeAddIn()
    InitializePreferences

    ' Any other startup tasks go here
End Sub
