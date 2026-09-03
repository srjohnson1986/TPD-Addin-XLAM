Attribute VB_Name = "modStartup"
'@Folder("TPD_Addin.Core")

'===========================================================
'  Add-in startup sequencing. InitializeAddIn is called from
'  modRibbonCallbacks.RibbonOnLoad and makes sure the hidden
'  _Resources (embedded logo) and _Preferences (saved settings)
'  sheets both exist before any ribbon command runs.
'===========================================================

Option Explicit

Public Sub InitializeAddIn()
    ' Create embedded resources if needed
    CreateResourcesSheetIfMissing

    ' Initialize preferences (GetPrefSheet creates _Preferences on first call)
    GetPrefSheet
    InitializePreferences

    ' Any other startup tasks go here
End Sub

