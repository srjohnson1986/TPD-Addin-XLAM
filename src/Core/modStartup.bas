Attribute VB_Name = "modStartup"
'@Folder("TPD_Addin.Core")

Public Sub InitializeAddIn()
    ' Create embedded resources if needed
    CreateResourcesSheetIfMissing

    ' Initialize preferences
    EnsurePreferenceSheet
    InitializePreferences

    ' Any other startup tasks go here
End Sub

