Attribute VB_Name = "modPreferences_Initializer"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
'  Preferences Initializer
'
'  Preference values no longer need seeding - LoadPref returns
'  the caller's default for any key that was never saved, and
'  each picker falls back to its own built-in column list. All
'  this does now is record which add-in version last ran for
'  the current user (support / diagnostics).
'===========================================================

Public Sub InitializePreferences()
    If LoadPref(PREF_VERSION, "") <> ADDIN_VERSION Then
        SavePref PREF_VERSION, ADDIN_VERSION
    End If
End Sub
