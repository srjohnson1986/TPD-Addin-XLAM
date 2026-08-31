Attribute VB_Name = "modPreferences_Initializer"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
' Preferences Initializer
' Ensures the _Preferences sheet exists and required keys exist
'===========================================================

Public Sub InitializePreferences()

    Dim ws As Worksheet
    Dim v As String

    ' Ensure sheet exists
    Set ws = GetPrefSheet()

    ' ---- Set defaults if missing ----

    If LoadPref(PREF_CUSTEQ_IMAGE, "") = "" Then
        SavePref PREF_CUSTEQ_IMAGE, ""
    End If

    If LoadPref(PREF_SPLIT_GROUPCOL, "") = "" Then
        SavePref PREF_SPLIT_GROUPCOL, ""
    End If

    If LoadPref(PREF_EXPORT_APPEND, "") = "" Then
        SavePref PREF_EXPORT_APPEND, ""
    End If

    If LoadPref(PREF_EXPORT_INCLUDEDATE, "") = "" Then
        SavePref PREF_EXPORT_INCLUDEDATE, "0"
    End If

    ' ---- Versioning (optional) ----
    If LoadPref(PREF_VERSION, "") = "" Then
        SavePref PREF_VERSION, "1.0"
    End If

End Sub

