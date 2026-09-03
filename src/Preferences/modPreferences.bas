Attribute VB_Name = "modPreferences"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
'  Preference Storage Module
'  Stores key/value settings inside a hidden sheet
'  named "_Preferences" within the add-in.
'
'  This is the portable, cross-platform way
'  to store persistent settings for an Excel add-in.
'===========================================================

Private Const PREF_SHEET As String = "_Preferences"


'-----------------------------------------------------------
' Ensure the preference sheet exists and return it
'-----------------------------------------------------------
Private Function PrefSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PREF_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create the sheet if missing
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = PREF_SHEET
        ws.Visible = xlSheetVeryHidden

        ' Add headings
        ws.Range("A1").value = "Key"
        ws.Range("B1").value = "Value"
    End If

    Set PrefSheet = ws
End Function


'-----------------------------------------------------------
' Save a preference (string value)
'-----------------------------------------------------------
Public Sub SavePref(key As String, value As String)
    Dim ws As Worksheet
    Dim f As Range

    Set ws = PrefSheet()

    ' Look for existing key
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)

    If f Is Nothing Then
        ' Append new key/value
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(nextRow, 1).value = key
        ws.Cells(nextRow, 2).value = value
    Else
        ' Update existing value
        f.Offset(0, 1).value = value
    End If
End Sub


'-----------------------------------------------------------
' Load a preference (returns defaultValue if missing)
'-----------------------------------------------------------
Public Function LoadPref(key As String, Optional defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim f As Range

    Set ws = PrefSheet()

    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)

    If f Is Nothing Then
        LoadPref = defaultValue
    Else
        LoadPref = CStr(f.Offset(0, 1).value)
    End If
End Function


'-----------------------------------------------------------
' Delete a preference (optional utility)
'-----------------------------------------------------------
Public Sub DeletePref(key As String)
    Dim ws As Worksheet
    Dim f As Range

    Set ws = PrefSheet()

    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)

    If Not f Is Nothing Then
        f.EntireRow.Delete
    End If
End Sub

' Save a collection of column names as CSV
Public Sub SaveColumnList(key As String, cols As Collection)
    Dim v As Variant
    Dim parts() As String
    Dim i As Long

    ReDim parts(1 To cols.Count)
    i = 1

    For Each v In cols
        parts(i) = CStr(v)
        i = i + 1
    Next v

    SavePref key, Join(parts, ",")
End Sub

' Load a CSV list of column names into a collection
Public Function LoadColumnList(key As String) As Collection
    Dim raw As String
    Dim parts As Variant
    Dim col As New Collection
    Dim p As Variant

    raw = LoadPref(key, "")

    If Len(raw) = 0 Then
        Set LoadColumnList = col
        Exit Function
    End If

    parts = Split(raw, ",")

    For Each p In parts
        col.Add Trim(CStr(p))
    Next p

    Set LoadColumnList = col
End Function


'-----------------------------------------------------------
' Clear ALL preferences (useful for debugging)
'-----------------------------------------------------------
Public Sub ClearAllPrefs()
    Dim ws As Worksheet
    Set ws = PrefSheet()

    ws.Rows("2:" & ws.Rows.Count).ClearContents
End Sub

' Single public accessor for the _Preferences sheet. Callers that only need
' to guarantee the sheet exists can ignore the return value; calling this
' creates _Preferences (very hidden, with Key/Value headings) on first use.
Public Function GetPrefSheet() As Worksheet
    Set GetPrefSheet = PrefSheet()   ' PrefSheet stays private to this module
End Function

