Attribute VB_Name = "modPreferences"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
'  Preference Storage Module
'
'  Key/value settings for the whole add-in, per Windows user,
'  stored in the registry via VBA's SaveSetting / GetSetting:
'      HKCU\Software\VB and VBA Program Settings\TPD_Addin\Preferences
'
'  Registry, not a sheet inside the .xlam: Excel never saves an
'  add-in on exit, so anything written into the add-in's own
'  workbook is discarded when Excel closes
'  ([#84](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/84)).
'  The registry is always writable (HKCU, no admin prompt), the
'  write lands immediately, and the store is global to the
'  add-in regardless of which workbook is open.
'
'  All keys live in modPreferences_KeyMap as PREF_* constants -
'  never pass a loose string literal.
'
'  Fallback: a key that was never saved on this machine makes
'  LoadPref return the caller's defaultValue and LoadColumnList
'  return an empty Collection; each picker then applies its own
'  built-in default column list. A fresh machine still opens
'  with sensible defaults.
'===========================================================

Private Const PREF_APP As String = "TPD_Addin"
Private Const PREF_SECTION As String = "Preferences"


'-----------------------------------------------------------
' Save one preference. Written to the registry immediately -
' no add-in save needed. A write failure (unusual for HKCU) is
' swallowed rather than thrown mid-flow; the value still holds
' for the rest of the session through the calling form.
'-----------------------------------------------------------
Public Sub SavePref(key As String, value As String)
    On Error Resume Next
    SaveSetting PREF_APP, PREF_SECTION, key, value
    On Error GoTo 0
End Sub


'-----------------------------------------------------------
' Load one preference, or defaultValue when the key has never
' been saved for the current user.
'-----------------------------------------------------------
Public Function LoadPref(key As String, Optional defaultValue As String = "") As String
    On Error Resume Next
    LoadPref = GetSetting(PREF_APP, PREF_SECTION, key, defaultValue)
    On Error GoTo 0
End Function


'-----------------------------------------------------------
' Remove one preference. Not an error if it isn't set.
'-----------------------------------------------------------
Public Sub DeletePref(key As String)
    On Error Resume Next
    DeleteSetting PREF_APP, PREF_SECTION, key
    On Error GoTo 0
End Sub


' Save a Collection of column names as one comma-separated value.
' An empty or Nothing collection is stored as "" (which reads
' back as "no saved list" -> the picker applies its defaults).
Public Sub SaveColumnList(key As String, cols As Collection)
    Dim parts() As String
    Dim v As Variant
    Dim i As Long

    If cols Is Nothing Then SavePref key, "": Exit Sub
    If cols.Count = 0 Then SavePref key, "": Exit Sub

    ReDim parts(0 To cols.Count - 1)
    i = 0
    For Each v In cols
        parts(i) = CStr(v)
        i = i + 1
    Next v

    SavePref key, Join(parts, ",")
End Sub


' Load a comma-separated column list into a Collection. Returns
' an empty Collection (never Nothing) when the key is unset, so
' a caller can read "no saved list" as "apply my defaults".
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
        If Len(Trim$(CStr(p))) > 0 Then col.Add Trim$(CStr(p))
    Next p

    Set LoadColumnList = col
End Function


'-----------------------------------------------------------
' Clear every add-in preference for the current user (debug
' aid). No-op if nothing has been saved yet.
'-----------------------------------------------------------
Public Sub ClearAllPrefs()
    On Error Resume Next
    DeleteSetting PREF_APP, PREF_SECTION
    On Error GoTo 0
End Sub


'===========================================================
'  Set TPD Defaults dialog support
'===========================================================

'-----------------------------------------------------------
' Canonical stored form of a column list is comma-separated.
' Tabs, carriage returns and new lines are all delimiters, so
' a heading row pasted straight out of Excel normalizes
' cleanly. Duplicates are kept; empty entries are dropped.
'-----------------------------------------------------------
Public Function NormalizeColumnList(ByVal rawText As String) As String
    Dim working As String
    Dim parts() As String
    Dim kept As Collection
    Dim entry As String
    Dim i As Long

    working = rawText
    working = Replace$(working, vbCrLf, ",")
    working = Replace$(working, vbCr, ",")
    working = Replace$(working, vbLf, ",")
    working = Replace$(working, vbTab, ",")

    parts = Split(working, ",")
    Set kept = New Collection

    For i = LBound(parts) To UBound(parts)
        entry = CollapseWhitespace(parts(i))
        If Len(entry) > 0 Then kept.Add entry
    Next i

    If kept.Count = 0 Then
        NormalizeColumnList = vbNullString
        Exit Function
    End If

    Dim result() As String
    ReDim result(1 To kept.Count)
    For i = 1 To kept.Count
        result(i) = kept(i)
    Next i

    NormalizeColumnList = Join(result, ", ")
End Function


'-----------------------------------------------------------
' Trims both ends and collapses internal whitespace runs to a
' single space, so a stored / typed name compares cleanly.
'-----------------------------------------------------------
Public Function CollapseWhitespace(ByVal someText As String) As String
    Dim working As String

    working = Replace$(Trim$(someText), vbTab, " ")
    Do While InStr(working, "  ") > 0
        working = Replace$(working, "  ", " ")
    Loop

    CollapseWhitespace = working
End Function


'-----------------------------------------------------------
' True when none of the dialog's four keys has ever been
' saved for the current user - drives the first-run notice.
' (Can't test "any pref exists": modPreferences_Initializer
' always stamps PREF_VERSION.)
'-----------------------------------------------------------
Public Function DefaultsNeverSaved() As Boolean
    DefaultsNeverSaved = _
        Len(GetSetting(PREF_APP, PREF_SECTION, PREF_EQLIST_COLUMNS, vbNullString)) = 0 And _
        Len(GetSetting(PREF_APP, PREF_SECTION, PREF_SCHEDULE_COLUMNS, vbNullString)) = 0 And _
        Len(GetSetting(PREF_APP, PREF_SECTION, PREF_SPLIT_COLUMNS, vbNullString)) = 0 And _
        Len(GetSetting(PREF_APP, PREF_SECTION, PREF_SPLIT_GROUPCOL, vbNullString)) = 0
End Function


'-----------------------------------------------------------
' Writes all four dialog keys together, rolling back to the
' prior values on any failure so the store never holds a
' half-applied set. Returns True on success. Unlike SavePref,
' this surfaces a write failure to the caller.
'-----------------------------------------------------------
Public Function SaveAllDefaults(ByVal eqListColumns As String, _
                                ByVal scheduleColumns As String, _
                                ByVal splitColumns As String, _
                                ByVal splitGroupColumn As String) As Boolean
    Dim priorEqList As String, priorSchedule As String
    Dim priorSplit As String, priorGroup As String

    priorEqList = GetSetting(PREF_APP, PREF_SECTION, PREF_EQLIST_COLUMNS, vbNullString)
    priorSchedule = GetSetting(PREF_APP, PREF_SECTION, PREF_SCHEDULE_COLUMNS, vbNullString)
    priorSplit = GetSetting(PREF_APP, PREF_SECTION, PREF_SPLIT_COLUMNS, vbNullString)
    priorGroup = GetSetting(PREF_APP, PREF_SECTION, PREF_SPLIT_GROUPCOL, vbNullString)

    On Error GoTo WriteFailed
    SaveSetting PREF_APP, PREF_SECTION, PREF_EQLIST_COLUMNS, eqListColumns
    SaveSetting PREF_APP, PREF_SECTION, PREF_SCHEDULE_COLUMNS, scheduleColumns
    SaveSetting PREF_APP, PREF_SECTION, PREF_SPLIT_COLUMNS, splitColumns
    SaveSetting PREF_APP, PREF_SECTION, PREF_SPLIT_GROUPCOL, splitGroupColumn
    On Error GoTo 0

    SaveAllDefaults = True
    Exit Function

WriteFailed:
    On Error Resume Next
    SaveSetting PREF_APP, PREF_SECTION, PREF_EQLIST_COLUMNS, priorEqList
    SaveSetting PREF_APP, PREF_SECTION, PREF_SCHEDULE_COLUMNS, priorSchedule
    SaveSetting PREF_APP, PREF_SECTION, PREF_SPLIT_COLUMNS, priorSplit
    SaveSetting PREF_APP, PREF_SECTION, PREF_SPLIT_GROUPCOL, priorGroup
    On Error GoTo 0
    SaveAllDefaults = False
End Function
