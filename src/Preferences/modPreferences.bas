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
' Which registry keys the Set TPD Defaults dialog owns. One
' list so SaveAllDefaults and DefaultsNeverSaved can't drift
' apart; the order here is the order SaveAllDefaults passes
' its values in.
'-----------------------------------------------------------
Private Function DefaultsPrefKeys() As Variant
    DefaultsPrefKeys = Array(PREF_EQLIST_COLUMNS, PREF_SCHEDULE_COLUMNS, _
                             PREF_SPLIT_COLUMNS, PREF_SPLIT_GROUPCOL)
End Function


'-----------------------------------------------------------
' Write several preferences in one call. keys and values are
' equal-length arrays (as returned by Array()); each pair is
' written and the function returns False if any single write
' failed, so the caller can react (frmSetTPDDefaults keeps its
' form open on False so the user can retry, which rewrites the
' whole set). No rollback: a half-written set on the very rare
' HKCU write failure is corrected by the next successful save,
' and the rollback it replaced used the same SaveSetting that
' had just failed anyway.
'-----------------------------------------------------------
Public Function SavePrefs(keys As Variant, values As Variant) As Boolean
    Dim i As Long

    If LBound(keys) <> LBound(values) Or UBound(keys) <> UBound(values) Then
        SavePrefs = False
        Exit Function
    End If

    SavePrefs = True
    On Error Resume Next
    For i = LBound(keys) To UBound(keys)
        Err.Clear
        SaveSetting PREF_APP, PREF_SECTION, CStr(keys(i)), CStr(values(i))
        If Err.Number <> 0 Then SavePrefs = False
    Next i
    On Error GoTo 0
End Function


'-----------------------------------------------------------
' True when every one of the given keys reads back empty for
' the current user - i.e. none has ever been saved. A value
' explicitly saved as "" also reads as empty here; that
' ambiguity is tracked in issue #104.
'-----------------------------------------------------------
Public Function AllPrefsUnset(keys As Variant) As Boolean
    Dim i As Long

    AllPrefsUnset = True
    For i = LBound(keys) To UBound(keys)
        If Len(LoadPref(CStr(keys(i)), "")) > 0 Then
            AllPrefsUnset = False
            Exit Function
        End If
    Next i
End Function


'-----------------------------------------------------------
' True when none of the dialog's keys has ever been saved for
' the current user - drives the first-run notice. (Can't test
' "any pref exists": modPreferences_Initializer always stamps
' PREF_VERSION.)
'-----------------------------------------------------------
Public Function DefaultsNeverSaved() As Boolean
    DefaultsNeverSaved = AllPrefsUnset(DefaultsPrefKeys())
End Function


'-----------------------------------------------------------
' Writes all four Set TPD Defaults keys. Returns True only if
' every write succeeded; on False the caller (frmSetTPDDefaults)
' keeps its form open so the user's edits aren't lost and a
' retry rewrites the whole set. No rollback - see SavePrefs.
'-----------------------------------------------------------
Public Function SaveAllDefaults(ByVal eqListColumns As String, _
                                ByVal scheduleColumns As String, _
                                ByVal splitColumns As String, _
                                ByVal splitGroupColumn As String) As Boolean
    SaveAllDefaults = SavePrefs(DefaultsPrefKeys(), _
        Array(eqListColumns, scheduleColumns, splitColumns, splitGroupColumn))
End Function
