Attribute VB_Name = "modPreferences_Defaults"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
'  What the add-in ships as, before a user saves anything:
'  the default column lists below, and the EQ List behaviour
'  toggle table at the bottom of this module.
'
'  These seed the Set Defaults dialog on first run, are what
'  its per-page "Restore defaults" buttons return to, and are
'  the fallback when a DefaultUser* key has never been saved
'  (LoadPref returns the value passed here).
'
'  Canonical form of a column list is comma-separated, matching
'  what the dialog stores and shows (see
'  modPreferences.NormalizeColumnList).
'===========================================================

Public Function DefaultEqListColumns() As String
    DefaultEqListColumns = "INTERNAL ID, Location, TYPE1, TYPE2, SIZE, RATING, " & _
                           "CONNECTION TYPE, Description, Manufacturer, Model, Details"
End Function

Public Function DefaultScheduleColumns() As String
    DefaultScheduleColumns = "Status, % Complete, Tasks, Start Date, End Date"
End Function

Public Function DefaultSplitColumns() As String
    DefaultSplitColumns = DefaultEqListColumns()
End Function

Public Function DefaultSplitGroupColumn() As String
    DefaultSplitGroupColumn = "Vendor"
End Function


'===========================================================
'  EQ List behaviour toggles (#97 / #119-#123)
'
'  One table for the five checkboxes on the Set Defaults EQ
'  List page. Each row carries everything about a toggle
'  except the code that acts on it:
'
'      key             the PREF_* constant it is stored under
'      shipsOn         its value on a machine that has never
'                      saved defaults - all off bar "add cell
'                      borders", which keeps the pre-#122 look
'      displayName     how it is named in a message to the user
'      needsPurchased  True when the toggle can only run if the
'                      source sheet has a PURCHASED column
'
'  Those four facts used to be hand-listed across four
'  routines - the dialog's Load / Save / Restore and the EQ
'  List flow's skipped-toggle notice - free to disagree with
'  each other. Adding a toggle now means a PREF_* constant, a
'  row here, a checkbox, and one line each in
'  LoadEqListToggles / SaveEqListToggles: the two places a
'  control has to be named out loud, kept explicit so the
'  compiler still checks the name.
'===========================================================

' One row per toggle, as "key|shipsOn|display name|needsPurchased" with 1/0
' for the two flags - a flat array of delimited strings rather than an array
' of arrays, which is how the add-in already stores its structured settings
' (column lists are comma-separated CSV). An earlier array-of-arrays version
' of this table raised an unhandled error in the lookup below; the cause was
' never pinned down, so this shape is the deliberate one - do not "simplify"
' it back to nested Array() calls without testing the lookup in Excel.
Private Const TOGGLE_SEP As String = "|"

Private Const TOGGLE_KEY As Long = 0
Private Const TOGGLE_SHIPS_ON As Long = 1
Private Const TOGGLE_NAME As Long = 2
Private Const TOGGLE_NEEDS_PURCHASED As Long = 3

Private Function EqListToggleTable() As Variant
    EqListToggleTable = Array( _
        PREF_EQLIST_REMOVE_PARENT_ROWS & "|0|Remove PARENT rows|1", _
        PREF_EQLIST_ADD_COUNT_COLUMN & "|0|Add EQ Count column|1", _
        PREF_EQLIST_PLAIN_PARENT_ROWS & "|0|Plain PARENT rows|1", _
        PREF_EQLIST_ADD_CELL_BORDERS & "|1|Add cell borders|0", _
        PREF_EQLIST_ADD_COLUMN_FILTERS & "|0|Add column filters|0")
End Function

' Every toggle key, in the order the checkboxes appear on the page - so a
' caller can walk the toggles without knowing which ones exist.
Public Function EqListToggleKeys() As Variant
    Dim toggleRows As Variant
    Dim keys() As String
    Dim i As Long

    toggleRows = EqListToggleTable()
    ReDim keys(LBound(toggleRows) To UBound(toggleRows))
    For i = LBound(toggleRows) To UBound(toggleRows)
        keys(i) = Split(CStr(toggleRows(i)), TOGGLE_SEP)(TOGGLE_KEY)
    Next i

    EqListToggleKeys = keys
End Function

' What this toggle is set to on a machine that has never saved defaults.
Public Function EqListToggleShipsOn(ByVal key As String) As Boolean
    EqListToggleShipsOn = (ToggleField(key, TOGGLE_SHIPS_ON, "0") = "1")
End Function

' How this toggle is named when a message has to mention it.
Public Function EqListToggleName(ByVal key As String) As String
    EqListToggleName = ToggleField(key, TOGGLE_NAME, vbNullString)
End Function

' True when this toggle needs a PURCHASED column on the source sheet.
Public Function EqListToggleNeedsPurchased(ByVal key As String) As Boolean
    EqListToggleNeedsPurchased = (ToggleField(key, TOGGLE_NEEDS_PURCHASED, "0") = "1")
End Function

' The saved value of one toggle, falling back to what it ships as. Every
' reader goes through here rather than LoadToggle directly, so no caller
' has to repeat "this one ships on".
Public Function LoadEqListToggle(ByVal key As String) As Boolean
    LoadEqListToggle = LoadToggle(key, EqListToggleShipsOn(key))
End Function

' One field of the row for key, or fallback if key isn't in the table.
Private Function ToggleField(ByVal key As String, ByVal fieldIndex As Long, _
                             ByVal fallback As String) As String
    Dim toggleRows As Variant
    Dim fields() As String
    Dim i As Long

    toggleRows = EqListToggleTable()
    For i = LBound(toggleRows) To UBound(toggleRows)
        fields = Split(CStr(toggleRows(i)), TOGGLE_SEP)
        If StrComp(fields(TOGGLE_KEY), key, vbTextCompare) = 0 Then
            ToggleField = fields(fieldIndex)
            Exit Function
        End If
    Next i

    ToggleField = fallback
End Function
