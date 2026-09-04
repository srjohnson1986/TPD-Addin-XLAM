Attribute VB_Name = "modPreferences_Defaults"
'@Folder("TPD_Addin.Preferences")

Option Explicit

'===========================================================
'  Hard-coded default column lists
'
'  These ship with the add-in, seed the Set TPD Defaults
'  dialog on first run, and are what its per-page "Restore
'  defaults" buttons return to. They are also the fallback
'  when a DefaultUser* key has never been saved (LoadPref
'  returns the value passed here).
'
'  Canonical form is comma-separated, matching what the
'  dialog stores and shows (see modPreferences.NormalizeColumnList).
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
