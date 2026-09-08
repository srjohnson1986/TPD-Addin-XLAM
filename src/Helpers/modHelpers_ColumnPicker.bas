Attribute VB_Name = "modHelpers_ColumnPicker"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  The column-picker behaviour shared by the three picker
'  dialogs - frmCustEQListColumnPicker, frmCustScheduleColumnPicker
'  and frmSplitSheetOptions. They differ only in their control
'  name prefix, their two preference keys and their shipped
'  default list, so building the grid and validating / saving
'  the ticked set live here once instead of three times over.
'
'  What stays on each form: its controls (.frx), its event
'  handlers - a standard module can't wire an event - and
'  anything genuinely its own, such as frmSplitSheetOptions'
'  group-column dropdown.
'===========================================================

Option Explicit

'-----------------------------------------------------------
' Builds the checkbox grid for headingList and ticks the
' columns the picker should open with: its own last-used
' selection (pickerKey), else the Set Defaults value
' (defaultKey), else the shipped list (#96, #98).
'-----------------------------------------------------------
Public Sub InitColumnPicker(ByVal fra As MSForms.Frame, _
                            ByVal headingList As Variant, _
                            ByVal rowsPerColumn As Long, _
                            ByVal baseName As String, _
                            ByVal pickerKey As String, _
                            ByVal defaultKey As String, _
                            ByVal shippedCsv As String)

    LayoutCheckboxes fra, headingList, rowsPerColumn, baseName
    ApplyColumnSelection fra, ResolveColumnList(Array(pickerKey, defaultKey), shippedCsv)
End Sub

'-----------------------------------------------------------
' The OK-button guard and save. Returns False - having already
' shown the reason - when nothing is ticked; otherwise writes
' the ticked set to the picker's own LastUsed* key and returns
' True. The one-click DefaultUser* key is never touched here:
' that one belongs to Set Defaults alone (#96).
'-----------------------------------------------------------
Public Function CommitColumnPicker(ByVal fra As MSForms.Frame, _
                                   ByVal pickerKey As String) As Boolean
    If Not HasColumnSelection(fra) Then
        MsgBox "Please select at least one column to keep.", vbExclamation
        Exit Function
    End If

    SaveColumnList pickerKey, GetSelectedColumns(fra)
    CommitColumnPicker = True
End Function
