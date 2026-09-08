Attribute VB_Name = "modHelpers_CheckboxSelection"
'@Folder("TPD_Addin.Helpers")

Option Explicit

'===========================================================
'  Returns a Collection of all selected checkbox captions
'===========================================================
Public Function GetSelectedColumns(fra As MSForms.Frame) As Collection
    Dim result As New Collection
    Dim ctrl As control

    For Each ctrl In fra.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.value = True Then
                result.Add ctrl.Caption
            End If
        End If
    Next ctrl

    Set GetSelectedColumns = result
End Function

'===========================================================
'  True if at least one checkbox in the frame is checked
'===========================================================
Public Function HasColumnSelection(fra As MSForms.Frame) As Boolean
    HasColumnSelection = (GetSelectedColumns(fra).Count > 0)
End Function

'===========================================================
'  Check exactly the checkboxes whose caption is in `resolved`
'  (a column list already run through
'  modPreferences.ResolveColumnList, so the LastUsed* ->
'  DefaultUser* -> shipped fallback chain is applied), and
'  clear the rest. Call after the checkboxes are built. All
'  three column pickers share this (#81, #96).
'===========================================================
Public Sub ApplyColumnSelection(fra As MSForms.Frame, resolved As Collection)
    Dim ctrl As control
    Dim v As Variant

    If resolved Is Nothing Then Exit Sub

    For Each ctrl In fra.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = False

            For Each v In resolved
                If StrComp(ctrl.Caption, CStr(v), vbTextCompare) = 0 Then
                    ctrl.value = True
                    Exit For
                End If
            Next v
        End If
    Next ctrl
End Sub

'===========================================================
'  Tick (isChecked = True) or clear every checkbox in the
'  frame - the Select all / Select none actions on the three
'  column pickers (#149). Walks fra.Controls, so it covers
'  boxes scrolled out of view as well as visible ones. No
'  persistence: cmdOK_Click still writes the LastUsed* key.
'===========================================================
Public Sub SetAllColumns(fra As MSForms.Frame, ByVal isChecked As Boolean)
    Dim ctrl As control

    For Each ctrl In fra.Controls
        If TypeName(ctrl) = "CheckBox" Then ctrl.value = isChecked
    Next ctrl
End Sub

'===========================================================
'  "Restore defaults" on a picker (#149): reset the ticks to
'  the user's configured default list (a DefaultUser* key,
'  written only by Set TPD Defaults), falling back to the
'  shipped list. Deliberately skips the picker's own LastUsed*
'  key via a single-key ResolveColumnList call, so the button
'  means the same thing it does on the Set Defaults tabs
'  (#149 option 1). Caller passes its own DefaultUser* key and
'  shipped CSV (e.g. PREF_EQLIST_COLUMNS, DefaultEqListColumns()).
'===========================================================
Public Sub RestoreColumnDefaults(fra As MSForms.Frame, _
                                 ByVal defaultUserKey As String, _
                                 ByVal shippedCsv As String)
    ApplyColumnSelection fra, ResolveColumnList(Array(defaultUserKey), shippedCsv)
End Sub

