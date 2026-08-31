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
'  Auto-select checkboxes based on saved column list
'===========================================================
Public Sub AutoSelectColumns(fra As MSForms.Frame, savedCols As Collection)
    Dim ctrl As control
    Dim v As Variant

    If savedCols Is Nothing Then Exit Sub

    For Each ctrl In fra.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = False

            For Each v In savedCols
                If StrComp(ctrl.Caption, CStr(v), vbTextCompare) = 0 Then
                    ctrl.value = True
                    Exit For
                End If
            Next v
        End If
    Next ctrl
End Sub

