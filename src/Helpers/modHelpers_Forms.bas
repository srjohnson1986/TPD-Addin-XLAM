Attribute VB_Name = "modHelpers_Forms"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Bridge functions that show a picker UserForm and hand the
'  result back to the caller.
'
'  NOTE: currently unused - every flow instantiates its form
'  directly. Kept for the moment; slated for cleanup with the
'  one-click EQ List refactor (#25).
'===========================================================

Option Explicit

Public Function ShowCustEQForm(headings As Variant, _
                               ByRef selectedCols As Collection, _
                               ByRef imgPath As String) As Boolean
    Dim frm As CustEQListColumnPickerForm

    Set frm = New CustEQListColumnPickerForm
    frm.LoadColumns headings
    frm.Show

    If frm.Cancelled Then
        ShowCustEQForm = False
    Else
        Set selectedCols = frm.GetSelectedColumns
        imgPath = frm.txtImgPath.Text
        ShowCustEQForm = (selectedCols.Count > 0)
    End If

    Unload frm
End Function

Public Function GetExportSuffixFromForm() As String
    Dim frm As frmFilenameOptions
    Dim txt As String
    Dim today As String

    Set frm = New frmFilenameOptions
    frm.Show

    If frm.Cancelled Then
        GetExportSuffixFromForm = "#CANCELLED#"
    Else
        txt = SanitizeFileText(frm.UserAppend)
        If frm.IncludeDate Then
            today = Format(Date, "mm-dd-yyyy")
            If txt <> "" Then
                txt = txt & " - " & today
            Else
                txt = today
            End If
        End If
        GetExportSuffixFromForm = txt
    End If

    Unload frm
End Function

Public Function ShowSplitSheetForm(headings As Variant, _
                                   ByRef groupColName As String, _
                                   ByRef selectedCols As Collection) As Boolean
    Dim frm As splitSheetByColumnOptionsForm

    Set frm = New splitSheetByColumnOptionsForm
    ' if needed, you can pass headings into the form here
    frm.Show

    If frm.groupColumn = "" Or frm.SelectedColumns Is Nothing Then
        ShowSplitSheetForm = False
    Else
        groupColName = frm.groupColumn
        Set selectedCols = frm.SelectedColumns
        ShowSplitSheetForm = True
    End If

    Unload frm
End Function


