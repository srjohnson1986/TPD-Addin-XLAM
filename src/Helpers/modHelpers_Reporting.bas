Attribute VB_Name = "modHelpers_Reporting"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  The end-of-run summary both batch flows show - "Split Sheet
'  by Column" and "Save Each Sheet to XLSX". It lived in
'  modHelpers_Export, which is the one module the Split flow
'  had to reach into despite exporting nothing; a shared
'  routine belongs in Helpers with the other shared code.
'===========================================================

Option Explicit

' Shows the end-of-run summary for a batch flow. "summary" is the lead line
' (already complete - e.g. "3 sheet(s) created." or "3 sheet(s) exported to:"
' + path). Each entry in "failures" is a "'name': reason" string; when there
' are any, they are listed under "<count> <failureNoun>" and the box is a
' warning titled errTitle, otherwise it's an info box titled okTitle.
Public Sub ReportBatchOutcome(ByVal summary As String, ByVal failures As Collection, _
                              ByVal failureNoun As String, _
                              ByVal errTitle As String, ByVal okTitle As String)
    Dim msg As String
    Dim f As Variant

    msg = summary

    If failures.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & failures.Count & " " & failureNoun
        For Each f In failures
            msg = msg & vbCrLf & "  " & f
        Next f
        MsgBox msg, vbExclamation, errTitle
    Else
        MsgBox msg, vbInformation, okTitle
    End If
End Sub
