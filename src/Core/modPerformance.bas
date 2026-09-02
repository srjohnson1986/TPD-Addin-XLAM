Attribute VB_Name = "modPerformance"
'@Folder("TPD_Addin.Core")

Option Explicit

Public Sub WithPerformance(action As String)

    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Select Case action

        Case "CreateCustEQList_Internal"
            CreateCustEQList_Internal

        Case "SplitSheetByColumn_Internal"
            SplitSheetByColumn_Internal

        Case "ExportSheets_Internal"
            ExportSheets_Internal

        ' Add more internal functions here as needed

        Case Else
            MsgBox "Unknown action: " & action, vbCritical

    End Select

CleanUp:
    ' Capture the error (if any) before any statement below clears Err.
    errNumber = Err.Number
    errDescription = Err.Description

    ' Always restore session-wide Excel settings, error or not.
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If errNumber <> 0 Then
        MsgBox "'" & action & "' did not finish." & vbCrLf & vbCrLf & _
               errDescription & " (error " & errNumber & ")", _
               vbExclamation, "TPD Add-in"
    End If

End Sub
