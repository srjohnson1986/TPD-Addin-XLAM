Attribute VB_Name = "modPerformance"
'@Folder("TPD_Addin.Core")

Option Explicit

Public Sub WithPerformance(action As String)
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
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

