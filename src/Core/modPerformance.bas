Attribute VB_Name = "modPerformance"
'@Folder("TPD_Addin.Core")

'===========================================================
'  WithPerformance runs one internal routine with screen
'  updating and events off and calculation manual, then
'  ALWAYS restores those three session-wide settings - on
'  success and on error alike (GEN-04 / #7). On error it also
'  shows the failure in a message box instead of leaving the
'  add-in silently half-done.
'
'  Callers pass one of the PERF_* constants below rather than
'  a bare string, so a mistyped routine name is a compile
'  error at the call site instead of a run-time "Unknown
'  action". Adding a flow = add a PERF_* constant, a Case in
'  the Select, and the WithPerformance call in the callback.
'===========================================================

Option Explicit

Public Const PERF_CREATE_CUST_EQ_LIST As String = "CreateCustEQList_Internal"
Public Const PERF_SPLIT_SHEET_BY_COLUMN As String = "SplitSheetByColumn_Internal"
Public Const PERF_EXPORT_SHEETS As String = "ExportSheets_Internal"
Public Const PERF_COUNT_EQ_ROWS As String = "CountEquipmentRows_Internal"

Public Sub WithPerformance(action As String)

    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Select Case action

        Case PERF_CREATE_CUST_EQ_LIST
            CreateCustEQList_Internal

        Case PERF_SPLIT_SHEET_BY_COLUMN
            SplitSheetByColumn_Internal

        Case PERF_EXPORT_SHEETS
            ExportSheets_Internal

        Case PERF_COUNT_EQ_ROWS
            CountEquipmentRows_Internal

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
