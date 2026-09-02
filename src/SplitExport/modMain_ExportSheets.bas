Attribute VB_Name = "modMain_ExportSheets"
'@Folder("TPD_Addin.SplitExport")

Option Explicit

' Save Each Sheet to XLSX never exports hidden sheets. Flip this to True (in
' source) if a build ever needs plain-hidden sheets included; the add-in's
' internal "_" sheets and very-hidden sheets are excluded regardless.
Private Const EXPORT_INCLUDE_HIDDEN_SHEETS As Boolean = False

Public Sub ExportSheets(control As IRibbonControl)
    WithPerformance "ExportSheets_Internal"
End Sub


Public Sub ExportSheets_Internal()

    Dim wb As Workbook
    Dim frm As frmFilenameOptions
    Dim suffix As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook available.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------------
    ' Show the filename options form
    '---------------------------------------------
    Set frm = New frmFilenameOptions
    frm.Show

    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If

    suffix = frm.UserAppend

    If frm.IncludeDate Then
        If Len(suffix) > 0 Then
            suffix = suffix & " - " & Format(Date, "mm-dd-yyyy")
        Else
            suffix = Format(Date, "mm-dd-yyyy")
        End If
    End If

    Unload frm

    '---------------------------------------------
    ' Call worker
    '---------------------------------------------
    ExportSheets_DoWork wb, suffix

End Sub

Public Sub ExportSheets_DoWork(wb As Workbook, suffix As String)

    Dim exportFolder As String
    Dim ws As Worksheet
    Dim exportCount As Long
    Dim failures As Collection
    Dim result As String

    ' Ensure export folder exists
    exportFolder = EnsureExportFolder(wb)
    If Len(exportFolder) = 0 Then Exit Sub

    '---------------------------------------------
    ' Export each eligible sheet
    '---------------------------------------------
    Set failures = New Collection

    For Each ws In wb.Worksheets
        If ShouldExportSheet(ws) Then
            result = ExportSheetToXLSX(ws, exportFolder, suffix)
            If Len(result) = 0 Then
                exportCount = exportCount + 1
            Else
                failures.Add result
            End If
        End If
    Next ws

    ReportExportOutcome exportCount, failures, exportFolder

End Sub

' True for sheets that should be written out. Anything whose name starts with
' "_" (the add-in's internal sheets) and any very-hidden sheet is always
' excluded; plain-hidden sheets are excluded unless EXPORT_INCLUDE_HIDDEN_SHEETS
' has been turned on in source.
Private Function ShouldExportSheet(ws As Worksheet) As Boolean
    If Left$(ws.name, 1) = "_" Then Exit Function

    Select Case ws.Visible
        Case xlSheetVisible
            ShouldExportSheet = True
        Case xlSheetHidden
            ShouldExportSheet = EXPORT_INCLUDE_HIDDEN_SHEETS
        Case Else                       ' xlSheetVeryHidden
            ShouldExportSheet = False
    End Select
End Function

Private Sub ReportExportOutcome(ByVal exportCount As Long, ByVal failures As Collection, ByVal exportFolder As String)
    Dim msg As String
    Dim f As Variant

    msg = exportCount & " sheet(s) exported to:" & vbCrLf & exportFolder

    If failures.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & failures.Count & " sheet(s) could not be exported:"
        For Each f In failures
            msg = msg & vbCrLf & "  " & f
        Next f
        MsgBox msg, vbExclamation, "Export finished with errors"
    Else
        MsgBox msg, vbInformation, "Export complete"
    End If
End Sub
