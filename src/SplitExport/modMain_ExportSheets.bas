Attribute VB_Name = "modMain_ExportSheets"
'@Folder("TPD_Addin.SplitExport")

Option Explicit

Public Sub ExportSheets(control As IRibbonControl)
    WithPerformance "ExportSheets_Internal"
End Sub


Public Sub ExportSheets_Internal()

    Dim wb As Workbook
    Dim frm As frmFilenameOptions
    Dim suffix As String
    Dim exportFolder As String

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

    'MsgBox "Sheets exported successfully.", vbInformation
    'MsgBox exportCount & " sheets exported successfully to:" & vbCrLf & exportFolder, vbInformation


End Sub

Public Sub ExportSheets_DoWork(wb As Workbook, suffix As String)

    Dim exportFolder As String
    Dim ws As Worksheet

    ' Ensure export folder exists
    exportFolder = EnsureExportFolder(wb)
    If Len(exportFolder) = 0 Then Exit Sub

    '---------------------------------------------
    ' Export each sheet
    '---------------------------------------------
    Dim exportCount As Long
    exportCount = 0
    
    For Each ws In wb.Worksheets
        ExportSheetToXLSX ws, exportFolder, suffix
        exportCount = exportCount + 1
    Next ws
    
    MsgBox exportCount & " sheets exported successfully to:" & vbCrLf & exportFolder, vbInformation


End Sub

