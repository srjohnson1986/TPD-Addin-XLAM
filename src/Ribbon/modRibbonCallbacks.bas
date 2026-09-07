Attribute VB_Name = "modRibbonCallbacks"
'@Folder("TPD_Addin.Ribbon")

'===========================================================
'  Bridges customUI14.xml onAction callbacks to the real
'  entry points, owns the global gRibbon reference, and runs
'  RibbonOnLoad -> modStartup.InitializeAddIn.
'===========================================================

Option Explicit

Public gRibbon As IRibbonUI


Public Sub RunCreateCustEQList(control As IRibbonControl)
    CreateCustEQList control
End Sub

Public Sub RunSplitSheetByColumn(control As IRibbonControl)
    SplitSheetByColumn control
End Sub

Public Sub RunCreateCustSchedule(control As IRibbonControl)
    CreateCustSchedule control
End Sub

' ---- Defaults group: one-click runs (no picker) --------------------------

Public Sub RunQuickEqList(control As IRibbonControl)
    CreateCustEQListFromDefaults control
End Sub

Public Sub RunQuickSplitSheets(control As IRibbonControl)
    SplitSheetByColumnFromDefaults control
End Sub

Public Sub RunQuickSchedule(control As IRibbonControl)
    CreateCustScheduleFromDefaults control
End Sub

Public Sub RunExportSheetsToXLSX(control As IRibbonControl)
    ExportSheets control
End Sub

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    InitializeAddIn
End Sub
