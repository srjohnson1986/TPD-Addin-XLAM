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

Public Sub RunExportSheetsToXLSX(control As IRibbonControl)
    ExportSheets control
End Sub

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    InitializeAddIn
End Sub
