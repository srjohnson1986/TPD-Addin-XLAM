Attribute VB_Name = "modRibbonCallbacks"
'@Folder("TPD_Addin.Ribbon")

'===========================================================
'  Every customUI14.xml onAction callback, in one module.
'  Each is a one-liner: hand WithPerformance the PERF_*
'  constant naming the flow, and modPerformance runs it with
'  screen updating / events / calculation managed.
'
'  Adding a ribbon button = a PERF_* constant, a Case in
'  modPerformance's Select, and a callback here whose name
'  matches the button's onAction attribute.
'
'  The one exception is Set Defaults: RunSetTPDDefaults lives
'  in modMain_SetTPDDefaults because it is not a forwarder -
'  it shows the dialog, then runs whichever one-click flow the
'  "Save & Run" button staged, after the form has unloaded.
'
'  Also owns the global gRibbon reference and RibbonOnLoad.
'===========================================================

Option Explicit

Public gRibbon As IRibbonUI


' ---- Sheet Tools group: the picker paths ---------------------------------

Public Sub RunCreateCustEQList(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_EQ_LIST
End Sub

Public Sub RunCreateCustSchedule(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_SCHEDULE
End Sub

' Named without the "Run" prefix its neighbours carry: the shipped ribbon's
' onAction says "SplitSheetByColumn", and the ribbon XML lives in the base
' .xlam, so renaming it would mean re-stamping the base file as well as
' editing customUI/customUI14.xml. Left as the ribbon asks for it.
Public Sub SplitSheetByColumn(control As IRibbonControl)
    WithPerformance PERF_SPLIT_SHEET_BY_COLUMN
End Sub

Public Sub RunExportSheetsToXLSX(control As IRibbonControl)
    WithPerformance PERF_EXPORT_SHEETS
End Sub


' ---- One-click group: straight from saved defaults, no picker ------------

Public Sub RunQuickEqList(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_EQ_LIST_DEFAULTS
End Sub

Public Sub RunQuickSchedule(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_SCHEDULE_DEFAULTS
End Sub

Public Sub RunQuickSplitSheets(control As IRibbonControl)
    WithPerformance PERF_SPLIT_SHEET_BY_COLUMN_DEFAULTS
End Sub


' ---- No ribbon button ----------------------------------------------------

' The "EQ Count" button was removed in #120 - the count is now an EQ List
' option instead. The flow and this callback are kept so the standalone
' command can be put back (or run by name) without rebuilding it.
Public Sub RunCountEquipmentRows(control As IRibbonControl)
    WithPerformance PERF_COUNT_EQ_ROWS
End Sub


Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    InitializeAddIn
End Sub
