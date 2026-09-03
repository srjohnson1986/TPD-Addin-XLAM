Attribute VB_Name = "modMain_CustScheduleHeader"
'@Folder("TPD_Addin.Schedule")

'===========================================================
'  "Default Schedule Header" ribbon entry point - a thin
'  wrapper that runs InsertDefaultCustScheduleHeader on the
'  active sheet.
'===========================================================

Option Explicit

Public Sub RunInsertDefaultScheduleHeader(control As IRibbonControl)
    InsertDefaultCustScheduleHeader ActiveSheet
End Sub

