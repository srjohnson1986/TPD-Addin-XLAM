Attribute VB_Name = "modMain_CustEQListHeader"
'@Folder("TPD_Addin.EQList")

'===========================================================
'  "Default EQ List Header" ribbon entry point - a thin
'  wrapper that runs InsertDefaultCustEQHeader on the active
'  sheet. ("header" = the TPD title block, not the data
'  heading row.)
'===========================================================

Option Explicit

Public Sub RunInsertDefaultCustEQHeader(control As IRibbonControl)
    InsertDefaultCustEQHeader ActiveSheet
End Sub
