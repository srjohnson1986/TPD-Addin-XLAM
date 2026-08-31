Attribute VB_Name = "modMain_CustEQList"
'@Folder("TPD_Addin.EQList")

Option Explicit

' Ribbon callback
Public Sub CreateCustEQList(control As IRibbonControl)
    'MsgBox "Ribbon callback fired"
    WithPerformance "CreateCustEQList_Internal"
End Sub

Public Sub CreateCustEQList_Internal()

    Dim wsSource As Worksheet
    Dim headers As Variant
    Dim frm As CustEQListColumnPickerForm
    Dim selectedCols As Collection
    Dim imgPath As String

    Set wsSource = GetFirstVisibleSheet()
    headers = GetHeaderList(wsSource, 1)

    Set frm = New CustEQListColumnPickerForm
    frm.LoadColumns headers
    frm.Show

    If frm.Cancelled Then Exit Sub

    Set selectedCols = GetSelectedColumns(frm.fraColumns)
    imgPath = frm.imgPath

    CreateCustEQList_DoWork wsSource, selectedCols, imgPath
End Sub

Public Sub CreateCustEQList_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal selectedCols As Collection, _
        ByVal imgPath As String)

    Dim wsNew As Worksheet
    Dim newSheetName As String

    newSheetName = GetNextAvailableSheetName("Customer EQ List")
    Set wsNew = CreateNewSheetAfterLast(ActiveWorkbook, newSheetName)

    CopyAllRowsPreserveGroups wsSource, wsNew
    DeleteUnselectedColumnsByHeader wsNew, selectedCols, 1

    FormatEQSheet wsNew, 1
    AutoFitUsedColumns wsNew
    
    InsertEQHeaderBlock wsNew
    SafeInsertLogoAtRight wsNew, imgPath, 1, 7
    
    SafeFreezePanes wsNew, 7
End Sub

