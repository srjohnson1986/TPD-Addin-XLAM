Attribute VB_Name = "modMain_CustEQList"
'@Folder("TPD_Addin.EQList")

'===========================================================
'  "Create Customer EQ List" flow: shows the column picker,
'  copies every row to a new sheet, drops the unselected
'  columns, formats it, and adds the EQ header block + logo +
'  frozen panes. _DoWork is the testable core (no UI).
'===========================================================

Option Explicit

' Ribbon callback
Public Sub CreateCustEQList(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_EQ_LIST
End Sub

Public Sub CreateCustEQList_Internal()

    Dim wsSource As Worksheet
    Dim headings As Variant
    Dim frm As CustEQListColumnPickerForm
    Dim selectedCols As Collection
    Dim imgPath As String

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub
    headings = modHelpers_Headers.GetHeadingList(wsSource, 1)

    Set frm = New CustEQListColumnPickerForm
    frm.LoadColumns headings
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

    CopyEntireSheetRows wsSource, wsNew
    modHelpers_Columns.DeleteUnselectedColumnsByHeading wsNew, selectedCols, 1

    FormatEQSheet wsNew, 1
    AutoFitUsedColumns wsNew
    
    InsertEQHeaderBlock wsNew
    SafeInsertLogoAtRight wsNew, imgPath, 1, EQ_HEADER_ROW_COUNT + 1

    SafeFreezePanes wsNew, EQ_HEADER_ROW_COUNT + 1
End Sub


