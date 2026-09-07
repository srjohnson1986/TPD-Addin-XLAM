Attribute VB_Name = "modMain_CustEQList"
'@Folder("TPD_Addin.EQList")

'===========================================================
'  "Create Customer EQ List" flow: shows the column picker,
'  copies every row to a new sheet, drops the unselected
'  columns, formats it, and adds the EQ header block + the
'  embedded default logo + frozen panes. _DoWork is the
'  testable core (no UI).
'===========================================================

Option Explicit

' Ribbon callback - "Create Customer EQ List" (opens the column picker)
Public Sub CreateCustEQList(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_EQ_LIST
End Sub

' Ribbon callback - "EQ List" one-click: no picker, columns come straight
' from the saved Set TPD Defaults value (or the shipped default list) (#96).
Public Sub CreateCustEQListFromDefaults(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_EQ_LIST_DEFAULTS
End Sub

Public Sub CreateCustEQList_Internal()

    Dim wsSource As Worksheet
    Dim headings As Variant
    Dim frm As CustEQListColumnPickerForm
    Dim selectedCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub
    headings = modHelpers_Headers.GetHeadingList(wsSource, 1)

    Set frm = New CustEQListColumnPickerForm
    frm.LoadColumns headings
    frm.Show

    If frm.Cancelled Then Exit Sub

    Set selectedCols = GetSelectedColumns(frm.fraColumns)

    CreateCustEQList_DoWork wsSource, selectedCols
End Sub

Public Sub CreateCustEQListFromDefaults_Internal()

    Dim wsSource As Worksheet
    Dim wantCols As Collection
    Dim presentCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    ' One-click precedence: Set TPD Defaults value -> shipped default list.
    ' (No picker LastUsed* layer here - that's the picker's alone.)
    Set wantCols = ResolveColumnList(Array(PREF_EQLIST_COLUMNS), DefaultEqListColumns())
    Set presentCols = HeadingsPresentOnSheet(wsSource, 1, wantCols)

    If presentCols.Count = 0 Then
        MsgBox "None of your default EQ List columns were found on '" & wsSource.name & "'." & vbCrLf & vbCrLf & _
               "Open TPD " & Chr$(187) & " Defaults " & Chr$(187) & " Set TPD Defaults to change the list, " & _
               "or use Create Customer EQ List to pick columns for this sheet.", _
               vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    CreateCustEQList_DoWork wsSource, presentCols
End Sub

Public Sub CreateCustEQList_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal selectedCols As Collection)

    Dim wsNew As Worksheet
    Dim newSheetName As String

    newSheetName = GetNextAvailableSheetName("Customer EQ List")
    Set wsNew = CreateNewSheetAfterLast(ActiveWorkbook, newSheetName)

    CopyEntireSheetRows wsSource, wsNew
    modHelpers_Columns.DeleteUnselectedColumnsByHeading wsNew, selectedCols, 1

    FormatEQSheet wsNew, 1
    AutoFitUsedColumns wsNew

    InsertEQHeaderBlock wsNew
    ' Embedded default logo, top-right of the heading row, scaled to the
    ' header block - same call the "Default EQ List Header" command uses.
    InsertDefaultLogo wsNew, EQ_HEADER_ROW_COUNT + 1, "right-top", EQ_HEADER_ROW_COUNT

    SafeFreezePanes wsNew, EQ_HEADER_ROW_COUNT + 1
End Sub


