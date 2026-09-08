Attribute VB_Name = "modMain_CustSchedule"
'@Folder("TPD_Addin.Schedule")

'===========================================================
'  "Create Customer Schedule" flow - the schedule twin of
'  modMain_CustEQList. Copies every row to a new "Customer
'  Schedule" sheet BEFORE any formatting, drops the unselected
'  columns, formats it, and adds the schedule header block +
'  the embedded default logo + frozen panes.
'  InsertDefaultCustScheduleHeader (modHelpers_SheetSetup)
'  inserts the header rows, the date text and the logo in one
'  call, so _DoWork does not add a second logo.
'  _DoWork is the testable core (no UI).
'===========================================================

Option Explicit

' "Create Customer Schedule" - the picker path. Entered from
' modRibbonCallbacks.RunCreateCustSchedule via WithPerformance.
Public Sub CreateCustSchedule_Internal()

    Dim wsSource As Worksheet
    Dim frm As frmCustScheduleColumnPicker

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    Set frm = New frmCustScheduleColumnPicker
    frm.LoadColumns modHelpers_Headers.GetHeadingList(wsSource, 1)
    frm.Show

    If frm.Cancelled Then Exit Sub

    CreateCustSchedule_DoWork wsSource, GetSelectedColumns(frm.fraColumns)
End Sub

' "Schedule" - the one-click path. Same resolution the EQ List one-click uses
' (saved Set Defaults value, else the shipped list, narrowed to the headings
' on the sheet); ResolveOneClickColumns shows the reason and returns Nothing
' when nothing matches (#96).
Public Sub CreateCustScheduleFromDefaults_Internal()

    Dim wsSource As Worksheet
    Dim presentCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    Set presentCols = ResolveOneClickColumns(wsSource, PREF_SCHEDULE_COLUMNS, _
                                             DefaultScheduleColumns(), _
                                             "Schedule", "Create Customer Schedule")
    If presentCols Is Nothing Then Exit Sub

    CreateCustSchedule_DoWork wsSource, presentCols
End Sub

Public Sub CreateCustSchedule_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal selectedCols As Collection)

    Dim wsNew As Worksheet
    Dim newSheetName As String

    newSheetName = GetNextAvailableSheetName("Customer Schedule")
    Set wsNew = CreateNewSheetAfterLast(ActiveWorkbook, newSheetName)

    CopyEntireSheetRows wsSource, wsNew
    ' Schedule keeps no structural rows - strip every carried-over fill (#130).
    modHelpers_SheetFormatting.StripDataRowFill wsNew, 1
    modHelpers_Columns.DeleteUnselectedColumnsByHeading wsNew, selectedCols, 1
    ' DeleteUnselectedColumnsByHeading keeps the survivors in source order;
    ' reorder them to match the chosen column order (#127).
    modHelpers_Columns.ReorderColumnsByHeading wsNew, selectedCols, 1

    ' "Tasks" is the schedule's long-text column - keep its natural left
    ' alignment, centre the rest (FormatDataTable defaults to "Description").
    FormatDataTable wsNew, 1, "Tasks"
    AutoFitUsedColumns wsNew

    ' Inserts the 5 header rows, the Date text/value, and the logo.
    InsertDefaultCustScheduleHeader wsNew

    SafeFreezePanes wsNew, SCHEDULE_HEADER_ROW_COUNT + 1
End Sub
