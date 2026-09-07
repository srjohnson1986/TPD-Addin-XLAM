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

' Ribbon callback - "Create Customer Schedule" (opens the column picker)
Public Sub CreateCustSchedule(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_SCHEDULE
End Sub

' Ribbon callback - "Schedule" one-click: no picker, columns come straight
' from the saved Set TPD Defaults value (or the shipped default list) (#96).
Public Sub CreateCustScheduleFromDefaults(control As IRibbonControl)
    WithPerformance PERF_CREATE_CUST_SCHEDULE_DEFAULTS
End Sub

Public Sub CreateCustSchedule_Internal()

    Dim wsSource As Worksheet
    Dim headings As Variant
    Dim frm As frmCustScheduleColumnPicker
    Dim selectedCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub
    headings = modHelpers_Headers.GetHeadingList(wsSource, 1)

    Set frm = New frmCustScheduleColumnPicker
    frm.LoadColumns headings
    frm.Show

    If frm.Cancelled Then Exit Sub

    Set selectedCols = GetSelectedColumns(frm.fraColumns)

    CreateCustSchedule_DoWork wsSource, selectedCols
End Sub

Public Sub CreateCustScheduleFromDefaults_Internal()

    Dim wsSource As Worksheet
    Dim wantCols As Collection
    Dim presentCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    ' One-click precedence: Set TPD Defaults value -> shipped default list.
    ' (No picker LastUsed* layer here - that's the picker's alone.)
    Set wantCols = ResolveColumnList(Array(PREF_SCHEDULE_COLUMNS), DefaultScheduleColumns())
    Set presentCols = HeadingsPresentOnSheet(wsSource, 1, wantCols)

    If presentCols.Count = 0 Then
        MsgBox "None of your default Schedule columns were found on '" & wsSource.name & "'." & vbCrLf & vbCrLf & _
               "Open TPD " & Chr$(187) & " Defaults " & Chr$(187) & " Set TPD Defaults to change the list, " & _
               "or use Create Customer Schedule to pick columns for this sheet.", _
               vbExclamation, "TPD Add-in"
        Exit Sub
    End If

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
    ' alignment, centre the rest (FormatEQSheet defaults to "Description").
    FormatEQSheet wsNew, 1, "Tasks"
    AutoFitUsedColumns wsNew

    ' Inserts the 5 header rows, the Date text/value, and the logo.
    InsertDefaultCustScheduleHeader wsNew

    SafeFreezePanes wsNew, SCHEDULE_HEADER_ROW_COUNT + 1
End Sub
