Attribute VB_Name = "modMain_CustEQList"
'@Folder("TPD_Addin.EQList")

'===========================================================
'  "Create Customer EQ List" flow: shows the column picker,
'  copies every row to a new sheet, drops the unselected
'  columns, formats it, and adds the EQ header block + the
'  embedded default logo + frozen panes. _DoWork is the
'  testable core (no UI) and also reads the EQ List behaviour
'  toggles from Set Defaults (#97 / #119-#123): drop PARENT
'  rows, inline EQ Count, plain PARENT rows, cell borders,
'  column filters. The three PARENT-row toggles need the
'  Purchased column, so _DoWork acts on it (and snapshots each
'  row's status) before the column projection drops it.
'===========================================================

Option Explicit

' "Create Customer EQ List" - the picker path. Entered from
' modRibbonCallbacks.RunCreateCustEQList via WithPerformance.
Public Sub CreateCustEQList_Internal()

    Dim wsSource As Worksheet
    Dim frm As frmCustEQListColumnPicker

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    Set frm = New frmCustEQListColumnPicker
    frm.LoadColumns modHelpers_Headers.GetHeadingList(wsSource, 1)
    frm.Show

    If frm.Cancelled Then Exit Sub

    CreateCustEQList_DoWork wsSource, GetSelectedColumns(frm.fraColumns)
End Sub

' "EQ List" - the one-click path: no picker, columns come straight from the
' saved Set Defaults value (or the shipped list), narrowed to what is on the
' sheet (#96). ResolveOneClickColumns shows the reason and returns Nothing
' when nothing matches; the Schedule flow resolves its columns the same way.
Public Sub CreateCustEQListFromDefaults_Internal()

    Dim wsSource As Worksheet
    Dim presentCols As Collection

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub

    Set presentCols = ResolveOneClickColumns(wsSource, PREF_EQLIST_COLUMNS, _
                                             DefaultEqListColumns(), _
                                             "EQ List", "Create Customer EQ List")
    If presentCols Is Nothing Then Exit Sub

    CreateCustEQList_DoWork wsSource, presentCols
End Sub

Public Sub CreateCustEQList_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal selectedCols As Collection)

    Dim wsNew As Worksheet
    Dim newSheetName As String

    ' EQ List behaviour toggles (#97). Read here so the picker and the
    ' one-click flow - both land in _DoWork - honour them the same way.
    Dim removeParents As Boolean, addCount As Boolean, plainParents As Boolean
    Dim addBorders As Boolean, addFilters As Boolean
    removeParents = LoadToggle(PREF_EQLIST_REMOVE_PARENT_ROWS)
    addCount = LoadToggle(PREF_EQLIST_ADD_COUNT_COLUMN)
    plainParents = LoadToggle(PREF_EQLIST_PLAIN_PARENT_ROWS)
    addBorders = LoadToggle(PREF_EQLIST_ADD_CELL_BORDERS, defaultOn:=True)
    addFilters = LoadToggle(PREF_EQLIST_ADD_COLUMN_FILTERS)

    Dim wantsPurchasedToggle As Boolean
    wantsPurchasedToggle = removeParents Or addCount Or plainParents

    newSheetName = GetNextAvailableSheetName("Customer EQ List")
    Set wsNew = CreateNewSheetAfterLast(ActiveWorkbook, newSheetName)

    CopyEntireSheetRows wsSource, wsNew
    ' Drop the source row fill now, while every source column is still here for
    ' the PURCHASED / INTERNAL ID / CUSTOMER ID lookup - keep it only on PARENT
    ' rows and rows with no IDs at all (#130). Same call, same result whether we
    ' got here from the picker or the one-click button.
    modHelpers_SheetFormatting.StripDataRowFill wsNew, 1, preserveStructuralRows:=True

    ' The PARENT-row toggles all read the Purchased column, which the chosen
    ' column list can drop. Act on it now, while it's still here: delete the
    ' PARENT rows if asked, then snapshot each surviving row's status so the
    ' EQ Count / plain-PARENT steps can run after Purchased is projected away
    ' (per #97 - Purchased only reaches the output sheet if the user picked it).
    Dim purchasedCol As Long
    Dim parentStatuses As Variant
    parentStatuses = Array()

    If wantsPurchasedToggle Then
        purchasedCol = modHelpers_Headers.FindHeadingIndex( _
            modHelpers_Headers.GetHeadingList(wsNew, 1), "PURCHASED", preferRightmost:=False)
        If purchasedCol = 0 Then
            removeParents = False: addCount = False: plainParents = False
        Else
            If removeParents Then DeleteParentRows wsNew, 1, purchasedCol
            parentStatuses = PurchasedStatuses(wsNew, 1, purchasedCol)
        End If
    End If

    modHelpers_Columns.DeleteUnselectedColumnsByHeading wsNew, selectedCols, 1
    ' DeleteUnselectedColumnsByHeading keeps the survivors in source order;
    ' reorder them to match the chosen column order (#127).
    modHelpers_Columns.ReorderColumnsByHeading wsNew, selectedCols, 1

    ' Inline EQ Count numbering (#120): insert the leftmost EQ COUNT column
    ' while the heading row is still row 1, before the title block goes on top.
    If addCount Then NumberEquipmentRowsFromStatuses wsNew, 1, parentStatuses

    FormatEQSheet wsNew, 1, addBorders:=addBorders     ' #122

    ' AutoFilter (#123) goes on while the heading is still row 1 and, crucially,
    ' BEFORE the autofit - so the columns are widened to clear the dropdown
    ' buttons instead of the arrow clipping the heading text (#147). The title
    ' block insert below shifts the filter range down to the heading row.
    If addFilters Then AddHeadingRowFilter wsNew, 1

    AutoFitUsedColumns wsNew

    ' Plain PARENT rows (#121): white fill + black, un-bolded text, overriding
    ' the grouping shade StripDataRowFill kept. Do it before the title block so
    ' the snapshot's data-row ordinals map straight onto sheet rows.
    If plainParents Then ApplyPlainParentRows wsNew, 1, parentStatuses

    InsertEQHeaderBlock wsNew
    ' Embedded default logo, top-right of the heading row, scaled to the
    ' header block - same call the "Default EQ List Header" command uses.
    InsertDefaultLogo wsNew, EQ_HEADER_ROW_COUNT + 1, "right-top", EQ_HEADER_ROW_COUNT

    SafeFreezePanes wsNew, EQ_HEADER_ROW_COUNT + 1

    If wantsPurchasedToggle And purchasedCol = 0 Then
        MsgBox "Some EQ List options need a 'Purchased' column, which '" & wsSource.name & _
               "' doesn't have, so they were skipped: " & SkippedPurchasedToggleNames() & "." & vbCrLf & vbCrLf & _
               "The list was still generated.", vbInformation, "TPD Add-in"
    End If
End Sub

' Deletes every data row (below headingRow) whose Purchased value is PARENT.
' Bottom-up so the row indices below stay valid as rows are removed.
Private Sub DeleteParentRows(ByVal ws As Worksheet, ByVal headingRow As Long, ByVal purchasedCol As Long)
    Dim r As Long

    For r = GetLastRow(ws) To headingRow + 1 Step -1
        If UCase$(Trim$(CStr(ws.Cells(r, purchasedCol).value))) = "PARENT" Then ws.Rows(r).Delete
    Next r
End Sub

' White fill, black non-bold text on each PARENT data row. statuses(k) is the
' Purchased status of the k-th data row below headingRow (from PurchasedStatuses,
' captured before the column was dropped). Column widths / borders / row height
' are left alone (#121).
Private Sub ApplyPlainParentRows(ByVal ws As Worksheet, ByVal headingRow As Long, ByVal statuses As Variant)
    Dim k As Long
    Dim lastCol As Long
    Dim rowRange As Range

    If Not IsArray(statuses) Then Exit Sub
    If UBound(statuses) < LBound(statuses) Then Exit Sub

    lastCol = GetLastCol(ws, headingRow)
    For k = LBound(statuses) To UBound(statuses)
        If statuses(k) = "PARENT" Then
            Set rowRange = ws.Range(ws.Cells(headingRow + k, 1), ws.Cells(headingRow + k, lastCol))
            rowRange.Interior.Color = vbWhite
            rowRange.Font.Color = vbBlack
            rowRange.Font.Bold = False
        End If
    Next k
End Sub

' Turns on AutoFilter for the heading row across the used width. Clears any
' existing filter first so a re-run doesn't just toggle it back off.
Private Sub AddHeadingRowFilter(ByVal ws As Worksheet, ByVal headingRow As Long)
    Dim lastCol As Long

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    lastCol = GetLastCol(ws, headingRow)
    ws.Range(ws.Cells(headingRow, 1), ws.Cells(GetLastRow(ws), lastCol)).AutoFilter
End Sub

' Comma-joined display names of the Purchased-dependent toggles the user has
' turned on - read from the saved prefs, since _DoWork forces its local flags
' False once it finds no Purchased column. Used only for the skip notice.
Private Function SkippedPurchasedToggleNames() As String
    Dim names As String

    If LoadToggle(PREF_EQLIST_REMOVE_PARENT_ROWS) Then names = names & ", Remove PARENT rows"
    If LoadToggle(PREF_EQLIST_ADD_COUNT_COLUMN) Then names = names & ", Add EQ Count column"
    If LoadToggle(PREF_EQLIST_PLAIN_PARENT_ROWS) Then names = names & ", Plain PARENT rows"

    If Len(names) > 0 Then SkippedPurchasedToggleNames = Mid$(names, 3)
End Function


