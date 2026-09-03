Attribute VB_Name = "modMain_SplitSheets"
'@Folder("TPD_Addin.SplitExport")

'===========================================================
'  "Split Sheet by Column" flow: one sheet per unique value in
'  a chosen group column, each carrying the selected columns'
'  rows. Per-value failures are collected and reported in one
'  summary. _DoWork is the testable core (no UI).
'===========================================================

Option Explicit

' Ribbon callback
Public Sub SplitSheetByColumn(control As IRibbonControl)
    WithPerformance PERF_SPLIT_SHEET_BY_COLUMN
End Sub


Public Sub SplitSheetByColumn_Internal()

    Dim wsSource As Worksheet
    Dim headings As Variant
    Dim frm As splitSheetByColumnOptionsForm
    Dim selectedCols As Collection
    Dim groupCol As String

    Set wsSource = GetFirstVisibleSheet()
    If wsSource Is Nothing Then Exit Sub
    headings = GetHeadingList(wsSource, 1)

    Set frm = New splitSheetByColumnOptionsForm
    frm.LoadColumns headings
    frm.Show

    If frm.Cancelled Then Exit Sub

    groupCol = frm.cboGroupColumn.value
    Set selectedCols = GetSelectedColumns(frm.fraColumns)

    Dim createdCount As Long
    Dim failures As Collection
    Set failures = New Collection

    createdCount = SplitSheetByColumn_DoWork(wsSource, groupCol, selectedCols, failures)

    ReportSplitOutcome createdCount, failures

End Sub


' Splits wsSource into one sheet per unique value in groupColumn, keeping only
' selectedCols. Returns the number of sheets successfully created; any group
' value that raised an error is added to failures as a "'value': reason" string
' and the run continues (mirrors ExportSheets_DoWork). failures may be passed in
' Nothing - it will be allocated.
Public Function SplitSheetByColumn_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal groupColumn As String, _
        ByVal selectedCols As Collection, _
        ByRef failures As Collection) As Long

    Dim headings As Variant
    Dim groupColIndex As Long
    Dim rawValues As Collection
    Dim uniqueValues As Collection
    Dim seen As Object
    Dim usedNames As Object
    Dim rawValue As Variant
    Dim cleaned As String
    Dim groupValue As Variant
    Dim createdCount As Long

    If failures Is Nothing Then Set failures = New Collection

    ' Sheet names produced this run - used to disambiguate distinct group
    ' values that sanitize to the same 31-char name (#57).
    Set usedNames = CreateObject("Scripting.Dictionary")
    usedNames.CompareMode = vbTextCompare

    ' Get headings and find the group column index
    headings = GetHeadingList(wsSource, 1)
    groupColIndex = modHelpers_Headers.FindHeadingIndex(headings, groupColumn, preferRightmost:=True)

    If groupColIndex = 0 Then
        MsgBox "Group column '" & groupColumn & "' not found.", vbExclamation
        Exit Function
    End If

    ' Build a cleaned, deduped list of the values to split on
    Set rawValues = GetUniqueValuesInColumn(wsSource, groupColIndex)
    Set seen = CreateObject("Scripting.Dictionary")
    Set uniqueValues = New Collection

    For Each rawValue In rawValues
        cleaned = NormalizeCellText(rawValue)
        If Len(cleaned) > 0 Then
            If Not seen.exists(cleaned) Then
                seen.Add cleaned, cleaned
                uniqueValues.Add cleaned
            End If
        End If
    Next rawValue

    ' One sheet per value. A failure on one value is recorded and skipped
    ' rather than aborting the whole run or vanishing into Debug.Print.
    For Each groupValue In uniqueValues
        On Error Resume Next
        Err.Clear
        SplitOneGroup wsSource, groupColIndex, selectedCols, groupValue, usedNames

        If Err.Number = 0 Then
            createdCount = createdCount + 1
        Else
            failures.Add "'" & CStr(groupValue) & "': " & Err.Description
            Debug.Print "ERROR ON:", groupValue, "ERR:", Err.Number, Err.Description
        End If
        On Error GoTo 0
    Next groupValue

    SplitSheetByColumn_DoWork = createdCount

End Function


' Builds (or clears and refills) the destination sheet for a single group value.
' Raises on any failure - the caller records it against this value. usedNames
' accumulates the sheet names created this run so colliding values don't
' overwrite each other.
Private Sub SplitOneGroup( _
        ByVal wsSource As Worksheet, _
        ByVal groupColIndex As Long, _
        ByVal selectedCols As Collection, _
        ByVal groupValue As Variant, _
        ByVal usedNames As Object)

    Dim wsNew As Worksheet
    Dim newName As String

    ' ShowAllData raises if AutoFilterMode is on but nothing is actually
    ' filtered, so guard it separately from the real work.
    On Error Resume Next
    If wsSource.AutoFilterMode Then wsSource.AutoFilter.ShowAllData
    On Error GoTo 0

    newName = UniqueRunName(SanitizeSheetName(CStr(groupValue)), usedNames)
    usedNames.Add newName, True

    Set wsNew = CreateOrClearSheet(ActiveWorkbook, newName)

    CopyFilteredRowsByColumns wsSource, wsNew, groupColIndex, selectedCols, groupValue, 1

    FormatSplitSheet wsNew, 1
    SafeFreezePanes wsNew, 1

End Sub

' Returns baseName if it hasn't been used this run, otherwise baseName with a
' " (2)" / " (3)" / ... suffix - trimming baseName so the result stays within
' Excel's 31-char sheet-name limit.
Private Function UniqueRunName(ByVal baseName As String, ByVal usedNames As Object) As String
    Dim n As Long
    Dim suffix As String
    Dim candidate As String

    If Not usedNames.exists(baseName) Then
        UniqueRunName = baseName
        Exit Function
    End If

    n = 2
    Do
        suffix = " (" & n & ")"
        candidate = Left$(baseName, 31 - Len(suffix)) & suffix
        n = n + 1
    Loop While usedNames.exists(candidate)

    UniqueRunName = candidate
End Function


Private Sub ReportSplitOutcome(ByVal createdCount As Long, ByVal failures As Collection)
    Dim msg As String
    Dim f As Variant

    msg = createdCount & " sheet(s) created."

    If failures.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & failures.Count & " value(s) could not be split out:"
        For Each f In failures
            msg = msg & vbCrLf & "  " & f
        Next f
        MsgBox msg, vbExclamation, "Split finished with errors"
    Else
        MsgBox msg, vbInformation, "Split complete"
    End If
End Sub
