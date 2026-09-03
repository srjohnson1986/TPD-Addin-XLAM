Attribute VB_Name = "modMain_SplitSheets"
'@Folder("TPD_Addin.SplitExport")

Option Explicit

' Ribbon callback
Public Sub SplitSheetByColumn(control As IRibbonControl)
    WithPerformance "SplitSheetByColumn_Internal"
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
    Dim rawValue As Variant
    Dim cleaned As String
    Dim groupValue As Variant
    Dim createdCount As Long

    If failures Is Nothing Then Set failures = New Collection

    ' Get headings and find the group column index
    headings = GetHeadingList(wsSource, 1)
    groupColIndex = modHelpers_Headers.FindHeadingIndex(headings, groupColumn)

    If groupColIndex = 0 Then
        MsgBox "Group column '" & groupColumn & "' not found.", vbExclamation
        Exit Function
    End If

    ' Build a cleaned, deduped list of the values to split on
    Set rawValues = GetUniqueValuesInColumn(wsSource, groupColIndex)
    Set seen = CreateObject("Scripting.Dictionary")
    Set uniqueValues = New Collection

    For Each rawValue In rawValues
        cleaned = CleanValue(rawValue)
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
        SplitOneGroup wsSource, groupColIndex, selectedCols, groupValue

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
' Raises on any failure - the caller records it against this value.
Private Sub SplitOneGroup( _
        ByVal wsSource As Worksheet, _
        ByVal groupColIndex As Long, _
        ByVal selectedCols As Collection, _
        ByVal groupValue As Variant)

    Dim wsNew As Worksheet
    Dim newName As String

    ' ShowAllData raises if AutoFilterMode is on but nothing is actually
    ' filtered, so guard it separately from the real work.
    On Error Resume Next
    If wsSource.AutoFilterMode Then wsSource.AutoFilter.ShowAllData
    On Error GoTo 0

    newName = SanitizeSheetName(CStr(groupValue))
    Set wsNew = CreateOrClearSheet(ActiveWorkbook, newName)

    CopyFilteredRowsByColumns wsSource, wsNew, groupColIndex, selectedCols, groupValue, 1

    FormatSplitSheet wsNew, 1
    SafeFreezePanes wsNew, 1

End Sub


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
