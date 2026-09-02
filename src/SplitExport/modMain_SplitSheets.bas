Attribute VB_Name = "modMain_SplitSheets"
'@Folder("TPD_Addin.SplitExport")

Option Explicit

' Ribbon callback
Public Sub SplitSheetByColumn(control As IRibbonControl)
    'MsgBox "Ribbon callback fired"
    WithPerformance "SplitSheetByColumn_Internal"
End Sub


Public Sub SplitSheetByColumn_Internal()

    Dim wsSource As Worksheet
    Dim headings As Variant
    Dim frm As splitSheetByColumnOptionsForm
    Dim selectedCols As Collection
    Dim groupCol As String

    Set wsSource = GetFirstVisibleSheet()
    headings = GetHeadingList(wsSource, 1)

    Set frm = New splitSheetByColumnOptionsForm
    frm.LoadColumns headings
    frm.Show

    If frm.Cancelled Then Exit Sub

    groupCol = frm.cboGroupColumn.value
    Set selectedCols = GetSelectedColumns(frm.fraColumns)

    Dim uniqueCount As Long
    uniqueCount = SplitSheetByColumn_DoWork(wsSource, groupCol, selectedCols)
    
    MsgBox uniqueCount & " sheets created successfully.", vbInformation

End Sub


Public Function SplitSheetByColumn_DoWork( _
        ByVal wsSource As Worksheet, _
        ByVal groupColumn As String, _
        ByVal selectedCols As Collection) As Long

    Dim headings As Variant
    Dim groupColIndex As Long
    Dim rawValues As Collection
    Dim uniqueValues As Collection
    Dim dict As Object
    Dim v As Variant
    Dim cleaned As String
    Dim wsNew As Worksheet
    Dim newName As String
    Dim keyValue As Variant

    ' Get headings and find the group column index
    headings = GetHeadingList(wsSource, 1)
    groupColIndex = modHelpers_Headers.FindHeadingIndex(headings, groupColumn)
    
    If groupColIndex = 0 Then
        MsgBox "Group column '" & groupColumn & "' not found.", vbExclamation
        Exit Function
    End If

    ' Get raw unique values from the sheet
    Set rawValues = GetUniqueValuesInColumn(wsSource, groupColIndex)

    ' Build cleaned, deduped list
    Set dict = CreateObject("Scripting.Dictionary")
    Set uniqueValues = New Collection

    For Each v In rawValues
        cleaned = CleanValue(v)
       
        If Len(cleaned) > 0 Then
            If Not dict.exists(cleaned) Then
                dict.Add cleaned, cleaned
                uniqueValues.Add cleaned
            End If
        End If
    Next v

    On Error GoTo LoopError
    ' Loop through cleaned unique values
    For Each keyValue In uniqueValues
    
        ' Ensure no filters interfere
        If wsSource.AutoFilterMode Then
            On Error Resume Next
            wsSource.AutoFilter.ShowAllData
            On Error GoTo 0
        End If

        ' Create or clear the destination sheet
        newName = SanitizeSheetName(CStr(keyValue))
        Set wsNew = CreateOrClearSheet(ActiveWorkbook, newName)
        
        ' Copy rows for this value
        CopyFilteredRowsByColumns _
            wsSource, _
            wsNew, _
            groupColIndex, _
            selectedCols, _
            keyValue, _
            1

        ' Format the new sheet
        If TypeName(wsNew) = "Worksheet" Then
            FormatSplitSheet wsNew, 1
            SafeFreezePanes wsNew, 1
        End If


    Next keyValue

    ' Return count of unique cleaned values
    SplitSheetByColumn_DoWork = uniqueValues.Count

Exit Function

LoopError:
    Debug.Print "ERROR ON:", keyValue, "ERR:", Err.Number, Err.Description
    Resume Next
End Function


