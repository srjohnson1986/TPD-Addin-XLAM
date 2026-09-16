Attribute VB_Name = "modHelpers_Columns"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Column-oriented helpers: case-insensitive collection
'  membership, deleting columns whose heading isn't in a
'  selected list, reordering the kept columns to a chosen
'  order, unique values in a column, and copying rows
'  between sheets (all rows, or only rows matching a value in
'  the group column) projected onto a chosen set of columns.
'
'  Sheet-facing routines (Public Sub/Function taking a Worksheet) are thin
'  wrappers around a pure counterpart that takes/returns plain values only -
'  HeadingsPresentIn, UniqueValues, BuildOrderedColumnMap, PlanColumnMoves
'  (#163, prep for Rubberduck unit tests). Extract new pure helpers the same
'  way.
'===========================================================

Option Explicit

Public Function CollectionContainsText(col As Collection, value As String) As Boolean
    Dim v As Variant
    For Each v In col
        If StrComp(CStr(v), value, vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next v
End Function

Public Sub DeleteUnselectedColumnsByHeading(ws As Worksheet, selectedHeadings As Collection, headingsRow As Long)
    Dim lastCol As Long
    Dim colHeading As String
    Dim c As Long

    lastCol = GetLastCol(ws, headingsRow)

    For c = lastCol To 1 Step -1
        colHeading = CStr(ws.Cells(headingsRow, c).value)
        If Not CollectionContainsText(selectedHeadings, colHeading) Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub

' Reorders ws's columns so their headings on headingsRow run left-to-right in
' orderedHeadings order, moving each wanted column into place with cut/insert.
' Headings not found at/after the current target position are skipped (already
' placed, or absent); any columns whose heading isn't in orderedHeadings are
' left untouched to the right of the ordered block.
'
' Pairs with DeleteUnselectedColumnsByHeading: that drops the unselected
' columns but leaves the survivors in source order, so the EQ List / Schedule
' flows call this straight after to honour the order set in Set TPD Defaults
' (#127). Cheap no-op when orderedHeadings is already in source order (e.g. the
' checkbox pickers, which can't express a custom order).
'
' Thin sheet-mutating wrapper - see PlanColumnMoves (#163).
Public Sub ReorderColumnsByHeading(ByVal ws As Worksheet, ByVal orderedHeadings As Collection, ByVal headingsRow As Long)
    Dim moves As Collection
    Dim mv As Variant

    Set moves = PlanColumnMoves(modHelpers_Headers.GetHeadingList(ws, headingsRow), orderedHeadings)

    For Each mv In moves
        ws.Columns(CLng(mv(0))).Cut
        ws.Columns(CLng(mv(1))).Insert Shift:=xlToRight
    Next mv
End Sub

' Pure (#163): the sequence of column moves - each a 1-based [from, to] pair,
' applied in order via Cut column `from` + Insert before column `to`
' (everything from `to` to `from`-1 shifts one column right, matching Excel's
' actual Cut/Insert semantics) - that reorders `headings` into
' `orderedHeadings` order. A heading in `orderedHeadings` not found at/after
' the current target position is skipped, same as ReorderColumnsByHeading.
Public Function PlanColumnMoves(ByVal headings As Variant, ByVal orderedHeadings As Collection) As Collection
    Dim moves As New Collection
    Dim arr() As String
    Dim n As Long, i As Long
    Dim targetCol As Long
    Dim h As Variant
    Dim curIdx As Long
    Dim j As Long
    Dim moved As String

    n = UBound(headings) - LBound(headings) + 1
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(headings(LBound(headings) + i - 1))
    Next i

    targetCol = 1
    For Each h In orderedHeadings
        curIdx = 0
        For j = targetCol To n
            If StrComp(arr(j), CStr(h), vbTextCompare) = 0 Then
                curIdx = j
                Exit For
            End If
        Next j

        If curIdx = targetCol Then
            targetCol = targetCol + 1
        ElseIf curIdx > targetCol Then
            moves.Add Array(curIdx, targetCol)

            ' Simulate the Cut+Insert so later scans see the post-move order.
            moved = arr(curIdx)
            For j = curIdx To targetCol + 1 Step -1
                arr(j) = arr(j - 1)
            Next j
            arr(targetCol) = moved

            targetCol = targetCol + 1
        End If
    Next h

    Set PlanColumnMoves = moves
End Function

' The one-click flows use this to avoid handing DeleteUnselectedColumnsByHeading
' a selection that matches nothing on the source sheet - which would delete
' every column (#96). Thin sheet-reading wrapper - see HeadingsPresentIn (#163).
Public Function HeadingsPresentOnSheet( _
        ByVal ws As Worksheet, _
        ByVal headingsRow As Long, _
        ByVal wanted As Collection) As Collection

    Set HeadingsPresentOnSheet = HeadingsPresentIn(modHelpers_Headers.GetHeadingList(ws, headingsRow), wanted)
End Function

' Pure (#163): the subset of `wanted` heading names that actually appear in
' `headings`, preserving `wanted`'s order and its (possibly dirty) spelling.
Public Function HeadingsPresentIn(ByVal headings As Variant, ByVal wanted As Collection) As Collection
    Dim present As New Collection
    Dim onSheet As New Collection
    Dim i As Long
    Dim w As Variant

    For i = LBound(headings) To UBound(headings)
        onSheet.Add CStr(headings(i))
    Next i

    For Each w In wanted
        If CollectionContainsText(onSheet, CStr(w)) Then present.Add CStr(w)
    Next w

    Set HeadingsPresentIn = present
End Function

' The column list a one-click ribbon button should act on: the user's saved
' Set Defaults value (defaultKey), else the shipped list, narrowed to the
' headings that are actually on ws. Returns Nothing - having shown the reason -
' when none of them is, which would otherwise hand
' DeleteUnselectedColumnsByHeading a selection matching nothing and take every
' column with it (#96). flowName / pickerCommand name the flow and its picker
' command in that message ("EQ List" / "Create Customer EQ List").
'
' Shared by the EQ List and Schedule one-click flows, which resolve columns
' identically - only their keys, shipped list and those two labels differ.
Public Function ResolveOneClickColumns(ByVal ws As Worksheet, _
                                       ByVal defaultKey As String, _
                                       ByVal shippedCsv As String, _
                                       ByVal flowName As String, _
                                       ByVal pickerCommand As String) As Collection
    Dim present As Collection

    ' No picker LastUsed* layer here - that one is the picker's alone (#96).
    Set present = HeadingsPresentOnSheet(ws, 1, ResolveColumnList(Array(defaultKey), shippedCsv))
    If present.Count = 0 Then
        MsgBox "None of your default " & flowName & " columns were found on '" & ws.name & "'." & vbCrLf & vbCrLf & _
               "Open TPD " & Chr$(187) & " One-click " & Chr$(187) & " Set Defaults to change the list, " & _
               "or use " & pickerCommand & " to pick columns for this sheet.", _
               vbExclamation, "TPD Add-in"
        Exit Function
    End If

    Set ResolveOneClickColumns = present
End Function

' Thin sheet-reading wrapper (skips the heading row) - see UniqueValues (#163).
Public Function GetUniqueValuesInColumn(ws As Worksheet, colIndex As Long) As Collection
    Dim lastRow As Long
    Dim raw() As Variant
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    ReDim raw(2 To lastRow)
    For i = 2 To lastRow
        raw(i) = ws.Cells(i, colIndex).value
    Next i

    Set GetUniqueValuesInColumn = UniqueValues(raw)
End Function

' Pure (#163): the distinct NormalizeCellText'd entries of `values`, skipping
' blanks, in first-seen order.
Public Function UniqueValues(ByVal values As Variant) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim v As String

    For i = LBound(values) To UBound(values)
        v = NormalizeCellText(values(i))
        If Len(v) > 0 Then
            If Not CollectionContainsText(result, v) Then result.Add v
        End If
    Next i

    Set UniqueValues = result
End Function


' Pure (#163): maps `selectedHeadings` (in that order) to their column index in
' `headings` - case-insensitive, rightmost match wins (matches
' FindHeadingIndex's default) - skipping any heading not found in `headings`
' and de-duplicating repeat indices. Returned as a Scripting.Dictionary so
' insertion order (== selectedHeadings order) survives on .Keys.
Public Function BuildOrderedColumnMap(ByVal headings As Variant, ByVal selectedHeadings As Collection) As Object
    Dim colMap As Object
    Dim h As Variant
    Dim srcIdx As Long

    Set colMap = CreateObject("Scripting.Dictionary")
    For Each h In selectedHeadings
        srcIdx = modHelpers_Headers.FindHeadingIndex(headings, CStr(h), preferRightmost:=True)
        If srcIdx > 0 Then
            If Not colMap.exists(srcIdx) Then colMap.Add srcIdx, headings(srcIdx)
        End If
    Next h

    Set BuildOrderedColumnMap = colMap
End Function

Public Sub CopyFilteredRowsByColumns( _
    ByVal wsSource As Worksheet, _
    ByVal wsDest As Worksheet, _
    ByVal groupColIndex As Long, _
    ByVal selectedHeadings As Collection, _
    ByVal matchValue As String, _
    ByVal headingsRow As Long)

    Dim lastRow As Long
    Dim headings As Variant
    Dim colMap As Object
    Dim srcColIndex As Variant
    Dim destCol As Long
    Dim destRow As Long
    Dim cell As Range

    lastRow = GetLastRow(wsSource)

    ' Column order comes from selectedHeadings, not the source sheet, so the
    ' output honours the order set in Set TPD Defaults (#127) - see
    ' BuildOrderedColumnMap (#163).
    headings = modHelpers_Headers.GetHeadingList(wsSource, headingsRow)
    Set colMap = BuildOrderedColumnMap(headings, selectedHeadings)

    ' Copy headings row
    destCol = 1
    For Each srcColIndex In colMap.Keys
        wsDest.Cells(headingsRow, destCol).value = wsSource.Cells(headingsRow, CLng(srcColIndex)).value
        destCol = destCol + 1
    Next srcColIndex

    ' Copy rows. Formatting / bold headings / freeze panes / autofit are the
    ' caller's job (FormatSplitSheet + SafeFreezePanes) - this routine only
    ' moves data.
    destRow = headingsRow + 1

    For Each cell In wsSource.Range(wsSource.Cells(headingsRow + 1, groupColIndex), wsSource.Cells(lastRow, groupColIndex))

        If StrComp(NormalizeCellText(cell.value), NormalizeCellText(matchValue), vbTextCompare) = 0 Then

            destCol = 1
            For Each srcColIndex In colMap.Keys
                wsDest.Cells(destRow, destCol).value = wsSource.Cells(cell.Row, CLng(srcColIndex)).value
                destCol = destCol + 1
            Next srcColIndex

            destRow = destRow + 1
        End If

    Next cell

End Sub

' Copies every row of wsSource onto wsDest verbatim. A whole-row copy (rather
' than a range copy) carries Excel's row grouping / outline levels across too,
' which the Customer EQ List needs.
Public Sub CopyEntireSheetRows(wsSource As Worksheet, wsDest As Worksheet)
    wsSource.Rows("1:" & wsSource.Rows.Count).Copy wsDest.Rows(1)
End Sub


