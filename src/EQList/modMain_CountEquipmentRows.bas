Attribute VB_Name = "modMain_CountEquipmentRows"
'@Folder("TPD_Addin.EQList")

'===========================================================
'  "EQ Count" flow: inserts a leftmost EQ COUNT column and
'  numbers every data row except those marked PARENT or
'  INCLUDED in the Purchased column, with leading-zero
'  formatting sized to the highest number written. A blank
'  Purchased value is a real line and is numbered by design
'  (#6, closed working-as-intended).
'===========================================================

Option Explicit

Private Const EQ_COUNT_COL As Long = 1

' The standalone "EQ Count" flow. Its ribbon button was removed in #120 (the
' count is an EQ List option now); the callback that reaches this lives with
' the others in modRibbonCallbacks.
Public Sub CountEquipmentRows_Internal()
    Dim wb As Workbook
    Dim ws As Worksheet

    If Not RequireActiveWorkbook(wb) Then Exit Sub

    ' Set fails (leaving ws Nothing) if the active sheet is a chart sheet
    On Error Resume Next
    Set ws = wb.ActiveSheet
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Select a worksheet and try again.", vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    NumberEquipmentRows ws
End Sub

' Inserts the EQ COUNT column on ws (heading row 1) and numbers its data rows.
' Reads the Purchased status straight off the sheet - the standalone "EQ Count"
' ribbon flow.
Public Sub NumberEquipmentRows(ws As Worksheet)
    Dim purchasedCol As Long

    purchasedCol = FindHeadingIndex(GetHeadingList(ws, 1), "PURCHASED", preferRightmost:=False)
    If purchasedCol = 0 Then
        MsgBox "No column with heading PURCHASED was found.", vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    NumberEquipmentRowsFromStatuses ws, 1, PurchasedStatuses(ws, 1, purchasedCol)
End Sub

' Uppercased, trimmed Purchased value for each data row (1-based from
' headingRow + 1). Returns Array() when there are no data rows. Row count
' comes from GetLastRow (the whole-sheet Find, #90) not the Purchased
' column's End(xlUp) - a blank Purchased value is a valid line (#6) and the
' last rows can legitimately be blank there.
Public Function PurchasedStatuses(ws As Worksheet, headingRow As Long, purchasedCol As Long) As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim arr() As String

    lastRow = GetLastRow(ws)
    If lastRow <= headingRow Then
        PurchasedStatuses = Array()
        Exit Function
    End If

    ReDim arr(1 To lastRow - headingRow)
    For r = headingRow + 1 To lastRow
        arr(r - headingRow) = UCase$(Trim$(CStr(ws.Cells(r, purchasedCol).value)))
    Next r
    PurchasedStatuses = arr
End Function

' Inserts a leftmost EQ COUNT column and numbers each data row whose captured
' status is neither PARENT nor INCLUDED, leading-zero-formatted to the highest
' number written. statuses is passed in rather than read here so the Customer
' EQ List flow can number after its Purchased column has been projected away
' (#120); statuses(k) is the status of the k-th data row below headingRow.
Public Sub NumberEquipmentRowsFromStatuses(ws As Worksheet, headingRow As Long, statuses As Variant)
    Dim k As Long
    Dim rowNum As Long
    Dim eqNumber As Long
    Dim lastDataRow As Long
    Dim countRange As Range

    ' EQ COUNT becomes the leftmost column; everything to its right shifts over.
    ws.Columns(EQ_COUNT_COL).Insert Shift:=xlToRight
    ws.Cells(headingRow, EQ_COUNT_COL).value = "EQ COUNT"

    ' The inserted column comes in with no formatting - copy the neighbouring
    ' heading cell's look onto it (bold, fill, borders, alignment, whatever the
    ' sheet's column headings carry) so EQ COUNT doesn't stand out.
    ws.Cells(headingRow, EQ_COUNT_COL + 1).Copy
    ws.Cells(headingRow, EQ_COUNT_COL).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    If Not IsArray(statuses) Then Exit Sub
    If UBound(statuses) < LBound(statuses) Then Exit Sub   ' heading row only

    eqNumber = 1
    For k = LBound(statuses) To UBound(statuses)
        rowNum = headingRow + k
        If statuses(k) = "PARENT" Or statuses(k) = "INCLUDED" Then
            ws.Cells(rowNum, EQ_COUNT_COL).value = ""
        Else
            ws.Cells(rowNum, EQ_COUNT_COL).value = eqNumber
            eqNumber = eqNumber + 1
        End If
    Next k

    lastDataRow = headingRow + UBound(statuses)
    Set countRange = ws.Range(ws.Cells(headingRow + 1, EQ_COUNT_COL), ws.Cells(lastDataRow, EQ_COUNT_COL))
    countRange.NumberFormat = String(Len(CStr( _
        Application.WorksheetFunction.Max(countRange))), "0")
End Sub
