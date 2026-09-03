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

' Ribbon callback
Public Sub RunCountEquipmentRows(control As IRibbonControl)
    WithPerformance PERF_COUNT_EQ_ROWS
End Sub

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

' Inserts the EQ COUNT column on ws and numbers its data rows.
Public Sub NumberEquipmentRows(ws As Worksheet)
    Dim purchasedCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim eqNumber As Long
    Dim purchasedStatus As String
    Dim maxNumber As Long
    Dim digits As Long

    purchasedCol = FindPurchasedColumn(ws)
    If purchasedCol = 0 Then
        MsgBox "No column with heading PURCHASED was found.", vbExclamation, "TPD Add-in"
        Exit Sub
    End If

    ' Insert EQ COUNT as the leftmost column; everything to its right
    ' (Purchased included) shifts one column right.
    ws.Columns(EQ_COUNT_COL).Insert Shift:=xlToRight
    ws.Cells(1, EQ_COUNT_COL).value = "EQ COUNT"
    purchasedCol = purchasedCol + 1

    lastRow = ws.Cells(ws.Rows.Count, purchasedCol).End(xlUp).Row
    If lastRow < 2 Then Exit Sub          ' heading row only - nothing to number

    eqNumber = 1
    For i = 2 To lastRow
        purchasedStatus = UCase$(Trim$(CStr(ws.Cells(i, purchasedCol).value)))

        If purchasedStatus = "PARENT" Or purchasedStatus = "INCLUDED" Then
            ws.Cells(i, EQ_COUNT_COL).value = ""
        Else
            ws.Cells(i, EQ_COUNT_COL).value = eqNumber
            eqNumber = eqNumber + 1
        End If
    Next i

    ' Leading-zero display format, widened to the highest number written
    maxNumber = Application.WorksheetFunction.Max( _
        ws.Range(ws.Cells(2, EQ_COUNT_COL), ws.Cells(lastRow, EQ_COUNT_COL)))
    digits = Len(CStr(maxNumber))
    ws.Range(ws.Cells(2, EQ_COUNT_COL), ws.Cells(lastRow, EQ_COUNT_COL)).NumberFormat = String(digits, "0")
End Sub

' Column index of the "PURCHASED" heading in row 1, or 0 if there isn't one.
Private Function FindPurchasedColumn(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If UCase$(Trim$(CStr(ws.Cells(1, c).value))) = "PURCHASED" Then
            FindPurchasedColumn = c
            Exit Function
        End If
    Next c
End Function
