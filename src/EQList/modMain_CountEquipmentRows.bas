Attribute VB_Name = "modMain_CountEquipmentRows"
'@Folder("TPD_Addin.EQList")

Option Explicit

Sub CountNonParentOrIncludedRows(ws As Worksheet)

    Dim wb As Workbook
    
    
    ' Ensure an active workbook exists
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "No active workbook detected. Please select a workbook and try again.", vbCritical
        Exit Sub
    End If
    
    ' Ensure the active sheet is a worksheet
    On Error Resume Next
    Set ws = wb.ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Or Not TypeOf ws Is Worksheet Then
        MsgBox "The active sheet is not a worksheet. Please select a worksheet and try again.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    Dim purchasedCol As Long
    Dim numberCol As Long
    Dim i As Long
    Dim counter As Long
    
    ' Find PURCHASED column
    purchasedCol = 0
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(UCase(ws.Cells(1, i).value)) = "PURCHASED" Then
            purchasedCol = i
            Exit For
        End If
    Next i
    
    If purchasedCol = 0 Then
        MsgBox "No column with heading PURCHASED was found.", vbCritical
        Exit Sub
    End If
    
    ' Insert EQ COUNT as the leftmost column
    ws.Columns(1).Insert Shift:=xlToRight
    ws.Cells(1, 1).value = "EQ COUNT"
    numberCol = 1

    
    ' Determine last row
    lastRow = ws.Cells(ws.Rows.Count, purchasedCol + 1).End(xlUp).Row
    
    counter = 1
    
    ' Number rows except PARENT or INCLUDED
    For i = 2 To lastRow
        Dim val As String
        val = UCase(Trim(ws.Cells(i, purchasedCol + 1).value))
        
        If val <> "PARENT" And val <> "INCLUDED" Then
            ws.Cells(i, numberCol).value = counter
            counter = counter + 1
        Else
            ws.Cells(i, numberCol).value = ""
        End If
    Next i
    
    
    ' -------------------------------
    ' Leading-zero display formatting
    ' -------------------------------
    
    Dim maxCount As Long
    Dim digits As Long
    
    maxCount = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, numberCol), ws.Cells(lastRow, numberCol)))
    digits = Len(CStr(maxCount))
    
    ws.Range(ws.Cells(2, numberCol), ws.Cells(lastRow, numberCol)).NumberFormat = String(digits, "0")

End Sub


Public Sub RunCountEquipmentRows(control As IRibbonControl)
    CountNonParentOrIncludedRows ActiveSheet
End Sub

