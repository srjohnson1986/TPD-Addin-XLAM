Attribute VB_Name = "modHelpers_Workbook"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Function GetFirstVisibleSheet() As Worksheet
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Visible = xlSheetVisible Then
            Set GetFirstVisibleSheet = sh
            Exit Function
        End If
    Next sh
End Function

Public Function CreateNewSheetAfterLast(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = sheetName
    Set CreateNewSheetAfterLast = ws
End Function

Public Function CreateOrClearSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet

    ' Try to get existing sheet EXACTLY by name
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new worksheet (not chart sheet)
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set CreateOrClearSheet = ws
End Function

Public Function IsWorkbookOpen(fileName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks(fileName)
    On Error GoTo 0
    IsWorkbookOpen = Not wb Is Nothing
End Function

Public Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Public Function GetNextAvailableSheetName(baseName As String) As String
    Dim nameToTry As String
    Dim n As Long

    nameToTry = baseName
    n = 1

    Do While SheetExists(ActiveWorkbook, nameToTry)
        n = n + 1
        nameToTry = baseName & " (" & n & ")"
    Loop

    GetNextAvailableSheetName = nameToTry
End Function

Function GetLastRow(ws As Worksheet) As Long
    GetLastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
End Function


Public Function GetLastCol(ws As Worksheet, headingsRow As Long) As Long
    GetLastCol = ws.Cells(headingsRow, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Sub DeleteSheetIfExists(wsName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(wsName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

