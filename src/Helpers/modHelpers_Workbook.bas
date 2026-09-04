Attribute VB_Name = "modHelpers_Workbook"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Sheet / workbook utilities shared across feature areas:
'  first visible sheet, create-or-clear a sheet by name, next
'  available sheet name, last-row / last-column lookups, and
'  open-workbook / sheet-exists checks.
'===========================================================

Option Explicit

' Sets wb to ActiveWorkbook and returns True; if there is no active workbook
' it shows a message and returns False. Use as the standard guard at the top
' of any flow that needs a workbook:  If Not RequireActiveWorkbook(wb) Then Exit Sub
Public Function RequireActiveWorkbook(ByRef wb As Workbook) As Boolean
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "TPD Add-in"
        Exit Function
    End If
    RequireActiveWorkbook = True
End Function

' Returns the first visible worksheet of the active workbook. Returns Nothing
' (after showing a message) when there is no active workbook or it has no
' visible worksheet - callers must check for Nothing and bail.
Public Function GetFirstVisibleSheet() As Worksheet
    Dim wb As Workbook
    Dim sh As Worksheet

    If Not RequireActiveWorkbook(wb) Then Exit Function

    For Each sh In wb.Worksheets
        If sh.Visible = xlSheetVisible Then
            Set GetFirstVisibleSheet = sh
            Exit Function
        End If
    Next sh

    MsgBox "This workbook has no visible worksheet to work from.", _
           vbExclamation, "TPD Add-in"
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

' Last row with content anywhere on the sheet (any column). Use this for
' "end of the data table"; use .End(xlUp) on a specific column only when you
' deliberately mean that one column (#89).
Public Function GetLastRow(ws As Worksheet) As Long
    ' All Find arguments are pinned: the unspecified ones otherwise inherit
    ' whatever the last Find call - or the user's Find & Replace dialog - left
    ' set, so results would vary between runs.
    ' Find returns Nothing on a completely empty sheet - treat that as row 1
    ' rather than letting .Row raise error 91 (see issue #33).
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                              MatchCase:=False)
    If found Is Nothing Then
        GetLastRow = 1
    Else
        GetLastRow = found.Row
    End If
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

