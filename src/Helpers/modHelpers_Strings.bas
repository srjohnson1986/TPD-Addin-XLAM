Attribute VB_Name = "modHelpers_Strings"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Function SanitizeSheetName(ByVal name As String) As String
    Dim badChars As Variant
    Dim c As Variant

    badChars = Array("/", "\", "?", "*", "[", "]", ":", "'", Chr(34))

    For Each c In badChars
        name = Replace(name, c, "_")
    Next c

    name = Trim(name)
    If Len(name) = 0 Then name = "Sheet"
    If Len(name) > 31 Then name = Left$(name, 31)

    SanitizeSheetName = name
End Function

Public Function SanitizeFileText(ByVal txt As String) As String
    Dim badChars As Variant
    Dim c As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each c In badChars
        txt = Replace(txt, c, "_")
    Next c

    txt = Trim(txt)
    SanitizeFileText = txt
End Function

Public Function CleanValue(val As Variant) As String
    Dim s As String
    On Error Resume Next

    If IsError(val) Or IsNull(val) Or IsEmpty(val) Then
        CleanValue = ""
        Exit Function
    End If

    s = CStr(val)

    ' Split on line breaks — keep only first line
    If InStr(s, vbLf) > 0 Then
        s = Split(s, vbLf)(0)
    End If
    If InStr(s, vbCr) > 0 Then
        s = Split(s, vbCr)(0)
    End If

    ' Remove non-breaking spaces
    s = Replace(s, Chr(160), " ")

    ' Collapse multiple spaces
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanValue = Trim(s)
End Function

Public Function IsInArray(val As String, arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If StrComp(Trim(val), Trim(v), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next v
End Function


Public Function CleanHeader(s As String) As String
    If Len(s) = 0 Then
        CleanHeader = ""
        Exit Function
    End If

    ' Normalize NBSP and Unicode whitespace
    s = Replace(s, Chr(160), " ")
    s = Replace(s, ChrW(8203), "")
    s = Replace(s, ChrW(8237), "")
    s = Replace(s, ChrW(8236), "")
    s = Replace(s, ChrW(9), " ")
    s = Replace(s, ChrW(10), " ")
    s = Replace(s, ChrW(13), " ")

    ' Collapse multiple spaces
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanHeader = Trim(s)
End Function



