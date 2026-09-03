Attribute VB_Name = "modHelpers_Strings"
'@Folder("TPD_Addin.Helpers")

Option Explicit

' Makes a string usable as an Excel sheet name: strips characters Excel
' forbids, trims, falls back to "Sheet" when empty, and truncates to 31.
' KNOWN LIMITATION: the 31-char truncation happens here, before any caller
' appends a " (2)" disambiguator, and two distinct source values that share
' the first 31 characters collapse to the same name - in the Split Sheet
' flow that means the second value's rows overwrite the first. See issue #37.
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

' Normalizes a cell value to a single comparable line of text: returns "" for
' error/null/empty, otherwise the first physical line only (splits on CR/LF),
' with non-breaking spaces converted to normal spaces, runs of spaces collapsed
' to one, and the result trimmed. Used wherever cell values are compared or
' deduped (Split Sheet group values, unique-value lists).
Public Function NormalizeCellText(val As Variant) As String
    Dim s As String
    On Error Resume Next

    If IsError(val) Or IsNull(val) Or IsEmpty(val) Then
        NormalizeCellText = ""
        Exit Function
    End If

    s = CStr(val)

    ' Split on line breaks - keep only first line
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

    NormalizeCellText = Trim(s)
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


