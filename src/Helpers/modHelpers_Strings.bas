Attribute VB_Name = "modHelpers_Strings"
'@Folder("TPD_Addin.Helpers")

Option Explicit

' Makes a string usable as an Excel sheet name: strips characters Excel
' forbids, trims, falls back to "Sheet" when empty, and truncates to 31.
' Two distinct source values can still collapse to the same name (shared
' first 31 chars, or differing only in stripped characters) - callers that
' create one sheet per value must dedupe the result themselves. The Split
' Sheet flow does this via UniqueRunName (#57).
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


