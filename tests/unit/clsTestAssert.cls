VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TPD_Addin.Tests")

'===========================================================
'  The assert object every test module uses (get one from
'  modTestSupport.NewTestAssert).
'
'  Why not Rubberduck's own: its Assert objects only do
'  anything while Rubberduck's Test Explorer is driving the
'  run. Called any other way - Application.Run, a headless
'  COM run - AreEqual and Fail are silent no-ops, so a test
'  "passes" whatever it asserted.
'
'  This class makes the comparison itself, so the verdict is
'  the same wherever the test runs, and reports a failure to
'  two places:
'    - modTestSupport.RecordAssertFailure, which is what the
'      headless runner (tools\Test-Unit.ps1) reads;
'    - Rubberduck's Fail, when Rubberduck is installed, so
'      Test Explorer shows the same failure. Outside Test
'      Explorer that call does nothing, which is fine.
'
'  Not part of /src - never imported by BuildAddin, never
'  shipped in the .xlam.
'===========================================================

Option Explicit

Private mRubberduck As Object   ' Nothing when Rubberduck isn't installed

Private Sub Class_Initialize()
    On Error Resume Next
    Set mRubberduck = CreateObject("Rubberduck.PermissiveAssertClass")
    On Error GoTo 0
End Sub

Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, _
                    Optional ByVal message As String = vbNullString)
    If Not ValuesEqual(expected, actual) Then
        Report "Expected " & Describe(expected) & " but was " & Describe(actual) & Suffix(message)
    End If
End Sub

Public Sub AreNotEqual(ByVal notExpected As Variant, ByVal actual As Variant, _
                       Optional ByVal message As String = vbNullString)
    If ValuesEqual(notExpected, actual) Then
        Report "Expected anything but " & Describe(notExpected) & Suffix(message)
    End If
End Sub

Public Sub IsTrue(ByVal condition As Boolean, Optional ByVal message As String = vbNullString)
    If Not condition Then Report "Expected True but was False" & Suffix(message)
End Sub

Public Sub IsFalse(ByVal condition As Boolean, Optional ByVal message As String = vbNullString)
    If condition Then Report "Expected False but was True" & Suffix(message)
End Sub

Public Sub Fail(Optional ByVal message As String = vbNullString)
    If Len(message) = 0 Then message = "Fail was called"
    Report message
End Sub

' Equality the way the tests were written for Rubberduck's permissive assert:
' value equality with VBA's usual numeric coercion (3 = 3& = 3#), text compared
' exactly (case-sensitive), and a number against its text form ("3" vs 3)
' treated as equal. Objects are equal only when they are the same object.
' Arrays are never equal - assert on their elements.
Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    On Error GoTo NotComparable

    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then ValuesEqual = (a Is b)
        Exit Function
    End If
    If IsArray(a) Or IsArray(b) Then Exit Function
    If IsNull(a) Or IsNull(b) Then
        ValuesEqual = (IsNull(a) And IsNull(b))
        Exit Function
    End If

    If (VarType(a) = vbString) Xor (VarType(b) = vbString) Then
        ValuesEqual = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
    ElseIf VarType(a) = vbString Then
        ValuesEqual = (StrComp(a, b, vbBinaryCompare) = 0)
    Else
        ValuesEqual = (a = b)
    End If
    Exit Function

NotComparable:
    ValuesEqual = False
End Function

Private Function Describe(ByVal v As Variant) As String
    On Error GoTo Unprintable

    If IsObject(v) Then
        If v Is Nothing Then Describe = "Nothing" Else Describe = "<" & TypeName(v) & " object>"
    ElseIf IsArray(v) Then
        Describe = "<" & TypeName(v) & " array>"
    ElseIf IsNull(v) Then
        Describe = "Null"
    ElseIf IsEmpty(v) Then
        Describe = "Empty"
    ElseIf VarType(v) = vbString Then
        Describe = """" & v & """"
    Else
        Describe = CStr(v) & " (" & TypeName(v) & ")"
    End If
    Exit Function

Unprintable:
    Describe = "<" & TypeName(v) & ">"
End Function

Private Function Suffix(ByVal message As String) As String
    If Len(message) > 0 Then Suffix = " - " & message
End Function

Private Sub Report(ByVal message As String)
    modTestSupport.RecordAssertFailure message

    If Not mRubberduck Is Nothing Then
        On Error Resume Next
        mRubberduck.Fail message
        On Error GoTo 0
    End If
End Sub
