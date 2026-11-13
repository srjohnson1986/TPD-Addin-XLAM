Attribute VB_Name = "modTestSupport"
'@Folder("TPD_Addin.Tests")

'===========================================================
'  Support for the unit tests in this folder: where a test
'  module gets its assert object, and where a failed assert
'  is recorded.
'
'  Two ways to run the tests, same verdicts either way:
'    - Rubberduck > Unit Tests > Test Explorer > Run All, in
'      a workbook holding /src plus this folder;
'    - tools\Test-Unit.ps1, which opens the built add-in over
'      COM, imports this folder plus a runner module it
'      generates, and reads the failures recorded here.
'
'  Not part of /src - never imported by BuildAddin, never
'  shipped in the .xlam.
'===========================================================

Option Explicit
Option Private Module

' Assert failures recorded since ResetAssertFailures was last called.
Private mFailures As Collection

' Notes recorded by the test that is running - see RecordTestNote.
Private mNotes As Collection

' Every test module's ModuleInitialize calls this.
Public Function NewTestAssert() As clsTestAssert
    Set NewTestAssert = New clsTestAssert
End Function

' Called by clsTestAssert each time an assert does not hold, and by the
' headless runner when a test raises an error.
Public Sub RecordAssertFailure(ByVal message As String)
    If mFailures Is Nothing Then Set mFailures = New Collection
    mFailures.Add message
End Sub

' The headless runner calls this before each test.
Public Sub ResetAssertFailures()
    Set mFailures = New Collection
    Set mNotes = New Collection
End Sub

' Something a test saw that is worth reporting but is not a pass or a fail -
' for example, where the add-in put a column when nothing says where it should
' go. The headless runner prints notes under the results; in Test Explorer
' they are simply dropped.
Public Sub RecordTestNote(ByVal message As String)
    If mNotes Is Nothing Then Set mNotes = New Collection
    mNotes.Add message
End Sub

' Every note recorded since the last reset, joined with "; ".
Public Function TestNoteText() As String
    Dim entry As Variant
    Dim joined As String

    If mNotes Is Nothing Then Exit Function
    For Each entry In mNotes
        If Len(joined) > 0 Then joined = joined & "; "
        joined = joined & CStr(entry)
    Next entry

    TestNoteText = joined
End Function

Public Function AssertFailureCount() As Long
    If Not mFailures Is Nothing Then AssertFailureCount = mFailures.Count
End Function

' Every recorded failure, joined with "; ".
Public Function AssertFailureText() As String
    Dim entry As Variant
    Dim joined As String

    If mFailures Is Nothing Then Exit Function
    For Each entry In mFailures
        If Len(joined) > 0 Then joined = joined & "; "
        joined = joined & CStr(entry)
    Next entry

    AssertFailureText = joined
End Function
