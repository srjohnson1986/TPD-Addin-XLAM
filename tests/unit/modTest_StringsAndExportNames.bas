Attribute VB_Name = "modTest_StringsAndExportNames"
'@Folder("TPD_Addin.Tests")
'@TestModule

'===========================================================
'  Unit tests for the naming rules: modHelpers_Strings
'  (SanitizeSheetName, SanitizeFileText, NormalizeCellText),
'  modMain_SplitSheets.UniqueRunName (sheet-name collisions,
'  #57) and modExport_Files (BuildExportSuffix,
'  PreviewExportFileName).
'
'  The SanitizeSheetName / UniqueRunName cases use the Vendor
'  values in tests/fixtures/EQ_List_Sample.xlsx, so the
'  "expected sheet name" table in tests/README.md is checked
'  here instead of by eye.
'
'  Not part of /src - never imported by BuildAddin, never
'  shipped in the .xlam. How to run: tests/README.md.
'===========================================================

Option Explicit
Option Private Module

Private Assert As clsTestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = NewTestAssert()
End Sub

' The same kind of dictionary SplitSheetByColumn_DoWork tracks used names in.
Private Function NewUsedNames() As Object
    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    Set NewUsedNames = names
End Function

' ---- SanitizeSheetName ----------------------------------------------------

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_IllegalCharacters_BecomeUnderscores()
    On Error GoTo TestFail

    ' 32 characters in; the trailing "?" becomes "_" and is then cut at 31.
    Assert.AreEqual "Acme Valve_Fittings _West__ Ltd", _
                    modHelpers_Strings.SanitizeSheetName("Acme Valve/Fittings [West]: Ltd?")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_QuotesAndApostrophes_BecomeUnderscores()
    On Error GoTo TestFail

    Assert.AreEqual "O_Brien _Co_", modHelpers_Strings.SanitizeSheetName("O'Brien ""Co""")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_LongName_CutAt31()
    On Error GoTo TestFail

    Dim result As String
    result = modHelpers_Strings.SanitizeSheetName("International Combustion & Controls Corporation of America")

    Assert.AreEqual "International Combustion & Cont", result
    Assert.AreEqual 31, Len(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_TwoLongNames_CollapseToTheSameName()
    On Error GoTo TestFail

    ' Why UniqueRunName exists: these differ only after character 31.
    Assert.AreEqual modHelpers_Strings.SanitizeSheetName("International Combustion & Controls Corporation of America"), _
                    modHelpers_Strings.SanitizeSheetName("International Combustion & Controls Corporation of Asia")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_SurroundingSpaces_Trimmed()
    On Error GoTo TestFail

    Assert.AreEqual "Alder Valve Co", modHelpers_Strings.SanitizeSheetName("  Alder Valve Co  ")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeSheetName")
Public Sub SanitizeSheetName_EmptyOrBlank_FallsBackToSheet()
    On Error GoTo TestFail

    Assert.AreEqual "Sheet", modHelpers_Strings.SanitizeSheetName("")
    Assert.AreEqual "Sheet", modHelpers_Strings.SanitizeSheetName("   ")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' ---- UniqueRunName --------------------------------------------------------

'@TestMethod("UniqueRunName")
Public Sub UniqueRunName_UnusedName_ReturnedAsIs()
    On Error GoTo TestFail

    Dim used As Object
    Set used = NewUsedNames()
    used.Add "Alder Valve Co", True

    Assert.AreEqual "Birch Pumps", modMain_SplitSheets.UniqueRunName("Birch Pumps", used)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueRunName")
Public Sub UniqueRunName_UsedShortName_GetsSuffix2()
    On Error GoTo TestFail

    Dim used As Object
    Set used = NewUsedNames()
    used.Add "Alder Valve Co", True

    Assert.AreEqual "Alder Valve Co (2)", modMain_SplitSheets.UniqueRunName("Alder Valve Co", used)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueRunName")
Public Sub UniqueRunName_Used31CharName_TrimmedToFitSuffix()
    On Error GoTo TestFail

    Dim used As Object
    Dim result As String
    Set used = NewUsedNames()
    used.Add "International Combustion & Cont", True

    result = modMain_SplitSheets.UniqueRunName("International Combustion & Cont", used)

    ' The base is cut to 27 characters, ending in a space, then " (2)" is added:
    ' hence the two spaces. This is the "...of Asia" row in tests/README.md.
    Assert.AreEqual "International Combustion &  (2)", result
    Assert.AreEqual 31, Len(result)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueRunName")
Public Sub UniqueRunName_Suffix2AlsoUsed_GetsSuffix3()
    On Error GoTo TestFail

    Dim used As Object
    Set used = NewUsedNames()
    used.Add "Alder Valve Co", True
    used.Add "Alder Valve Co (2)", True

    Assert.AreEqual "Alder Valve Co (3)", modMain_SplitSheets.UniqueRunName("Alder Valve Co", used)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("UniqueRunName")
Public Sub UniqueRunName_NameDifferingOnlyInCase_CountsAsUsed()
    On Error GoTo TestFail

    ' Excel sheet names are case-insensitive, and so is the dictionary.
    Dim used As Object
    Set used = NewUsedNames()
    used.Add "Alder Valve Co", True

    Assert.AreEqual "ALDER VALVE CO (2)", modMain_SplitSheets.UniqueRunName("ALDER VALVE CO", used)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' ---- SanitizeFileText -----------------------------------------------------

'@TestMethod("SanitizeFileText")
Public Sub SanitizeFileText_IllegalCharacters_BecomeUnderscores()
    On Error GoTo TestFail

    Assert.AreEqual "a_b_c_d_e_f_g_h_i_j", modHelpers_Strings.SanitizeFileText("a\b/c:d*e?f""g<h>i|j")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SanitizeFileText")
Public Sub SanitizeFileText_SurroundingSpaces_Trimmed()
    On Error GoTo TestFail

    Assert.AreEqual "Rev A", modHelpers_Strings.SanitizeFileText("  Rev A  ")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' ---- NormalizeCellText ----------------------------------------------------

'@TestMethod("NormalizeCellText")
Public Sub NormalizeCellText_MultiLine_KeepsFirstLineOnly()
    On Error GoTo TestFail

    Assert.AreEqual "Gate Valve", modHelpers_Strings.NormalizeCellText("Gate Valve" & vbLf & "Class 150")
    Assert.AreEqual "Gate Valve", modHelpers_Strings.NormalizeCellText("Gate Valve" & vbCrLf & "Class 150")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("NormalizeCellText")
Public Sub NormalizeCellText_SpacesAndNbsp_CollapsedAndTrimmed()
    On Error GoTo TestFail

    Assert.AreEqual "alder valve co", modHelpers_Strings.NormalizeCellText(" alder   valve" & Chr(160) & "co ")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("NormalizeCellText")
Public Sub NormalizeCellText_EmptyNullAndError_AreEmptyText()
    On Error GoTo TestFail

    Dim nothingSet As Variant

    Assert.AreEqual "", modHelpers_Strings.NormalizeCellText(nothingSet)
    Assert.AreEqual "", modHelpers_Strings.NormalizeCellText(Null)
    Assert.AreEqual "", modHelpers_Strings.NormalizeCellText(CVErr(2042))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("NormalizeCellText")
Public Sub NormalizeCellText_Number_BecomesItsText()
    On Error GoTo TestFail

    Dim tag As Variant
    tag = 1042

    Assert.AreEqual "1042", modHelpers_Strings.NormalizeCellText(tag)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' ---- BuildExportSuffix / PreviewExportFileName ----------------------------

'@TestMethod("BuildExportSuffix")
Public Sub BuildExportSuffix_TextOnly_Trimmed()
    On Error GoTo TestFail

    Assert.AreEqual "Rev A", modExport_Files.BuildExportSuffix("  Rev A  ", False)
    Assert.AreEqual "", modExport_Files.BuildExportSuffix("   ", False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("BuildExportSuffix")
Public Sub BuildExportSuffix_WithDate_AppendsTodaysDate()
    On Error GoTo TestFail

    Dim today As String
    today = modHelpers_SheetSetup.GetTodaysDate()

    Assert.AreEqual today, modExport_Files.BuildExportSuffix("", True)
    Assert.AreEqual "Rev A - " & today, modExport_Files.BuildExportSuffix("Rev A", True)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PreviewExportFileName")
Public Sub PreviewExportFileName_SheetAndSuffix_Joined()
    On Error GoTo TestFail

    Assert.AreEqual "EQ List - Rev A.xlsx", modExport_Files.PreviewExportFileName("EQ List", "Rev A", False)
    Assert.AreEqual "EQ List.xlsx", modExport_Files.PreviewExportFileName("EQ List", "", False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PreviewExportFileName")
Public Sub PreviewExportFileName_NoSheetName_UsesSheet1()
    On Error GoTo TestFail

    Assert.AreEqual "Sheet1.xlsx", modExport_Files.PreviewExportFileName("", "", False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("PreviewExportFileName")
Public Sub PreviewExportFileName_IllegalCharacters_Sanitized()
    On Error GoTo TestFail

    Assert.AreEqual "A_B - x_y.xlsx", modExport_Files.PreviewExportFileName("A/B", "x:y", False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
