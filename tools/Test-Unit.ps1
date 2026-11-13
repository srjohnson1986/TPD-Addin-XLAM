<#
.SYNOPSIS
    Runs the unit tests in /tests/unit headlessly against the built add-in and
    exits non-zero if any fail.

.DESCRIPTION
    Opens build\TPD_Addin.xlam read-only in a hidden Excel started for this
    run, imports every module in /tests/unit plus a small runner module
    generated for the run (in memory only - the file on disk is never saved),
    runs each '@TestMethod sub, and prints one line per failure.

    Needs Excel and "Trust access to the VBA project object model", like the
    build - so this is a local check, not a CI job.

    HOW A TEST IS CALLED. The generated runner calls every test by name in
    ordinary VBA ("modTest_X.Some_Test" inside a Select Case), under an
    On Error handler, so an error raised inside a test is that test's failure.
    It deliberately does not use Application.Run with the test's name as a
    string. The first version of this script did, and its very first run hung
    on a "Microsoft Visual Basic" box in the hidden Excel. A direct call has
    ordinary VBA error handling and nothing to guess about.

    The tests use tests\unit\clsTestAssert.cls, which makes its own
    comparison, so a failed assert is a real failure here. Rubberduck's own
    Assert objects are silent no-ops outside Rubberduck's Test Explorer.

    RUNNER SELF-CHECK. Every run also imports a three-test canary module: one
    test that must pass, one that must fail on an assert, one that must fail
    by raising an error. If the runner reports anything else for those three,
    the run is thrown out (exit 2) - a runner that cannot report a failure
    proves nothing by reporting none.

    WATCHDOG. The Excel work runs in a child PowerShell process. If it has not
    finished after -TimeoutSeconds, Excel is sitting on a dialog nobody can
    see (a compile error, a MsgBox, or the VBE in break mode). The script then
    prints which step it was on and the text of any dialog that Excel is
    showing, ends the Excel process IT started - identified by window handle,
    never any other Excel - and exits 3. Nothing is retried.

    Exit codes: 0 all passed, 1 test failures, 2 could not run or runner
    self-check failed, 3 timed out.

.PARAMETER Build
    Run tools\Build-TPDAddin.ps1 first. Without it, the script refuses to run
    against a build that is older than anything in /src or missing.

.PARAMETER PlanOnly
    Print the tests that would run and the generated runner module, and stop.
    Does not start Excel and does not need a build.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Test-Unit.ps1
    powershell -ExecutionPolicy Bypass -File tools\Test-Unit.ps1 -Build
#>

[CmdletBinding()]
param(
    [string]$XlamPath,

    # Which tests to run. 'unit' is tests\unit: pure functions, no workbook.
    # 'flows' is tests\flows: whole flows run against copies of the fixture
    # workbooks in tests\fixtures. Flow tests use copy and paste, so they
    # overwrite the clipboard while they run.
    [ValidateSet('unit', 'flows')][string]$Suite = 'unit',

    # Default: 120 for unit tests, 300 for flow tests.
    [int]$TimeoutSeconds = 0,
    [switch]$Build,
    [switch]$PlanOnly,

    # Internal: set when the script re-launches itself to do the Excel work.
    [switch]$Worker,
    [string]$WorkDir,

    # Internal: print what the watchdog would report for this process id.
    [int]$ProbeWindowsOfPid
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$unitDir  = Join-Path (Join-Path $repoRoot 'tests') 'unit'
$srcDir   = Join-Path $repoRoot 'src'
if ([string]::IsNullOrEmpty($XlamPath)) {
    $XlamPath = Join-Path (Join-Path $repoRoot 'build') 'TPD_Addin.xlam'
}

$flowsDir   = Join-Path (Join-Path $repoRoot 'tests') 'flows'
$fixtureDir = Join-Path (Join-Path $repoRoot 'tests') 'fixtures'

$suiteLabel = 'unit'
if ($Suite -eq 'flows') { $suiteLabel = 'flow' }
if ($TimeoutSeconds -le 0) {
    $TimeoutSeconds = 120
    if ($Suite -eq 'flows') { $TimeoutSeconds = 300 }
}

$canaryModule = 'modTest_RunnerCanary'
$runnerModule = 'modTestDispatch'

# Win32 helpers: which process owns a window, and what text a process's
# visible windows and dialogs are showing. Read-only.
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace TpdTest
{
    public static class Win32
    {
        private delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumProc callback, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, StringBuilder lParam,
                                                        uint flags, uint timeoutMs, out IntPtr result);

        public static uint ProcessIdOf(IntPtr hWnd)
        {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            return pid;
        }

        private static string TextOf(IntPtr hWnd)
        {
            StringBuilder text = new StringBuilder(2048);
            IntPtr ignored;
            // WM_GETTEXT, SMTO_ABORTIFHUNG, one second
            SendMessageTimeout(hWnd, 0x000D, (IntPtr)text.Capacity, text, 0x0002, 1000, out ignored);
            return text.ToString().Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string ClassOf(IntPtr hWnd)
        {
            StringBuilder text = new StringBuilder(256);
            GetClassName(hWnd, text, text.Capacity);
            return text.ToString();
        }

        // One line per visible top-level window of the process; for a dialog
        // box, one more line per piece of text on it.
        public static string[] VisibleWindowText(uint processId)
        {
            List<string> lines = new List<string>();
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint owner;
                GetWindowThreadProcessId(hWnd, out owner);
                if (owner != processId || !IsWindowVisible(hWnd)) return true;

                string cls = ClassOf(hWnd);
                lines.Add((cls == "#32770" ? "dialog: " : "window: ") + TextOf(hWnd));
                if (cls == "#32770")
                {
                    EnumChildWindows(hWnd, delegate(IntPtr child, IntPtr unused2)
                    {
                        string text = TextOf(child);
                        if (text.Length > 0) lines.Add("    " + ClassOf(child) + ": " + text);
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            return lines.ToArray();
        }
    }
}
'@

if ($ProbeWindowsOfPid) {
    [TpdTest.Win32]::VisibleWindowText([uint32]$ProbeWindowsOfPid) | ForEach-Object { Write-Host $_ }
    exit 0
}

# ===========================================================================
# Worker: the only part that touches Excel. Reads imports.txt from $WorkDir,
# writes excel.pid as soon as Excel exists and result.txt (or error.txt) when
# done. The VBA runner appends to progress.log as it goes.
# ===========================================================================
if ($Worker) {
    $excel = $null
    $wb    = $null
    try {
        $imports = @(Get-Content -LiteralPath (Join-Path $WorkDir 'imports.txt') | Where-Object { $_ })

        $excel = New-Object -ComObject Excel.Application

        # Record which process is ours before doing anything that could hang,
        # so the parent's watchdog can end this Excel and no other.
        $excelPid = [TpdTest.Win32]::ProcessIdOf([IntPtr]$excel.Hwnd)
        [System.IO.File]::WriteAllText((Join-Path $WorkDir 'excel.pid'), [string]$excelPid)

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1    # msoAutomationSecurityLow, this process only

        # Open(Filename, UpdateLinks, ReadOnly) - read-only, so a copy the user
        # has loaded elsewhere is not disturbed and nothing can be saved back.
        $wb = $excel.Workbooks.Open($XlamPath, 0, $true)

        foreach ($file in $imports) {
            [void]$wb.VBProject.VBComponents.Import($file)
        }

        $result = $excel.Run("'" + $wb.Name + "'!$runnerModule.RunAllTestsHeadless", (Join-Path $WorkDir 'progress.log'))
        [System.IO.File]::WriteAllText((Join-Path $WorkDir 'result.txt'), [string]$result)
    }
    catch {
        [System.IO.File]::WriteAllText((Join-Path $WorkDir 'error.txt'), $_.Exception.Message)
    }
    finally {
        if ($wb)    { try { $wb.Close($false) } catch { } }
        if ($excel) {
            try { $excel.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        Remove-Variable excel, wb -ErrorAction SilentlyContinue
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    exit 0
}

# ===========================================================================
# Parent
# ===========================================================================

function Stop-WithMessage {
    param([string]$Message, [int]$Code)
    Write-Host $Message -ForegroundColor Red
    exit $Code
}

if (-not (Test-Path -LiteralPath $unitDir)) { Stop-WithMessage "No tests folder at $unitDir" 2 }

# ---- 1. A build that matches /src ------------------------------------------

if (-not $PlanOnly) {
    if ($Build) {
        Write-Host 'Building the add-in first...'
        # A child process, so the build script's own "exit" cannot end this one.
        $buildProc = Start-Process -FilePath (Get-Process -Id $PID).Path -Wait -PassThru -NoNewWindow -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $PSScriptRoot 'Build-TPDAddin.ps1')`"")
        if ($buildProc.ExitCode -ne 0) { Stop-WithMessage "Build failed (exit $($buildProc.ExitCode)) - not running tests." 2 }
    }

    if (-not (Test-Path -LiteralPath $XlamPath)) {
        Stop-WithMessage "Built add-in not found: $XlamPath`nRun tools\Build-TPDAddin.ps1, or pass -Build." 2
    }
    $XlamPath = (Resolve-Path -LiteralPath $XlamPath).Path

    $builtAt  = (Get-Item -LiteralPath $XlamPath).LastWriteTime
    $newerSrc = @(Get-ChildItem -Path $srcDir -Recurse -Include *.bas, *.cls, *.frm, *.frx |
                    Where-Object { $_.LastWriteTime -gt $builtAt } | Sort-Object LastWriteTime -Descending)
    if ($newerSrc.Count -gt 0) {
        Stop-WithMessage ("The build is older than /src ($($newerSrc.Count) file(s) changed since, e.g. $($newerSrc[0].Name)).`n" +
                          'Run tools\Build-TPDAddin.ps1, or pass -Build, so the tests run against the current code.') 2
    }
}

# ---- 2. Find the tests by their annotations --------------------------------

# What gets scanned for tests, and what is only imported alongside.
$testDir      = $unitDir
$supportFiles = @()
if ($Suite -eq 'flows') {
    $testDir      = $flowsDir
    $supportFiles = @('clsTestAssert.cls', 'modTestSupport.bas' | ForEach-Object { Get-Item -LiteralPath (Join-Path $unitDir $_) })
}
if (-not (Test-Path -LiteralPath $testDir)) { Stop-WithMessage "No tests folder at $testDir" 2 }
$testFiles = @(Get-ChildItem -Path $testDir -File | Where-Object { $_.Extension -in '.bas', '.cls' } | Sort-Object Name)

$callLines = New-Object System.Collections.Generic.List[string]   # body of RunAllTestsHeadless
$caseLines = New-Object System.Collections.Generic.List[string]   # Select Case in CallProc
$expected  = New-Object System.Collections.Generic.List[string]   # "module.proc" for every real test
$script:nextId = 0

# Gives a sub an id and a "Case id: module.sub" line in the generated runner.
function New-ProcId {
    param([string]$Module, [string]$Proc)
    if ([string]::IsNullOrEmpty($Proc)) { return 0 }
    $script:nextId++
    $script:caseLines.Add("        Case $($script:nextId): $Module.$Proc")
    return $script:nextId
}

function Add-ModuleToPlan {
    param([string]$Path)

    $text = [System.IO.File]::ReadAllText($Path)
    if ($text -notmatch "(?m)^'@TestModule\b") { return 0 }

    $nameMatch = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"')
    if (-not $nameMatch.Success) { throw "No Attribute VB_Name line in $Path" }
    $module = $nameMatch.Groups[1].Value

    $hooks = @{ ModuleInitialize = ''; ModuleCleanup = ''; TestInitialize = ''; TestCleanup = '' }
    $tests = New-Object System.Collections.Generic.List[string]
    $pending = $null

    foreach ($line in [regex]::Split($text, "\r\n|\n|\r")) {
        $annotation = [regex]::Match($line, "^\s*'@(TestMethod|ModuleInitialize|ModuleCleanup|TestInitialize|TestCleanup)\b")
        if ($annotation.Success) { $pending = $annotation.Groups[1].Value; continue }
        if ($null -eq $pending) { continue }

        $sub = [regex]::Match($line, '^\s*(?:Public\s+|Private\s+)?Sub\s+([A-Za-z_]\w*)')
        if ($sub.Success) {
            if ($pending -eq 'TestMethod') { $tests.Add($sub.Groups[1].Value) } else { $hooks[$pending] = $sub.Groups[1].Value }
            $pending = $null
        }
    }

    $initId    = New-ProcId $module $hooks.TestInitialize
    $cleanupId = New-ProcId $module $hooks.TestCleanup

    if ($hooks.ModuleInitialize) {
        $id = New-ProcId $module $hooks.ModuleInitialize
        $script:callLines.Add("    report = report & RunHook($id, `"$module.$($hooks.ModuleInitialize)`")")
    }
    foreach ($t in $tests) {
        $id = New-ProcId $module $t
        $script:callLines.Add("    report = report & RunTest($id, `"$module.$t`", $initId, $cleanupId)")
        if ($module -ne $script:canaryModule) { $script:expected.Add("$module.$t") }
    }
    if ($hooks.ModuleCleanup) {
        $id = New-ProcId $module $hooks.ModuleCleanup
        $script:callLines.Add("    report = report & RunHook($id, `"$module.$($hooks.ModuleCleanup)`")")
    }
    return $tests.Count
}

# ---- 3. Work folder, canary module, generated runner -----------------------

$work = Join-Path ([System.IO.Path]::GetTempPath()) ('TPD_UnitRun_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null

$canaryPath = Join-Path $work ($canaryModule + '.bas')
# The annotation lines are assembled from pieces: a line starting with '@ would
# end a here-string.
$q = "'"
$canaryText = @(
    "Attribute VB_Name = `"$canaryModule`"",
    "$q@TestModule",
    'Option Explicit',
    'Option Private Module',
    '',
    'Private Assert As clsTestAssert',
    '',
    "$q@ModuleInitialize",
    'Public Sub ModuleInitialize()',
    '    Set Assert = NewTestAssert()',
    'End Sub',
    '',
    "$q@TestMethod(`"Canary`")",
    'Public Sub Canary_MustPass()',
    '    Assert.AreEqual 3, 1 + 2',
    '    Assert.IsTrue True',
    'End Sub',
    '',
    "$q@TestMethod(`"Canary`")",
    'Public Sub Canary_MustFailOnAssert()',
    '    Assert.AreEqual "expected", "actual"',
    'End Sub',
    '',
    "$q@TestMethod(`"Canary`")",
    'Public Sub Canary_MustFailOnError()',
    '    Err.Raise 5, , "raised on purpose by the runner self-check"',
    'End Sub',
    ''
) -join "`r`n"
[System.IO.File]::WriteAllText($canaryPath, $canaryText, [System.Text.Encoding]::ASCII)

# The fixed part of the runner. {{CALLS}} and {{CASES}} are filled in below.
$runnerTemplate = @'
Attribute VB_Name = "modTestDispatch"
' Generated by tools\Test-Unit.ps1 for one headless run. Never saved.
' No "Option Private Module": RunAllTestsHeadless is called from outside the
' project.
Option Explicit

Private mLogPath As String

Public Function RunAllTestsHeadless(ByVal logPath As String) As String
    Dim report As String
    mLogPath = logPath
{{SETUP}}
{{CALLS}}
    LogStep "(all tests finished)"
    RunAllTestsHeadless = report
End Function

Private Function RunHook(ByVal procId As Long, ByVal procLabel As String) As String
    Dim problem As String
    problem = CallProc(procId, procLabel)
    If Len(problem) > 0 Then RunHook = "ERROR|" & procLabel & "|" & OneLine(problem) & vbLf
End Function

Private Function RunTest(ByVal procId As Long, ByVal procLabel As String, _
                         ByVal initId As Long, ByVal cleanupId As Long) As String
    Dim problem As String

    modTestSupport.ResetAssertFailures

    If initId <> 0 Then
        problem = CallProc(initId, procLabel & " (TestInitialize)")
        If Len(problem) > 0 Then modTestSupport.RecordAssertFailure "TestInitialize raised " & problem
    End If
    If modTestSupport.AssertFailureCount = 0 Then
        problem = CallProc(procId, procLabel)
        If Len(problem) > 0 Then modTestSupport.RecordAssertFailure "Raised " & problem
    End If
    If cleanupId <> 0 Then
        problem = CallProc(cleanupId, procLabel & " (TestCleanup)")
        If Len(problem) > 0 Then modTestSupport.RecordAssertFailure "TestCleanup raised " & problem
    End If

    If modTestSupport.AssertFailureCount = 0 Then
        RunTest = "PASS|" & procLabel & vbLf
    Else
        RunTest = "FAIL|" & procLabel & "|" & OneLine(modTestSupport.AssertFailureText) & vbLf
    End If

    If Len(modTestSupport.TestNoteText) > 0 Then
        RunTest = RunTest & "NOTE|" & procLabel & "|" & OneLine(modTestSupport.TestNoteText) & vbLf
    End If
End Function

' Calls one sub directly, by id. Returns "" when it ran to the end, otherwise
' the error it raised.
Private Function CallProc(ByVal procId As Long, ByVal procLabel As String) As String
    LogStep procLabel
    On Error GoTo Raised
    Select Case procId
{{CASES}}
        Case Else: Err.Raise 5, , "no sub registered under id " & procId
    End Select
    Exit Function
Raised:
    CallProc = "error " & Err.Number & ": " & Err.Description
End Function

' Written before each call, so a run that hangs shows where it stopped.
Private Sub LogStep(ByVal procLabel As String)
    Dim fileNo As Integer
    On Error Resume Next
    fileNo = FreeFile
    Open mLogPath For Append As #fileNo
    Print #fileNo, procLabel
    Close #fileNo
End Sub

' Results are one line each with "|" between fields.
Private Function OneLine(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    OneLine = Replace(s, "|", "/")
End Function
'@

$exitCode = 2
try {
    [void](Add-ModuleToPlan -Path $canaryPath)
    $moduleCounts = [ordered]@{}
    foreach ($f in $testFiles) {
        $n = Add-ModuleToPlan -Path $f.FullName
        if ($n -gt 0) { $moduleCounts[$f.BaseName] = $n }
    }
    if ($expected.Count -eq 0) { Stop-WithMessage "No '@TestMethod subs found under $testDir" 2 }

    # Flow tests work on copies of the fixtures, in this run's temp folder, so
    # nothing a test does can reach the tracked files.
    $setupLines = @()
    if ($Suite -eq 'flows') {
        $fixtureCopy = Join-Path $work 'fixtures'
        New-Item -ItemType Directory -Path $fixtureCopy -Force | Out-Null
        $fixtures = @(Get-ChildItem -LiteralPath $fixtureDir -Filter *.xlsx -File)
        if ($fixtures.Count -eq 0) { Stop-WithMessage "No fixture workbooks in $fixtureDir" 2 }
        $fixtures | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $fixtureCopy }
        $setupLines = @("    modFlowSupport.SetFixtureFolder `"$fixtureCopy`"")
    }

    $runnerText = $runnerTemplate.Replace('{{SETUP}}', ($setupLines -join "`r`n")).Replace('{{CALLS}}', ($callLines -join "`r`n")).Replace('{{CASES}}', ($caseLines -join "`r`n"))
    $runnerText = ($runnerText -replace "`r?`n", "`r`n") + "`r`n"
    $runnerPath = Join-Path $work ($runnerModule + '.bas')
    [System.IO.File]::WriteAllText($runnerPath, $runnerText, [System.Text.Encoding]::ASCII)

    if ($PlanOnly) {
        Write-Host "$($expected.Count) $suiteLabel test(s) in $($moduleCounts.Count) module(s), plus 3 canaries. Excel not started."
        Write-Host "---- generated $runnerModule.bas ----"
        Write-Host $runnerText
        $exitCode = 0
        return
    }

    $importList = @($supportFiles | ForEach-Object { $_.FullName }) + @($testFiles | ForEach-Object { $_.FullName }) + $canaryPath + $runnerPath
    [System.IO.File]::WriteAllLines((Join-Path $work 'imports.txt'), [string[]]$importList)

    # ---- 4. Run the worker under a watchdog --------------------------------

    $shortXlam = $XlamPath.Substring($repoRoot.Length).TrimStart('\', '/')
    $heading = (Get-Culture).TextInfo.ToTitleCase($suiteLabel)
    Write-Host "$heading tests (headless): $($expected.Count) test(s) in $($moduleCounts.Count) module(s), against $shortXlam"

    $hostExe = (Get-Process -Id $PID).Path
    $workerArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"",
                    '-Worker', '-WorkDir', "`"$work`"", '-XlamPath', "`"$XlamPath`"")
    $child = Start-Process -FilePath $hostExe -ArgumentList $workerArgs -PassThru -WindowStyle Hidden

    if (-not $child.WaitForExit($TimeoutSeconds * 1000)) {
        $exitCode = 3
        Write-Host "TIMED OUT after $TimeoutSeconds s." -ForegroundColor Red

        $progress = Join-Path $work 'progress.log'
        if (Test-Path -LiteralPath $progress) {
            $steps = @(Get-Content -LiteralPath $progress | Where-Object { $_ })
            Write-Host "  It was on step $($steps.Count): $($steps[-1])" -ForegroundColor Red
        } else {
            Write-Host '  No test had started: it stopped while opening the add-in or importing the test modules.' -ForegroundColor Red
        }

        $pidFile = Join-Path $work 'excel.pid'
        if (Test-Path -LiteralPath $pidFile) {
            $excelPid = [int](Get-Content -LiteralPath $pidFile -Raw)
            $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($proc -and $proc.ProcessName -ieq 'EXCEL') {
                $shown = @([TpdTest.Win32]::VisibleWindowText([uint32]$excelPid))
                if ($shown.Count -gt 0) {
                    Write-Host '  Excel was showing:' -ForegroundColor Red
                    $shown | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
                } else {
                    Write-Host '  Excel was showing no visible window or dialog.' -ForegroundColor Red
                }
                Stop-Process -Id $excelPid -Force
                Write-Host "  Ended the Excel this run started (PID $excelPid). No other Excel was touched." -ForegroundColor Red
            } else {
                Write-Host "  The Excel this run started (PID $excelPid) had already gone." -ForegroundColor Red
            }
        } else {
            Write-Host '  Excel had not been started yet.' -ForegroundColor Red
        }
        try { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } catch { }
        Write-Host '  Not retrying - a second run would stop in the same place.' -ForegroundColor Red
        return
    }

    $errorFile  = Join-Path $work 'error.txt'
    $resultFile = Join-Path $work 'result.txt'
    if (Test-Path -LiteralPath $errorFile) {
        $msg = Get-Content -LiteralPath $errorFile -Raw
        if ($msg -match 'not trusted|programmatic access') {
            $msg = "Excel blocked access to the VBA project. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model'."
        }
        Write-Host "Could not run the tests: $msg" -ForegroundColor Red
        return
    }
    if (-not (Test-Path -LiteralPath $resultFile)) {
        Write-Host 'Could not run the tests: the Excel step ended without writing a result.' -ForegroundColor Red
        return
    }

    # ---- 5. Read the results ----------------------------------------------

    $verdicts = @{}
    $hookErrors = New-Object System.Collections.Generic.List[string]
    $notes = New-Object System.Collections.Generic.List[string]
    foreach ($line in [regex]::Split((Get-Content -LiteralPath $resultFile -Raw), "\r\n|\n|\r")) {
        if (-not $line) { continue }
        $f = $line.Split('|')
        switch ($f[0]) {
            'PASS'  { $verdicts[$f[1]] = [pscustomobject]@{ Passed = $true;  Reason = '' } }
            'FAIL'  { $verdicts[$f[1]] = [pscustomobject]@{ Passed = $false; Reason = ($f[2..($f.Length - 1)] -join '|') } }
            'ERROR' { $hookErrors.Add("$($f[1]): $($f[2..($f.Length - 1)] -join '|')") }
            'NOTE'  { $notes.Add("$($f[1])`n        $($f[2..($f.Length - 1)] -join '|')") }
        }
    }

    # Runner self-check: the three canaries must come back pass / fail / fail.
    $canaryWant = [ordered]@{ 'Canary_MustPass' = $true; 'Canary_MustFailOnAssert' = $false; 'Canary_MustFailOnError' = $false }
    $canaryProblems = New-Object System.Collections.Generic.List[string]
    foreach ($name in $canaryWant.Keys) {
        $key = "$canaryModule.$name"
        if (-not $verdicts.ContainsKey($key)) { $canaryProblems.Add("$name was not reported at all") }
        elseif ($verdicts[$key].Passed -ne $canaryWant[$name]) {
            $got = 'FAIL'; if ($verdicts[$key].Passed) { $got = 'PASS' }
            $canaryProblems.Add("$name came back $got")
        }
    }
    if ($canaryProblems.Count -gt 0) {
        Write-Host 'Runner self-check FAILED - results below cannot be trusted:' -ForegroundColor Red
        $canaryProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        return
    }
    Write-Host '  runner self-check ok (a passing test, a failed assert and a raised error were each reported correctly)'

    $failed  = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($mod in $moduleCounts.Keys) {
        $names = @($expected | Where-Object { $_.StartsWith("$mod.") })
        $ok = 0
        foreach ($t in $names) {
            if (-not $verdicts.ContainsKey($t)) { $missing.Add($t) }
            elseif ($verdicts[$t].Passed) { $ok++ }
            else { $failed.Add("$t`n        $($verdicts[$t].Reason)") }
        }
        $color = 'Green'; if ($ok -ne $names.Count) { $color = 'Red' }
        Write-Host ("  {0,-40} {1,2} of {2,2} passed" -f $mod, $ok, $names.Count) -ForegroundColor $color
    }

    foreach ($h in $hookErrors) { Write-Host "  ERROR in setup/teardown: $h" -ForegroundColor Red }
    foreach ($m in $missing)    { Write-Host "  NOT RUN: $m" -ForegroundColor Red }
    foreach ($x in $failed)     { Write-Host "  FAIL: $x" -ForegroundColor Red }
    foreach ($note in $notes)   { Write-Host "  note: $note" -ForegroundColor Yellow }

    Write-Host ''
    $bad = $failed.Count + $missing.Count + $hookErrors.Count
    if ($bad -gt 0) {
        Write-Host "$($failed.Count) failed, $($missing.Count) not run, $($hookErrors.Count) setup error(s), of $($expected.Count) test(s)." -ForegroundColor Red
        $exitCode = 1
    } else {
        Write-Host "All $($expected.Count) $suiteLabel test(s) passed." -ForegroundColor Green
        $exitCode = 0
    }
}
catch {
    Write-Host "Test-Unit.ps1 failed before it could report: $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 2
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    exit $exitCode
}
