<#
.SYNOPSIS
    The local pre-merge gate: every automated check this repo has, in order,
    stopping at the first failure, with a summary to paste into a PR.

.DESCRIPTION
    Steps, in order:

      1. Static checks self-test   tools\Test-StaticChecks.ps1 -SelfTest
      2. Static checks             tools\Test-StaticChecks.ps1
      3. Excel is free             nothing has build\TPD_Addin.xlam or the
                                   builder workbook open
      4. Builder source            the text in tools\builder-src is what is
                                   in tools\TPD_Builder.xlsm, so the build
                                   macro that runs is the one that was reviewed
      5. Build                     tools\Build-TPDAddin.ps1 (which also writes
                                   customUI\customUI14.xml into the build)
      6. Build is complete         build.log: every /src module imported,
                                   0 skipped
      7. Ribbon XML in the build   customUI14.xml inside the built .xlam is
                                   the same as customUI\customUI14.xml
      8. Source drift              tools\Test-SourceDrift.ps1
      9. Unit tests                tools\Test-Unit.ps1
     10. Flow tests                tools\Test-Unit.ps1 -Suite flows: the EQ
                                   List, Schedule and Split flows run against
                                   copies of the fixtures. These use copy and
                                   paste, so the clipboard is overwritten.

    Steps 1 and 2 need no Excel. Steps 4, 5, 8, 9 and 10 each start a hidden
    Excel.

    Each script runs in its own child process under a time limit. If a step
    that drives Excel runs out of time, this script prints the text of any
    dialog that Excel is showing, ends the Excel process that step started
    (it reports its own process id; no other Excel is touched), and stops.
    It does not retry.

    There is no separate "does a macro run" smoke check: the unit tests open
    the built add-in and call into it, which covers that and more.

    WHAT THIS DOES NOT COVER, and a person still has to do:
      - Debug > Compile VBAProject in the VBE (the only real compile check)
      - a Rubberduck Test Explorer run
      - the ribbon flows by hand against tests\fixtures (tests\README.md):
        the flow tests check what the flows produce, not the dialogs, the
        ribbon, Save Each Sheet, or how the result looks

    Exit code 0 if every step passed, 1 otherwise.

.PARAMETER StepTimeoutSeconds
    Time limit for each step that runs a script (default 240).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Test-All.ps1
#>

[CmdletBinding()]
param(
    [int]$StepTimeoutSeconds = 240,

    # Run only the steps that need no Excel (1 and 2) - for when Excel is
    # busy. Not a substitute for the full run.
    [switch]$NoExcel
)

$ErrorActionPreference = 'Stop'

$toolsDir  = $PSScriptRoot
$repoRoot  = Split-Path -Parent $toolsDir
$buildDir  = Join-Path $repoRoot 'build'
$xlamPath  = Join-Path $buildDir 'TPD_Addin.xlam'
$builderPath = Join-Path $toolsDir 'TPD_Builder.xlsm'
$builderSrc  = Join-Path $toolsDir 'builder-src'
$buildLog  = Join-Path $buildDir 'build.log'
$ribbonXml = Join-Path (Join-Path $repoRoot 'customUI') 'customUI14.xml'
$hostExe   = (Get-Process -Id $PID).Path

Add-Type -AssemblyName System.IO.Compression.FileSystem

$work = Join-Path ([System.IO.Path]::GetTempPath()) ('TPD_Gate_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null

$results = New-Object System.Collections.Generic.List[object]
$started = Get-Date

# ---------------------------------------------------------------------------
# Running one script as a step
# ---------------------------------------------------------------------------

# Runs a tools script in a child process with a time limit. Returns
# @{ Passed; Detail; Output }. $UsesExcel adds -PidFile, so a hung Excel can
# be identified and ended.
function Invoke-ScriptStep {
    param([string]$Script, [string[]]$ScriptArgs = @(), [switch]$UsesExcel, [int]$TimeoutSeconds,
          # Shown in the summary when the step passes, for a script whose last line says nothing useful.
          [string]$OkDetail)

    $tag     = [Guid]::NewGuid().ToString('N')
    $outFile = Join-Path $work "$tag.out"
    $errFile = Join-Path $work "$tag.err"
    $pidFile = Join-Path $work "$tag.pid"

    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $toolsDir $Script)`"") + $ScriptArgs
    if ($UsesExcel) { $argList += @('-PidFile', "`"$pidFile`"") }

    $child = Start-Process -FilePath $hostExe -ArgumentList $argList -PassThru -WindowStyle Hidden `
                           -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    $null = $child.Handle    # without this, ExitCode is empty once the process has gone

    $timedOut = -not $child.WaitForExit($TimeoutSeconds * 1000)
    if (-not $timedOut) { $child.WaitForExit() }    # let the redirected output finish flushing

    $output = @()
    if (Test-Path -LiteralPath $outFile) { $output += @(Get-Content -LiteralPath $outFile) }
    if (Test-Path -LiteralPath $errFile) { $output += @(Get-Content -LiteralPath $errFile | Where-Object { $_ }) }

    if ($timedOut) {
        $notes = New-Object System.Collections.Generic.List[string]
        if ($UsesExcel -and (Test-Path -LiteralPath $pidFile)) {
            $excelPid = [int](Get-Content -LiteralPath $pidFile -Raw)
            $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($proc -and $proc.ProcessName -ieq 'EXCEL') {
                $shown = @(& $hostExe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $toolsDir 'Test-Unit.ps1') -ProbeWindowsOfPid $excelPid)
                if ($shown.Count -gt 0) { $notes.Add('Excel was showing:'); $shown | ForEach-Object { $notes.Add("  $_") } }
                Stop-Process -Id $excelPid -Force
                $notes.Add("Ended the Excel this step started (PID $excelPid). No other Excel was touched.")
            }
        } elseif ($UsesExcel) {
            $notes.Add('Excel had not been started yet.')
        }
        try { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } catch { }
        return @{ Passed = $false; Detail = "timed out after $TimeoutSeconds s"; Output = @($output) + @($notes) }
    }

    $passed   = ($child.ExitCode -eq 0)
    $lastLine = @($output | Where-Object { $_ -match '[A-Za-z]' } | Select-Object -Last 1)
    $detail = ''
    if ($lastLine.Count -gt 0) { $detail = $lastLine[0].Trim() }
    if ($passed -and $OkDetail) { $detail = $OkDetail }
    if (-not $passed) { $detail = "exit code $($child.ExitCode): $detail" }
    return @{ Passed = $passed; Detail = $detail; Output = $output }
}

# ---------------------------------------------------------------------------
# Checks done here, with no Excel
# ---------------------------------------------------------------------------

function Test-ExcelIsFree {
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($lock in @((Join-Path $buildDir '~$TPD_Addin.xlam'), (Join-Path $toolsDir '~$TPD_Builder.xlsm'))) {
        if (Test-Path -LiteralPath $lock) { $problems.Add("$(Split-Path -Leaf $lock) exists - that file is open in Excel") }
    }
    if ($problems.Count -eq 0 -and (Test-Path -LiteralPath $xlamPath)) {
        try {
            $stream = [System.IO.File]::Open($xlamPath, 'Open', 'ReadWrite', 'None')
            $stream.Close()
        } catch {
            $problems.Add('build\TPD_Addin.xlam cannot be opened for writing - it is loaded as an add-in or open somewhere')
        }
    }
    if ($problems.Count -gt 0) {
        return @{ Passed = $false; Detail = 'close Excel (or unload the add-in) and run again'; Output = $problems }
    }
    return @{ Passed = $true; Detail = 'no lock on the built add-in or the builder workbook'; Output = @() }
}

function Get-NormalizedXml {
    param([string]$Text)
    return ($Text -replace "`r`n", "`n").TrimStart([char]0xFEFF).Trim()
}

# Compares customUI/customUI14.xml inside an .xlam (a zip) with the repo's copy.
function Test-RibbonXml {
    param([string]$Path, [string]$What)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{ Passed = $false; Detail = "$What not found"; Output = @("Missing: $Path") }
    }
    $inside = $null
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $zip.Entries | Where-Object { $_.FullName -ieq 'customUI/customUI14.xml' } | Select-Object -First 1
        if ($entry) {
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $inside = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
    } catch {
        return @{ Passed = $false; Detail = "could not read $What"; Output = @($_.Exception.Message) }
    } finally {
        if ($zip) { $zip.Dispose() }
    }

    if ($null -eq $inside) {
        return @{ Passed = $false; Detail = "$What has no customUI/customUI14.xml"; Output = @() }
    }
    $repoText = [System.IO.File]::ReadAllText($ribbonXml)
    if ((Get-NormalizedXml $inside) -ceq (Get-NormalizedXml $repoText)) {
        return @{ Passed = $true; Detail = "ribbon XML in $What matches customUI\customUI14.xml"; Output = @() }
    }
    return @{ Passed = $false
              Detail = "ribbon XML in $What differs from customUI\customUI14.xml"
              Output = @('The ribbon ships inside the base .xlam. After editing customUI\customUI14.xml, re-stamp the',
                         'base file (CONTRIBUTING.md > "Cutting a release", step 4) or the build carries the old ribbon.') }
}

function Test-BuildLog {
    if (-not (Test-Path -LiteralPath $buildLog)) {
        return @{ Passed = $false; Detail = 'build.log was not written'; Output = @() }
    }
    $logText  = Get-Content -LiteralPath $buildLog -Raw
    $srcCount = @(Get-ChildItem -Path (Join-Path $repoRoot 'src') -Recurse -Include *.bas, *.cls, *.frm).Count
    $done = [regex]::Match($logText, 'BUILD COMPLETE.*?(\d+) component\(s\) imported/injected, (\d+) skipped')
    if (-not $done.Success) {
        return @{ Passed = $false; Detail = 'build.log has no BUILD COMPLETE line'; Output = @($logText -split "`r?`n") }
    }
    $imported = [int]$done.Groups[1].Value
    $skipped  = [int]$done.Groups[2].Value
    if ($skipped -ne 0 -or $imported -ne $srcCount) {
        return @{ Passed = $false
                  Detail = "$imported imported, $skipped skipped, but /src has $srcCount module(s)"
                  Output = @($logText -split "`r?`n") }
    }
    return @{ Passed = $true; Detail = "$imported of $srcCount /src module(s) imported, 0 skipped"; Output = @() }
}

# ---------------------------------------------------------------------------
# The steps
# ---------------------------------------------------------------------------

$steps = @(
    @{ Name = 'Static checks self-test'; Run = { Invoke-ScriptStep -Script 'Test-StaticChecks.ps1' -ScriptArgs @('-SelfTest') -TimeoutSeconds $StepTimeoutSeconds } }
    @{ Name = 'Static checks';           Run = { Invoke-ScriptStep -Script 'Test-StaticChecks.ps1' -TimeoutSeconds $StepTimeoutSeconds } }
    @{ Name = 'Excel is free';           Excel = $true; Run = { Test-ExcelIsFree } }
    @{ Name = 'Builder source';          Excel = $true; Run = { Invoke-ScriptStep -Script 'Test-SourceDrift.ps1' -UsesExcel -TimeoutSeconds $StepTimeoutSeconds `
                                                                 -ScriptArgs @('-XlamPath', "`"$builderPath`"", '-SourceRoot', "`"$builderSrc`"", '-MacrosDisabled') } }
    @{ Name = 'Build';                   Excel = $true; Run = { Invoke-ScriptStep -Script 'Build-TPDAddin.ps1' -UsesExcel -TimeoutSeconds $StepTimeoutSeconds -OkDetail 'BuildAddin ran to the end' } }
    @{ Name = 'Build is complete';       Excel = $true; Run = { Test-BuildLog } }
    @{ Name = 'Ribbon XML in the build'; Excel = $true; Run = { Test-RibbonXml -Path $xlamPath -What 'the built add-in' } }
    @{ Name = 'Source drift';            Excel = $true; Run = { Invoke-ScriptStep -Script 'Test-SourceDrift.ps1' -UsesExcel -TimeoutSeconds $StepTimeoutSeconds } }
    # Test-Unit.ps1 has its own watchdog (120 s); the limit here only has to outlast it.
    @{ Name = 'Unit tests';              Excel = $true; Run = { Invoke-ScriptStep -Script 'Test-Unit.ps1' -TimeoutSeconds ($StepTimeoutSeconds + 60) } }
    # The flow suite has a 300 s watchdog of its own.
    @{ Name = 'Flow tests';              Excel = $true; Run = { Invoke-ScriptStep -Script 'Test-Unit.ps1' -ScriptArgs @('-Suite', 'flows') -TimeoutSeconds ($StepTimeoutSeconds + 180) } }
)

$failedStep = $null
try {
    $n = 0
    foreach ($step in $steps) {
        $n++
        if ($failedStep) {
            $results.Add([pscustomobject]@{ Name = $step.Name; State = 'skipped'; Detail = ''; Seconds = 0 })
            continue
        }
        if ($NoExcel -and $step.Excel) {
            $results.Add([pscustomobject]@{ Name = $step.Name; State = 'skipped'; Detail = 'not run (-NoExcel)'; Seconds = 0 })
            continue
        }

        Write-Host ("[{0}/{1}] {2}..." -f $n, $steps.Count, $step.Name) -ForegroundColor Cyan
        $clock = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $r = & $step.Run
        } catch {
            $r = @{ Passed = $false; Detail = "the step itself failed: $($_.Exception.Message)"; Output = @() }
        }
        $clock.Stop()

        $state = 'FAILED'
        if ($r.Passed) { $state = 'ok' }
        $results.Add([pscustomobject]@{ Name = $step.Name; State = $state; Detail = $r.Detail; Seconds = [math]::Round($clock.Elapsed.TotalSeconds) })

        if ($r.Passed) {
            Write-Host "      $($r.Detail)"
        } else {
            $failedStep = $step.Name
            Write-Host "      FAILED: $($r.Detail)" -ForegroundColor Red
            foreach ($line in @($r.Output)) { Write-Host "      | $line" -ForegroundColor DarkGray }
        }
    }
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# Summary - plain text, meant to be pasted into a PR
# ---------------------------------------------------------------------------

$branch = ''
$commit = ''
$dirty  = ''
try {
    $branch = (& git -C $repoRoot rev-parse --abbrev-ref HEAD 2>$null)
    $commit = (& git -C $repoRoot rev-parse --short HEAD 2>$null)
    if (@(& git -C $repoRoot status --porcelain 2>$null).Count -gt 0) { $dirty = ', with uncommitted changes' }
} catch { }

$total = [math]::Round(((Get-Date) - $started).TotalSeconds)
Write-Host ''
Write-Host '----- TPD add-in local gate -----'
Write-Host ("{0}  |  {1} @ {2}{3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $branch, $commit, $dirty)
foreach ($row in $results) {
    $time = ''
    if ($row.State -ne 'skipped') { $time = "$($row.Seconds)s" }
    Write-Host ("  {0,-7} {1,-25} {2,4}  {3}" -f $row.State, $row.Name, $time, $row.Detail)
}
$ranOk   = @($results | Where-Object { $_.State -eq 'ok' }).Count
$skipped = @($results | Where-Object { $_.State -eq 'skipped' }).Count
if ($failedStep) {
    Write-Host "RESULT: FAILED at '$failedStep' after $total s. Later steps were not run."
} elseif ($skipped -gt 0) {
    Write-Host "RESULT: PARTIAL - $ranOk step(s) passed in $total s, $skipped not run (-NoExcel). This is not a full pass."
} else {
    Write-Host "RESULT: all $($results.Count) steps passed in $total s."
}
Write-Host 'Not covered: Debug > Compile VBAProject, a Test Explorer run, the dialogs and ribbon, how the output looks.'
Write-Host '---------------------------------'

if ($failedStep) { exit 1 }
exit 0
