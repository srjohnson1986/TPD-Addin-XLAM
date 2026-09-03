<#
.SYNOPSIS
    Drives Excel via COM to run TPD_Builder.xlsm's BuildAddin macro headlessly,
    rebuilding build/TPD_Addin.xlam from /src without opening Excel by hand.

.DESCRIPTION
    Paths default to this repo (derived from the script's own location), so a
    plain `pwsh tools\Build-TPDAddin.ps1` from anywhere works. See
    CONTRIBUTING.md > "Rebuilding a testable .xlam" for the full workflow.

    A clean run means "a build exists" - NOT "the build is good". Load the
    result as an add-in and exercise it against the fixtures in /tests before
    trusting it. This script does not touch git, run tests, or validate output.

.NOTES
    ONE-TIME MANUAL SETUP (security settings - deliberately not scripted):
      File > Options > Trust Center > Trust Center Settings > Macro Settings
      > check "Trust access to the VBA project object model."

    Requires build/_base/TPD_Addin_base.xlam (gitignored - a known-good base
    .xlam supplying the sheets / ribbon / styles / embedded logo that live
    outside /src). The BuildAddin macro copies it, imports every module from
    /src, and writes build/TPD_Addin.xlam.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1
#>

param(
    [string]$BuilderPath = (Join-Path $PSScriptRoot 'TPD_Builder.xlsm'),
    [string]$LogPath     = (Join-Path (Split-Path $PSScriptRoot -Parent) 'build\build.log')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $BuilderPath)) {
    Write-Error "Builder workbook not found: $BuilderPath"
    exit 1
}

# Truncate the log so the output below is only this run's.
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
if (Test-Path $LogPath) { Clear-Content $LogPath }

$excel = $null
$wb = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # msoAutomationSecurityLow (1): suppress the "macros disabled" security
    # prompt, which would otherwise pop a modal dialog with nobody to click it.
    # Affects only THIS Excel process, not your normal settings.
    $excel.AutomationSecurity = 1

    Write-Host 'Opening builder workbook...'
    # Open(Filename, UpdateLinks, ReadOnly)
    $wb = $excel.Workbooks.Open($BuilderPath, 0, $false)

    Write-Host 'Running BuildAddin(silent:=True)...'
    $excel.Run('BuildAddin', $true)

    Write-Host 'Done. Log output:'
    Write-Host '----------------------------------------'
    if (Test-Path $LogPath) {
        Get-Content $LogPath | ForEach-Object { Write-Host $_ }
    } else {
        Write-Warning "No log file was written - check that TPD_Builder.xlsm's log path matches -LogPath ($LogPath)."
    }
    Write-Host '----------------------------------------'

    # "CANCELLED" or a nonzero skip count = a failed build.
    if ((Test-Path $LogPath) -and (Select-String -Path $LogPath -Pattern 'CANCELLED|Skipped:' -Quiet)) {
        Write-Warning 'Build reported problems - see log above. Not treating this as a clean build.'
        $exitCode = 2
    }
}
catch {
    $err = $_.Exception.Message
    if ($err -match 'not trusted|programmatic access') {
        Write-Error "Excel blocked programmatic VBA project access. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model' and re-run."
    } else {
        Write-Error "Build failed: $err"
    }
    $exitCode = 1
}
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Remove-Variable excel, wb -ErrorAction SilentlyContinue
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
