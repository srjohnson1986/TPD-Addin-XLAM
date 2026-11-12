<#
.SYNOPSIS
    Puts the VBA in tools\builder-src into tools\TPD_Builder.xlsm, replacing
    the workbook's own copy of each module, and saves the workbook.

.DESCRIPTION
    The reverse of tools\Export-BuilderSource.ps1. Use it after editing
    tools\builder-src\*.bas as text (or after pulling someone else's change
    to it), then commit the workbook and the text together.

    Only standard modules and classes are replaced. The workbook's document
    modules (ThisWorkbook, Sheet1) are left as they are - they hold no code.

    The workbook is opened with macros disabled, so nothing in it runs. A
    module is removed, the workbook saved and reopened, and only then is the
    new text imported: importing over a module that was removed in the same
    session leaves the old one behind under a "1" suffix.

    Needs Excel and "Trust access to the VBA project object model". Close
    the builder workbook first.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Import-BuilderSource.ps1
#>

param(
    [string]$BuilderPath = (Join-Path $PSScriptRoot 'TPD_Builder.xlsm'),
    [string]$SourceDir   = (Join-Path $PSScriptRoot 'builder-src'),

    # Optional: a file to write this run's Excel process id to, so a caller's
    # watchdog can end this Excel and no other if the run hangs.
    [string]$PidFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $BuilderPath)) { Write-Error "Builder workbook not found: $BuilderPath"; exit 1 }
if (-not (Test-Path -LiteralPath $SourceDir))   { Write-Error "Source folder not found: $SourceDir"; exit 1 }
$BuilderPath = (Resolve-Path -LiteralPath $BuilderPath).Path

$lock = Join-Path (Split-Path -Parent $BuilderPath) ('~$' + (Split-Path -Leaf $BuilderPath))
if (Test-Path -LiteralPath $lock) { Write-Error 'The builder workbook is open in Excel. Close it and run again.'; exit 1 }

$documentModules = @('ThisWorkbook', 'Sheet1', 'Sheet2', 'Sheet3')
$files = @(Get-ChildItem -LiteralPath $SourceDir -File | Where-Object { $_.Extension -in '.bas', '.cls' } |
             Where-Object { $documentModules -notcontains $_.BaseName } | Sort-Object Name)
if ($files.Count -eq 0) { Write-Error "No .bas/.cls modules to import in $SourceDir"; exit 1 }

$excel = $null
$wb = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    if ($PidFile) {
        Add-Type -Namespace TpdTools -Name Win32 -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
        $excelPid = [uint32]0
        [void][TpdTools.Win32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        [System.IO.File]::WriteAllText($PidFile, [string]$excelPid)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3    # msoAutomationSecurityForceDisable: nothing in the workbook runs

    Write-Host "Opening $BuilderPath (macros disabled)..."
    $wb = $excel.Workbooks.Open($BuilderPath, 0, $false)
    if ($wb.ReadOnly) { throw 'The builder workbook opened read-only - it is open somewhere else.' }

    $removed = 0
    foreach ($f in $files) {
        $existing = $null
        try { $existing = $wb.VBProject.VBComponents.Item($f.BaseName) } catch { }
        if ($existing) { $wb.VBProject.VBComponents.Remove($existing); $removed++ }
    }
    if ($removed -gt 0) {
        $wb.Save()
        $wb.Close($false)
        $wb = $excel.Workbooks.Open($BuilderPath, 0, $false)
    }

    foreach ($f in $files) {
        [void]$wb.VBProject.VBComponents.Import($f.FullName)
        Write-Host "  imported $($f.Name)"
    }
    $wb.Save()
    Write-Host "Saved. $($files.Count) module(s) in the workbook now match $SourceDir"
}
catch {
    $err = $_.Exception.Message
    if ($err -match 'not trusted|programmatic access') {
        Write-Error "Excel blocked programmatic VBA project access. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model' and re-run."
    } else {
        Write-Error "Import failed: $err"
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
