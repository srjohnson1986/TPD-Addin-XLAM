<#
.SYNOPSIS
    Exports the VBA inside tools\TPD_Builder.xlsm to text files in
    tools\builder-src, so a change to the build macro shows up in a diff.

.DESCRIPTION
    TPD_Builder.xlsm is a binary file. Its BuildAddin macro decides what goes
    into every build, and until now a change to it could not be reviewed.

    The workbook stays the thing that runs; tools\builder-src is a readable
    mirror of it. After editing the macro in the VBE:

        1. save TPD_Builder.xlsm and close Excel
        2. run this script
        3. commit the workbook and tools\builder-src together

    tools\Test-All.ps1 checks the two still match (tools\Test-SourceDrift.ps1
    with -SourceRoot), so a workbook edit that was never exported fails the
    gate.

    The workbook is opened read-only with macros disabled; nothing in it runs
    and it is not saved. Needs Excel and "Trust access to the VBA project
    object model".

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Export-BuilderSource.ps1
#>

param(
    [string]$BuilderPath = (Join-Path $PSScriptRoot 'TPD_Builder.xlsm'),
    [string]$OutDir      = (Join-Path $PSScriptRoot 'builder-src'),

    # Optional: a file to write this run's Excel process id to, so a caller's
    # watchdog can end this Excel and no other if the run hangs.
    [string]$PidFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $BuilderPath)) {
    Write-Error "Builder workbook not found: $BuilderPath"
    exit 1
}
$BuilderPath = (Resolve-Path -LiteralPath $BuilderPath).Path
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

# vbext_ComponentType: 1=StandardModule, 2=ClassModule, 3=MSForm, 100=Document
$extByType = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }

$excel = $null
$wb = $null
$exitCode = 0
$written = New-Object System.Collections.Generic.List[string]

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

    Write-Host "Opening $BuilderPath (read-only, macros disabled)..."
    $wb = $excel.Workbooks.Open($BuilderPath, 0, $true)

    foreach ($comp in $wb.VBProject.VBComponents) {
        $ext = $extByType[[int]$comp.Type]
        if (-not $ext) { continue }
        $target = Join-Path $OutDir ($comp.Name + $ext)
        $comp.Export($target)
        $written.Add($comp.Name + $ext)
        Write-Host ("  {0,-28} {1,5} line(s)" -f ($comp.Name + $ext), $comp.CodeModule.CountOfLines)
    }
}
catch {
    $err = $_.Exception.Message
    if ($err -match 'not trusted|programmatic access') {
        Write-Error "Excel blocked programmatic VBA project access. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model' and re-run."
    } else {
        Write-Error "Export failed: $err"
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

if ($exitCode -eq 0) {
    # A component that no longer exists in the workbook should not linger as a file.
    $stale = @(Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' } |
                 Where-Object { $written -notcontains $_.Name -and $_.Extension -ne '.frx' })
    foreach ($s in $stale) { Write-Host "  note: $($s.Name) is in builder-src but not in the workbook - delete it if that is right." -ForegroundColor Yellow }
    Write-Host "Exported $($written.Count) component(s) to $OutDir"
}
exit $exitCode
