<#
.SYNOPSIS
    Makes a base file for the build (build\_base\TPD_Addin_base.xlam) from an
    existing add-in .xlam, by removing all of its code.

.DESCRIPTION
    The build starts from a base .xlam that supplies what /src cannot: the
    worksheets, the styles, the _Resources sheet with the embedded DefaultLogo
    shape, and a ribbon part for the build to overwrite. It must hold no
    standard modules or forms - the builder cannot re-import a form that is
    already there.

    The base file is not in the repo (it carries the TPD logo, which is not
    under the MIT licence), so until now it existed only as one hand-made
    file. This script makes one from any working copy of the add-in - a
    released TPD_Addin.xlam, or build\TPD_Addin.xlam:

      1. copies the source file to a temp folder (the original is not touched)
      2. opens the copy with macros disabled
      3. removes every standard module, class and form, and empties the
         document modules (ThisWorkbook, Sheet1-3)
      4. saves it, reopens it read-only and reports what is left
      5. copies it to -OutPath

    An existing base file is never overwritten unless you pass -Force, and
    then it is first copied beside itself with a timestamp in its name.

    Refresh the base file only when something outside /src changed: a
    worksheet, a style, the logo shape. A ribbon change does not need it -
    the build writes customUI\customUI14.xml into every build.

    Needs Excel and "Trust access to the VBA project object model".

.PARAMETER From
    The add-in .xlam to strip. Required.

.PARAMETER OutPath
    Where to write the base file. Default: build\_base\TPD_Addin_base.xlam.

.PARAMETER Force
    Replace an existing file at -OutPath (after backing it up beside itself).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\New-BaseFile.ps1 -From "$env:USERPROFILE\Downloads\TPD_Addin.xlam"
#>

param(
    [Parameter(Mandatory = $true)][string]$From,
    [string]$OutPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'build\_base\TPD_Addin_base.xlam'),
    [switch]$Force,

    # Optional: a file to write this run's Excel process id to, so a caller's
    # watchdog can end this Excel and no other if the run hangs.
    [string]$PidFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $From)) { Write-Error "Source add-in not found: $From"; exit 1 }
$From = (Resolve-Path -LiteralPath $From).Path

if ((Test-Path -LiteralPath $OutPath) -and -not $Force) {
    Write-Error "$OutPath already exists. Pass -Force to replace it (it is backed up first), or give a different -OutPath."
    exit 1
}

$work = Join-Path ([System.IO.Path]::GetTempPath()) ('TPD_Base_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null
$copy = Join-Path $work 'TPD_Addin_base.xlam'
Copy-Item -LiteralPath $From -Destination $copy

$excel = $null
$wb = $null
$exitCode = 0
$report = New-Object System.Collections.Generic.List[string]

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
    $excel.AutomationSecurity = 3    # msoAutomationSecurityForceDisable: nothing in the add-in runs

    Write-Host "Stripping a copy of $From ..."
    $wb = $excel.Workbooks.Open($copy, 0, $false)

    # vbext_ComponentType: 1 standard module, 2 class, 3 form, 100 document
    $doomed = @($wb.VBProject.VBComponents | Where-Object { $_.Type -in 1, 2, 3 })
    foreach ($c in $doomed) { $wb.VBProject.VBComponents.Remove($c) }

    $emptied = 0
    foreach ($c in @($wb.VBProject.VBComponents | Where-Object { $_.Type -eq 100 })) {
        if ($c.CodeModule.CountOfLines -gt 0) { $c.CodeModule.DeleteLines(1, $c.CodeModule.CountOfLines); $emptied++ }
    }

    $wb.Save()
    $wb.Close($false)

    # Reopen and report what is actually in the saved file.
    $wb = $excel.Workbooks.Open($copy, 0, $true)
    $left      = @($wb.VBProject.VBComponents | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Type = [int]$_.Type; Lines = $_.CodeModule.CountOfLines } })
    $codeLeft  = @($left | Where-Object { $_.Type -ne 100 -or $_.Lines -gt 0 })
    $sheets    = @($wb.Worksheets | ForEach-Object { $_.Name })
    $logo      = $false
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq '_Resources') {
            foreach ($shape in $ws.Shapes) { if ($shape.Name -eq 'DefaultLogo') { $logo = $true } }
        }
    }

    $report.Add("removed $($doomed.Count) module(s)/form(s); emptied $emptied document module(s)")
    $report.Add("left in the file: $(($left | ForEach-Object { $_.Name }) -join ', ')")
    $report.Add("worksheets: $($sheets -join ', ')")
    $report.Add("DefaultLogo shape on _Resources: $(if ($logo) { 'present' } else { 'MISSING' })")

    if ($codeLeft.Count -gt 0) { throw "code is still present after stripping: $(($codeLeft | ForEach-Object { $_.Name }) -join ', ')" }
    if (-not $logo) { throw 'the stripped file has no DefaultLogo shape on a _Resources sheet - it would build an add-in with no header logo.' }
}
catch {
    $err = $_.Exception.Message
    if ($err -match 'not trusted|programmatic access') {
        Write-Error "Excel blocked programmatic VBA project access. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model' and re-run."
    } else {
        Write-Error "Could not make a base file: $err"
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

foreach ($line in $report) { Write-Host "  $line" }

if ($exitCode -eq 0) {
    # The ribbon part must be there for the build to overwrite.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null
    for ($attempt = 1; $attempt -le 20 -and -not $zip; $attempt++) {
        try { $zip = [System.IO.Compression.ZipFile]::OpenRead($copy) } catch { Start-Sleep -Milliseconds 500 }
    }
    $hasRibbon = $false
    if ($zip) {
        try { $hasRibbon = [bool]($zip.Entries | Where-Object { $_.FullName -ieq 'customUI/customUI14.xml' }) } finally { $zip.Dispose() }
    }
    Write-Host "  ribbon part (customUI/customUI14.xml): $(if ($hasRibbon) { 'present' } else { 'MISSING' })"
    if (-not $hasRibbon) {
        Write-Error 'The stripped file has no ribbon part, so the build would have no ribbon to replace. Not writing a base file.'
        $exitCode = 1
    }
}

if ($exitCode -eq 0) {
    $outDir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (Test-Path -LiteralPath $OutPath) {
        $backup = Join-Path $outDir (([System.IO.Path]::GetFileNameWithoutExtension($OutPath)) + '.backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlam')
        Copy-Item -LiteralPath $OutPath -Destination $backup
        Write-Host "  previous base file kept as $backup"
    }
    Copy-Item -LiteralPath $copy -Destination $OutPath -Force
    Write-Host "Base file written: $OutPath ($((Get-Item -LiteralPath $OutPath).Length) bytes)"
}

Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
exit $exitCode
