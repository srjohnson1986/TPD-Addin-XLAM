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
    .xlam supplying the sheets / styles / embedded logo that live outside
    /src; tools\New-BaseFile.ps1 makes one from a released .xlam). The
    BuildAddin macro copies it, imports every module from /src, and writes
    build/TPD_Addin.xlam.

    THE RIBBON. After the macro has finished and Excel has closed, this script
    writes customUI\customUI14.xml into the built file, replacing the copy that
    came from the base. So the repo's XML is what ships, and editing it needs
    no change to the base file. The ribbon uses only Excel's built-in icons
    (imageMso); a ribbon that needed its own image files would need more than
    this does.

    The macro is told where the repo is (this script's parent folder), so the
    build does not depend on any path saved inside the builder workbook.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Build-TPDAddin.ps1
#>

param(
    [string]$BuilderPath = (Join-Path $PSScriptRoot 'TPD_Builder.xlsm'),
    [string]$LogPath     = (Join-Path (Split-Path $PSScriptRoot -Parent) 'build\build.log'),

    # Optional: a file to write this run's Excel process id to. tools\Test-All.ps1
    # passes it so that, if the build hangs, it can end this Excel and no other.
    [string]$PidFile
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path $PSScriptRoot -Parent

if (-not (Test-Path $BuilderPath)) {
    Write-Error "Builder workbook not found: $BuilderPath"
    exit 1
}

# Replaces customUI/customUI14.xml inside the built .xlam (a zip) with the
# repo's copy. Returns a short description of what it did; throws if it could not.
function Set-RibbonXml {
    param([string]$PackagePath, [string]$XmlPath)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $entryName = 'customUI/customUI14.xml'
    $newBytes  = [System.IO.File]::ReadAllBytes($XmlPath)
    $normalize = { param($bytes) ([System.Text.Encoding]::UTF8.GetString($bytes)).TrimStart([char]0xFEFF).Replace("`r`n", "`n").Trim() }

    # Excel can keep hold of the file for a moment after it has quit.
    $zip = $null
    for ($attempt = 1; $attempt -le 20 -and -not $zip; $attempt++) {
        try { $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, 'Update') }
        catch { Start-Sleep -Milliseconds 500 }
    }
    if (-not $zip) { throw "could not open $PackagePath for writing" }

    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -ieq $entryName } | Select-Object -First 1
        if (-not $entry) {
            throw "the built file has no $entryName part. The base file carries no ribbon, and this step can only replace one, not add one."
        }

        $buffer = New-Object System.IO.MemoryStream
        $in = $entry.Open()
        try { $in.CopyTo($buffer) } finally { $in.Dispose() }

        if ((& $normalize $buffer.ToArray()) -ceq (& $normalize $newBytes)) {
            return 'already the same as customUI\customUI14.xml'
        }

        $entry.Delete()
        $fresh = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $out = $fresh.Open()
        try { $out.Write($newBytes, 0, $newBytes.Length) } finally { $out.Dispose() }
        return 'replaced with customUI\customUI14.xml'
    }
    finally {
        $zip.Dispose()
    }
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
    if ($PidFile) {
        Add-Type -Namespace TpdTools -Name Win32 -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
        $excelPid = [uint32]0
        [void][TpdTools.Win32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        [System.IO.File]::WriteAllText($PidFile, [string]$excelPid)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # msoAutomationSecurityLow (1): suppress the "macros disabled" security
    # prompt, which would otherwise pop a modal dialog with nobody to click it.
    # Affects only THIS Excel process, not your normal settings.
    $excel.AutomationSecurity = 1

    Write-Host 'Opening builder workbook...'
    # Open(Filename, UpdateLinks, ReadOnly)
    $wb = $excel.Workbooks.Open($BuilderPath, 0, $false)

    Write-Host 'Running BuildAddin(silent:=True, repoRoot)...'
    $excel.Run('BuildAddin', $true, $repoRoot)

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

# The ribbon goes in last, with Excel closed: the repo's XML replaces the base file's.
if ($exitCode -eq 0) {
    try {
        $did = Set-RibbonXml -PackagePath (Join-Path $repoRoot 'build\TPD_Addin.xlam') `
                             -XmlPath (Join-Path $repoRoot 'customUI\customUI14.xml')
        Write-Host "Ribbon XML in the build: $did"
    }
    catch {
        Write-Warning "Could not write the ribbon XML into the build: $($_.Exception.Message)"
        $exitCode = 3
    }
}

exit $exitCode
