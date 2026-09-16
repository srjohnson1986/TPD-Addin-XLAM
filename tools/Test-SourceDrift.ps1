<#
.SYNOPSIS
    Confirms /src is exactly what's inside a built .xlam - i.e. that nothing
    was edited in the VBE and built without being exported back to /src.

.DESCRIPTION
    Opens the given .xlam via Excel COM, exports every VBA component (module/
    class/form) to a temp folder using the same VBComponent.Export the VBE
    itself uses, then text-diffs each export against the matching /src file.

    This needs Excel + "Trust access to the VBA project object model" and is
    NOT run in GitHub Actions (hosted runners have no Office install) - it's
    a local pre-release gate. Run it after Build-TPDAddin.ps1, before tagging
    a release. See CONTRIBUTING.md > "Cutting a release".

    UserForm .frx resource files are intentionally NOT compared - they're
    binary OLE streams that can differ byte-for-byte between two builds of
    identical form content (embedded GUIDs/timestamps), so a byte diff would
    be noise. Only each form's .frm text is compared.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Test-SourceDrift.ps1
    powershell -ExecutionPolicy Bypass -File tools\Test-SourceDrift.ps1 -XlamPath build\TPD_Addin.xlam
#>

param(
    [string]$XlamPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'build\TPD_Addin.xlam'),
    [string]$RepoRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $XlamPath)) {
    Write-Error "Built .xlam not found: $XlamPath (run tools\Build-TPDAddin.ps1 first)"
    exit 1
}

$srcRoot  = Join-Path $RepoRoot 'src'
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("TPD_SourceDrift_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

# vbext_ComponentType: 1=StandardModule, 2=ClassModule, 3=MSForm, 100=Document
$extByType = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }

$srcFiles = Get-ChildItem -Path $srcRoot -Recurse -Include *.bas, *.cls, *.frm

function Get-NormalizedText {
    param([string]$Path)
    $text = (Get-Content -Path $Path -Raw) -replace "`r`n", "`n"
    # VBA's Export() writes a non-deterministic number of consecutive blank
    # lines - observed both mid-file (right after the Attribute block) and
    # at end-of-file (a code-free document module exported with 2 vs 4
    # trailing blanks) - unrelated to any real code change. Collapse any
    # run of blank lines to one before comparing, and trim the file end.
    $text = $text -replace "`n{3,}", "`n`n"
    return $text.TrimEnd("`n")
}

function Compare-NormalizedText {
    param([string]$PathA, [string]$PathB)
    return (Get-NormalizedText $PathA) -eq (Get-NormalizedText $PathB)
}

$excel = $null
$wb = $null
$mismatches = New-Object System.Collections.Generic.List[string]
$builtNames = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    Write-Host "Opening $XlamPath..."
    $wb = $excel.Workbooks.Open($XlamPath, 0, $true)  # read-only

    foreach ($vbComp in $wb.VBProject.VBComponents) {
        $ext = $extByType[[int]$vbComp.Type]
        if (-not $ext) { continue }  # skip unknown/ActiveX-designer component types

        [void]$builtNames.Add($vbComp.Name)

        $exportPath = Join-Path $tempRoot ($vbComp.Name + $ext)
        $vbComp.Export($exportPath)

        $match = $srcFiles | Where-Object { $_.BaseName -ieq $vbComp.Name -and $_.Extension -ieq $ext }
        if (-not $match) {
            $mismatches.Add("In built .xlam but not found in /src: $($vbComp.Name)$ext")
            continue
        }
        if ($match.Count -gt 1) {
            $mismatches.Add("Multiple /src files named $($vbComp.Name)$ext - ambiguous: $($match.FullName -join ', ')")
            continue
        }

        if (-not (Compare-NormalizedText -PathA $exportPath -PathB $match.FullName)) {
            $relPath = $match.FullName.Substring($RepoRoot.Length).TrimStart('\', '/')
            $mismatches.Add("Source drift: $relPath does not match what's built into the .xlam (re-export from the VBE)")
        }
    }
}
catch {
    $err = $_.Exception.Message
    if ($err -match 'not trusted|programmatic access') {
        Write-Error "Excel blocked programmatic VBA project access. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model' and re-run."
    } else {
        Write-Error "Drift check failed: $err"
    }
    exit 1
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
    Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# Files that exist in /src but never showed up as a component in the built
# .xlam - usually means the base file or build step skipped something.
foreach ($f in $srcFiles) {
    if (-not $builtNames.Contains($f.BaseName)) {
        $relPath = $f.FullName.Substring($RepoRoot.Length).TrimStart('\', '/')
        $mismatches.Add("In /src but not found in built .xlam: $relPath")
    }
}

Write-Host ""
if ($mismatches.Count -gt 0) {
    Write-Host "$($mismatches.Count) drift issue(s) found:" -ForegroundColor Red
    $mismatches | ForEach-Object { Write-Host "  FAIL: $_" -ForegroundColor Red }
    exit 1
} else {
    Write-Host "No drift - /src matches the built .xlam ($($builtNames.Count) components)." -ForegroundColor Green
    exit 0
}
