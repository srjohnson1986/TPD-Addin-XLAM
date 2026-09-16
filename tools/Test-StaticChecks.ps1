<#
.SYNOPSIS
    Static source checks for TPD_Addin's VBA source tree. No Excel/COM required -
    pure text/XML parsing, so this runs the same locally (pwsh, any OS) and in
    GitHub Actions on a normal hosted runner.

.DESCRIPTION
    Three checks, all against /src and customUI/customUI14.xml:

      1. Ribbon callbacks  - every onAction/onLoad named in customUI14.xml has a
                              matching Public/Private/Friend Sub or Function in /src.
      2. Option Explicit   - every module that contains code (a Sub/Function/Property)
                              declares Option Explicit. Code-free document modules
                              (e.g. an unused Sheet class) are exempt.
      3. @Folder tags      - every module has a well-formed '@Folder("TPD_Addin...")
                              annotation; every .frm is tagged exactly
                              '@Folder("TPD_Addin.Forms") per the project convention.

    Exits non-zero if any check fails, so this can be used directly as a CI gate.

.EXAMPLE
    pwsh -File tools/Test-StaticChecks.ps1
#>

[CmdletBinding()]
param(
    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrEmpty($RepoRoot)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $RepoRoot = Split-Path -Parent $scriptDir
}

$srcRoot     = Join-Path $RepoRoot 'src'
$customUIXml = Join-Path $RepoRoot 'customUI\customUI14.xml'

$failures = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    param([string]$Message)
    $script:failures.Add($Message)
    Write-Host "  FAIL: $Message" -ForegroundColor Red
}

# ---------------------------------------------------------------------------
# Gather every VBA source module once.
# ---------------------------------------------------------------------------

$moduleFiles = Get-ChildItem -Path $srcRoot -Recurse -Include *.bas, *.cls, *.frm |
    Sort-Object FullName

if ($moduleFiles.Count -eq 0) {
    throw "No .bas/.cls/.frm files found under $srcRoot - is RepoRoot correct?"
}

$subFunctionPattern = '(?im)^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)'
$anyProcPattern     = '(?im)^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+|Default\s+)*(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)'
$optionExplicitPattern = '(?im)^\s*Option\s+Explicit\s*$'
$folderTagPattern      = '(?im)^\s*''@Folder\("TPD_Addin(\.[A-Za-z0-9_]+)*"\)\s*$'

$moduleInfo = foreach ($file in $moduleFiles) {
    $text = Get-Content -Path $file.FullName -Raw
    [pscustomobject]@{
        File        = $file
        RelativePath = $file.FullName.Substring($RepoRoot.Length).TrimStart('\', '/')
        Text        = $text
        ProcNames   = [regex]::Matches($text, $subFunctionPattern) | ForEach-Object { $_.Groups[4].Value }
        HasAnyProc  = [regex]::IsMatch($text, $anyProcPattern)
    }
}

# ---------------------------------------------------------------------------
# Check 1 - every ribbon onAction/onLoad callback exists somewhere in /src.
# ---------------------------------------------------------------------------

Write-Host "`n[1/3] Ribbon callbacks (customUI/customUI14.xml)" -ForegroundColor Cyan

if (-not (Test-Path $customUIXml)) {
    Add-Failure "customUI14.xml not found at $customUIXml"
} else {
    [xml]$ribbonXml = Get-Content -Path $customUIXml -Raw

    $allProcNames = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $moduleInfo) {
        foreach ($name in $m.ProcNames) { [void]$allProcNames.Add($name) }
    }

    $callbacks = New-Object System.Collections.Generic.List[pscustomobject]

    $onLoad = $ribbonXml.customUI.onLoad
    if ($onLoad) {
        $callbacks.Add([pscustomobject]@{ Name = $onLoad; Source = 'customUI@onLoad' })
    }

    $actionNodes = $ribbonXml.SelectNodes('//*[@onAction]')
    foreach ($node in $actionNodes) {
        $id = $node.GetAttribute('id')
        $callbacks.Add([pscustomobject]@{ Name = $node.onAction; Source = "$($node.Name) id=`"$id`"" })
    }

    if ($callbacks.Count -eq 0) {
        Add-Failure "No onAction/onLoad callbacks found in customUI14.xml - XML may have failed to parse as expected."
    }

    $checked = 0
    foreach ($cb in $callbacks) {
        $checked++
        if (-not $allProcNames.Contains($cb.Name)) {
            Add-Failure "Ribbon callback '$($cb.Name)' ($($cb.Source)) has no matching Sub/Function anywhere in /src."
        }
    }

    if ($checked -gt 0) {
        Write-Host "  Checked $checked callback(s) against $($allProcNames.Count) declared Sub/Function name(s)."
    }
}

# ---------------------------------------------------------------------------
# Check 2 - Option Explicit in every module that has code.
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Option Explicit" -ForegroundColor Cyan

$checkedCount = 0
foreach ($m in $moduleInfo) {
    if (-not $m.HasAnyProc) { continue }  # code-free document module - exempt
    $checkedCount++
    if (-not [regex]::IsMatch($m.Text, $optionExplicitPattern)) {
        Add-Failure "Missing 'Option Explicit': $($m.RelativePath)"
    }
}
Write-Host "  Checked $checkedCount module(s) with code (of $($moduleInfo.Count) total; code-free document modules are exempt)."

# ---------------------------------------------------------------------------
# Check 3 - @Folder annotations well-formed, and every .frm tagged Forms.
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] @Folder annotations" -ForegroundColor Cyan

foreach ($m in $moduleInfo) {
    $folderMatch = [regex]::Match($m.Text, $folderTagPattern)
    if (-not $folderMatch.Success) {
        Add-Failure "Missing or malformed '@Folder(""TPD_Addin...."")' annotation: $($m.RelativePath)"
        continue
    }

    if ($m.File.Extension -ieq '.frm' -and $folderMatch.Value.Trim() -ne '''@Folder("TPD_Addin.Forms")') {
        Add-Failure "UserForm not tagged '@Folder(""TPD_Addin.Forms"")' (every .frm must use this exact tag): $($m.RelativePath)"
    }
}
Write-Host "  Checked $($moduleInfo.Count) module(s)."

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
if ($failures.Count -gt 0) {
    Write-Host "$($failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
} else {
    Write-Host "All static checks passed ($($moduleInfo.Count) modules)." -ForegroundColor Green
    exit 0
}
