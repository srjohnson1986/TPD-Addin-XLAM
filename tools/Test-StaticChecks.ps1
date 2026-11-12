<#
.SYNOPSIS
    Static source checks for TPD_Addin's VBA source tree. No Excel/COM required -
    pure text/XML parsing, so this runs the same locally (Windows PowerShell 5.1
    or pwsh, any OS) and in GitHub Actions on a normal hosted runner.

.DESCRIPTION
    Eleven checks against /src, /tests/unit, /tools, customUI/customUI14.xml, README.md and
    CHANGELOG.md. Each failure carries a short id (in brackets below), which is
    what -SelfTest asserts on.

       1. Ribbon callbacks   [ribbon-callback]  every onAction/onLoad named in
                             customUI14.xml has a matching Sub/Function in /src.
       2. Option Explicit    [option-explicit]  every module that contains code
                             declares it. Code-free document modules are exempt.
       3. @Folder tags       [folder-tag]       every module has a well-formed
                             '@Folder("TPD_Addin...") tag; every .frm is tagged
                             exactly '@Folder("TPD_Addin.Forms").
       4. File shape         [file-shape] [line-endings] [duplicate-name]
                             CRLF only, no BOM, Attribute VB_Name matches the
                             file name, .frm header names its own form and .frx,
                             the .frx exists, no two modules share a name. A
                             hand-written .frm that gets this wrong is imported
                             by the builder as a plain module.
       5. Declarations first [decl-after-proc]  no module-level code after the
                             first Sub/Function/Property. A Const or Dim down
                             there does not resolve, and nothing says so until
                             a flow reaches it at run time (2026-09-08).
       6. Parenthesised call [paren-call]       no statement of the form
                             "MySub (arg)" for a Sub/Function declared in this
                             repo. The parens force by-value evaluation, so an
                             object argument raises run-time 438 (#29).
       7. String-named procs [string-proc-ref]  every string literal handed to
                             Application.OnTime / Application.Run names a
                             Sub/Function in a standard module in /src.
       8. WithPerformance    [perf-parity]      every PERF_* constant has a Case
                             that calls the Sub its value names, that Sub
                             exists, and some caller uses the constant.
       9. Constants declared [undeclared-const] every UPPER_SNAKE identifier
                             used in code is declared in that module or Public
                             in a standard module. Catches a mistyped PREF_* /
                             PERF_* that Option Explicit would only report at
                             Debug > Compile.
      10. Version sync       [version-sync]     ADDIN_VERSION, the README's
                             pinned download link and the newest CHANGELOG
                             heading all name the same version.
      11. Scripts parse      [ps-syntax]        every tools/*.ps1 parses. A script
                             with a syntax error fails only when it is
                             run, which for the Excel tools is rarely.

    These are text heuristics, not a compiler. They cover the mistakes this
    project has actually shipped; they do not replace Debug > Compile VBAProject.

    Check 4's line-ending test can only fail on a local working tree: a git
    checkout always writes CRLF for these files (.gitattributes), so in CI it
    passes by construction. It is there for the hand-written file you have not
    committed yet.

    Exits non-zero if any check fails, so this can be used directly as a CI gate.

.PARAMETER SelfTest
    Instead of checking the repo, build a small synthetic source tree in the
    temp folder, confirm it passes clean, then break it one way at a time and
    confirm each break is reported by exactly the expected check. Run this
    after changing any check: a check that never fires proves nothing.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Test-StaticChecks.ps1
    powershell -ExecutionPolicy Bypass -File tools\Test-StaticChecks.ps1 -SelfTest
#>

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'

$script:Quiet    = $false
$script:Failures = [System.Collections.Generic.List[object]]::new()
$script:Latin1   = [System.Text.Encoding]::GetEncoding(28591)

$rxOpts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
# Everything up to an apostrophe that is not inside a string literal.
$script:RxComment   = New-Object System.Text.RegularExpressions.Regex('^((?:[^"'']|"[^"]*")*)''.*$')
$script:RxString    = New-Object System.Text.RegularExpressions.Regex('"[^"]*"')
$script:RxProcStart = New-Object System.Text.RegularExpressions.Regex('^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)', $rxOpts)
$script:RxProcEnd   = New-Object System.Text.RegularExpressions.Regex('^\s*End\s+(Sub|Function|Property)\b', $rxOpts)
# Case-sensitive on purpose: the VBE re-cases every identifier to match its
# declaration, so an all-caps name with an underscore is a constant by convention.
$script:RxUpper     = New-Object System.Text.RegularExpressions.Regex('(?<![\w.])[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+(?!\w)')

function Write-Line {
    param([string]$Text, [string]$Color)
    if ($script:Quiet) { return }
    if ($Color) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
}

function Add-Failure {
    param([string]$Id, [string]$Message)
    $script:Failures.Add([pscustomobject]@{ Id = $Id; Message = $Message })
    Write-Line "  FAIL [$Id]: $Message" 'Red'
}

function Add-UpperNames {
    param($Set, [string]$Code)
    foreach ($hit in $script:RxUpper.Matches($Code)) { [void]$Set.Add($hit.Value) }
}

# ---------------------------------------------------------------------------
# VBA text -> logical lines. Comments are dropped, " _" continuations are
# joined, and each line comes back twice: Raw (string literals intact) and
# Code (every literal blanked to ""), so a check can pattern-match Code without
# tripping over text that only looks like code.
# ---------------------------------------------------------------------------
function Get-VbaLines {
    param([string]$Text)

    $out      = [System.Collections.Generic.List[object]]::new()
    $physical = [regex]::Split($Text, "\r\n|\n|\r")
    $accCode  = $null
    $accRaw   = $null
    $accNo    = 0

    for ($i = 0; $i -lt $physical.Length; $i++) {
        $raw = $script:RxComment.Replace($physical[$i], '$1')
        if ($raw -match '^\s*Rem(\s|$)') { $raw = '' }
        $code = $script:RxString.Replace($raw, '""')

        $continues = $code -match '(^|\s)_\s*$'
        if ($continues) {
            $code = $code -replace '_\s*$', ''
            $raw  = $raw  -replace '_\s*$', ''
        }

        if ($null -eq $accCode) {
            $accCode = $code
            $accRaw  = $raw
            $accNo   = $i + 1
        } else {
            $accCode = $accCode.TrimEnd() + ' ' + $code.TrimStart()
            $accRaw  = $accRaw.TrimEnd()  + ' ' + $raw.TrimStart()
        }

        if (-not $continues) {
            $out.Add([pscustomobject]@{ N = $accNo; Code = $accCode; Raw = $accRaw })
            $accCode = $null
        }
    }
    if ($null -ne $accCode) {
        $out.Add([pscustomobject]@{ N = $accNo; Code = $accCode; Raw = $accRaw })
    }
    return ,$out
}

# Splits an argument list on its top-level commas (not inside parens/strings).
function Split-TopLevelArgs {
    param([string]$Text)

    $parts = [System.Collections.Generic.List[string]]::new()
    $sb    = New-Object System.Text.StringBuilder
    $depth = 0
    $inStr = $false
    foreach ($ch in $Text.ToCharArray()) {
        if ($inStr) {
            [void]$sb.Append($ch)
            if ($ch -eq [char]'"') { $inStr = $false }
            continue
        }
        if ($ch -eq [char]'"') { $inStr = $true; [void]$sb.Append($ch); continue }
        if ($ch -eq [char]'(') { $depth++ }
        if ($ch -eq [char]')') { $depth-- }
        if ($ch -eq [char]',' -and $depth -le 0) {
            $parts.Add($sb.ToString().Trim())
            [void]$sb.Clear()
            continue
        }
        [void]$sb.Append($ch)
    }
    $parts.Add($sb.ToString().Trim())
    return ,$parts
}

# ---------------------------------------------------------------------------
# One pass over a module: its procedures, where the first one starts, any
# module-level code after it, and the UPPER_SNAKE names it declares.
# ---------------------------------------------------------------------------
function Get-ModuleAnalysis {
    param($File, [string]$Root, [string]$Area)

    $bytes = [System.IO.File]::ReadAllBytes($File.FullName)
    $text  = $script:Latin1.GetString($bytes)
    $lines = Get-VbaLines $text
    $ext   = $File.Extension.ToLowerInvariant()
    $isStandard = ($ext -eq '.bas')

    # Skip the designer / class header: real code starts after Attribute VB_Name.
    $start = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Code -match '^Attribute\s+VB_Name\b') { $start = $i + 1; break }
    }

    $procs    = [System.Collections.Generic.List[object]]::new()
    $late     = [System.Collections.Generic.List[object]]::new()
    $body     = [System.Collections.Generic.List[object]]::new()
    $declared = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $public   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $inProc      = $false
    $seenProc    = $false
    $block       = $null    # 'Enum' / 'Type' while inside one
    $blockPublic = $false

    for ($i = $start; $i -lt $lines.Count; $i++) {
        $ln   = $lines[$i]
        $code = $ln.Code.Trim()
        if ($code.Length -eq 0) { continue }
        if ($code -match '^Attribute\s') { continue }

        if ($code.StartsWith('#')) {
            if ($code -match '^#Const\s+([A-Za-z_]\w*)') { [void]$declared.Add($Matches[1]) }
            continue
        }

        if ($inProc) {
            $body.Add([pscustomobject]@{ N = $ln.N; Code = $ln.Code; Raw = $ln.Raw; InProc = $true })
            if ($script:RxProcEnd.IsMatch($code)) { $inProc = $false; continue }
            if ($code -match '^(Dim|Static|Const)\b') { Add-UpperNames $declared $code }
            if ($code -match '^([A-Za-z_]\w*):\s*$')   { [void]$declared.Add($Matches[1]) }
            continue
        }

        $body.Add([pscustomobject]@{ N = $ln.N; Code = $ln.Code; Raw = $ln.Raw; InProc = $false })

        $ps = $script:RxProcStart.Match($code)
        if ($ps.Success) {
            $procs.Add([pscustomobject]@{
                Name  = $ps.Groups[3].Value
                Kind  = ($ps.Groups[2].Value -replace '\s+.*$', '')
                Scope = $ps.Groups[1].Value
                N     = $ln.N
            })
            Add-UpperNames $declared $code    # parameter names
            $inProc   = $true
            $seenProc = $true
            continue
        }

        # Module-level code. After the first procedure, VBA accepts nothing here.
        if ($seenProc) { $late.Add([pscustomobject]@{ N = $ln.N; Text = $code }) }

        if ($block) {
            if ($code -match '^End\s+(Enum|Type)\b') { $block = $null; continue }
            if ($block -eq 'Enum' -and $code -match '^([A-Za-z_]\w*)') {
                [void]$declared.Add($Matches[1])
                if ($blockPublic) { [void]$public.Add($Matches[1]) }
            }
            continue
        }
        if ($code -match '^(?:(Public|Private)\s+)?(Enum|Type)\s+[A-Za-z_]\w*') {
            $block       = $Matches[2]
            $blockPublic = ($isStandard -and $Matches[1] -ne 'Private')
            continue
        }

        Add-UpperNames $declared $code
        if ($isStandard -and $code -match '^(Public|Global)\s') { Add-UpperNames $public $code }
    }

    [pscustomobject]@{
        File       = $File
        Rel        = $File.FullName.Substring($Root.Length).TrimStart('\', '/')
        Area       = $Area
        Ext        = $ext
        BaseName   = $File.BaseName
        Bytes      = $bytes
        Text       = $text
        IsStandard = $isStandard
        Procs      = $procs
        Late       = $late
        Body       = $body
        Declared   = $declared
        Public     = $public
    }
}

# ---------------------------------------------------------------------------
# The checks. Returns nothing; failures accumulate in $script:Failures.
# ---------------------------------------------------------------------------
function Invoke-StaticChecks {
    param([string]$Root)

    $Root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\', '/')
    $script:Failures = [System.Collections.Generic.List[object]]::new()

    $srcRoot     = Join-Path $Root 'src'
    $testRoot    = Join-Path (Join-Path $Root 'tests') 'unit'
    $customUIXml = Join-Path (Join-Path $Root 'customUI') 'customUI14.xml'
    $readmePath  = Join-Path $Root 'README.md'
    $changelog   = Join-Path $Root 'CHANGELOG.md'
    $total       = 11

    if (-not (Test-Path -LiteralPath $srcRoot)) {
        throw "No src folder under $Root - is RepoRoot correct?"
    }

    $srcFiles = @(Get-ChildItem -Path $srcRoot -Recurse -Include *.bas, *.cls, *.frm | Sort-Object FullName)
    if ($srcFiles.Count -eq 0) {
        throw "No .bas/.cls/.frm files found under $srcRoot - is RepoRoot correct?"
    }
    $testFiles = @()
    if (Test-Path -LiteralPath $testRoot) {
        $testFiles = @(Get-ChildItem -Path $testRoot -Recurse -Include *.bas, *.cls, *.frm | Sort-Object FullName)
    }

    $srcModules  = @(foreach ($f in $srcFiles)  { Get-ModuleAnalysis -File $f -Root $Root -Area 'src' })
    $testModules = @(foreach ($f in $testFiles) { Get-ModuleAnalysis -File $f -Root $Root -Area 'tests' })
    $modules     = @($srcModules) + @($testModules)

    # The leading comma stops PowerShell unrolling the (empty) set into $null.
    $newSet = { ,([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)) }

    $srcCallables = & $newSet    # Sub/Function names in /src (any module kind)
    $stdCallables = & $newSet    # ...in standard modules only
    $allCallables = & $newSet    # ...in /src and /tests/unit
    $moduleNames  = & $newSet
    $publicUpper  = & $newSet
    foreach ($m in $modules) {
        [void]$moduleNames.Add($m.BaseName)
        foreach ($p in $m.Procs) {
            if ($p.Kind -ieq 'Property') { continue }
            [void]$allCallables.Add($p.Name)
            if ($m.Area -eq 'src') {
                [void]$srcCallables.Add($p.Name)
                if ($m.IsStandard) { [void]$stdCallables.Add($p.Name) }
            }
        }
        if ($m.Area -eq 'src') { foreach ($n in $m.Public) { [void]$publicUpper.Add($n) } }
    }

    # -- 1 ------------------------------------------------------------------
    Write-Line "`n[1/$total] Ribbon callbacks (customUI/customUI14.xml)" 'Cyan'

    if (-not (Test-Path -LiteralPath $customUIXml)) {
        Add-Failure 'ribbon-callback' "customUI14.xml not found at $customUIXml"
    } else {
        [xml]$ribbonXml = Get-Content -LiteralPath $customUIXml -Raw
        $callbacks = [System.Collections.Generic.List[object]]::new()

        $onLoad = $ribbonXml.customUI.onLoad
        if ($onLoad) { $callbacks.Add([pscustomobject]@{ Name = $onLoad; Source = 'customUI@onLoad' }) }
        foreach ($node in $ribbonXml.SelectNodes('//*[@onAction]')) {
            $id = $node.GetAttribute('id')
            $callbacks.Add([pscustomobject]@{ Name = $node.onAction; Source = "$($node.Name) id=`"$id`"" })
        }

        if ($callbacks.Count -eq 0) {
            Add-Failure 'ribbon-callback' 'No onAction/onLoad callbacks found in customUI14.xml - XML may have failed to parse as expected.'
        }
        foreach ($cb in $callbacks) {
            if (-not $srcCallables.Contains($cb.Name)) {
                Add-Failure 'ribbon-callback' "Ribbon callback '$($cb.Name)' ($($cb.Source)) has no matching Sub/Function anywhere in /src."
            }
        }
        Write-Line "  Checked $($callbacks.Count) callback(s) against $($srcCallables.Count) declared Sub/Function name(s)."
    }

    # -- 2 ------------------------------------------------------------------
    Write-Line "`n[2/$total] Option Explicit" 'Cyan'

    $checked = 0
    foreach ($m in $modules) {
        if ($m.Procs.Count -eq 0) { continue }    # code-free document module - exempt
        $checked++
        if (-not [regex]::IsMatch($m.Text, '(?im)^\s*Option\s+Explicit\s*$')) {
            Add-Failure 'option-explicit' "Missing 'Option Explicit': $($m.Rel)"
        }
    }
    Write-Line "  Checked $checked module(s) with code (of $($modules.Count); code-free document modules are exempt)."

    # -- 3 ------------------------------------------------------------------
    Write-Line "`n[3/$total] @Folder annotations" 'Cyan'

    foreach ($m in $modules) {
        $tag = [regex]::Match($m.Text, '(?im)^\s*''@Folder\("TPD_Addin(\.[A-Za-z0-9_]+)*"\)\s*$')
        if (-not $tag.Success) {
            Add-Failure 'folder-tag' "Missing or malformed '@Folder(""TPD_Addin...."")' annotation: $($m.Rel)"
            continue
        }
        if ($m.Ext -eq '.frm' -and $tag.Value.Trim() -ne '''@Folder("TPD_Addin.Forms")') {
            Add-Failure 'folder-tag' "UserForm not tagged '@Folder(""TPD_Addin.Forms"")' (every .frm must use this exact tag): $($m.Rel)"
        }
    }
    Write-Line "  Checked $($modules.Count) module(s)."

    # -- 4 ------------------------------------------------------------------
    Write-Line "`n[4/$total] File shape and line endings" 'Cyan'

    foreach ($m in $modules) {
        $base = $m.BaseName

        if ($m.Bytes.Length -ge 3 -and $m.Bytes[0] -eq 0xEF -and $m.Bytes[1] -eq 0xBB -and $m.Bytes[2] -eq 0xBF) {
            Add-Failure 'file-shape' "Starts with a UTF-8 BOM (the VBE writes none, and the importer will not see line 1): $($m.Rel)"
            continue
        }

        $bareLf = [regex]::Matches($m.Text, '(?<!\r)\n').Count
        $bareCr = [regex]::Matches($m.Text, '\r(?!\n)').Count
        if (($bareLf + $bareCr) -gt 0) {
            Add-Failure 'line-endings' "$($bareLf + $bareCr) line ending(s) are not CRLF (the builder imports a non-CRLF .frm as a plain module): $($m.Rel)"
        }

        $firstLines = [regex]::Split($m.Text, "\r\n|\n|\r")
        $vbName = [regex]::Match($m.Text, '(?m)^Attribute VB_Name = "([^"]*)"')
        if (-not $vbName.Success) {
            Add-Failure 'file-shape' "No 'Attribute VB_Name' line: $($m.Rel)"
        } elseif ($vbName.Groups[1].Value -ine $base) {
            Add-Failure 'file-shape' "Attribute VB_Name is ""$($vbName.Groups[1].Value)"" but the file is named $base$($m.Ext): $($m.Rel)"
        }

        switch ($m.Ext) {
            '.bas' {
                if ($firstLines[0] -notmatch '^Attribute VB_Name = ') {
                    Add-Failure 'file-shape' "Line 1 of a .bas must be its 'Attribute VB_Name' line: $($m.Rel)"
                }
            }
            '.cls' {
                if ($firstLines[0].TrimEnd() -ne 'VERSION 1.0 CLASS') {
                    Add-Failure 'file-shape' "Line 1 of a .cls must be 'VERSION 1.0 CLASS': $($m.Rel)"
                }
            }
            '.frm' {
                if ($firstLines[0].TrimEnd() -ne 'VERSION 5.00') {
                    Add-Failure 'file-shape' "Line 1 of a .frm must be 'VERSION 5.00': $($m.Rel)"
                }
                $beginRx = '^Begin \{[0-9A-Fa-f-]{36}\} ' + [regex]::Escape($base) + '\s*$'
                if ($firstLines.Length -lt 2 -or $firstLines[1] -notmatch $beginRx) {
                    Add-Failure 'file-shape' "Line 2 of a .frm must be 'Begin {GUID} $base': $($m.Rel)"
                }
                $blobRx = '(?m)^\s*OleObjectBlob\s*=\s*"' + [regex]::Escape($base) + '\.frx":'
                if (-not [regex]::IsMatch($m.Text, $blobRx)) {
                    Add-Failure 'file-shape' "No 'OleObjectBlob = ""$base.frx"":...' line: $($m.Rel)"
                }
                $frx = Join-Path $m.File.DirectoryName ($base + '.frx')
                if (-not (Test-Path -LiteralPath $frx)) {
                    Add-Failure 'file-shape' "No $base.frx beside the form: $($m.Rel)"
                }
            }
        }
    }

    foreach ($frx in @(Get-ChildItem -Path $srcRoot -Recurse -Filter *.frx)) {
        $frm = Join-Path $frx.DirectoryName ($frx.BaseName + '.frm')
        if (-not (Test-Path -LiteralPath $frm)) {
            Add-Failure 'file-shape' "Orphan .frx with no .frm beside it: $($frx.FullName.Substring($Root.Length).TrimStart('\', '/'))"
        }
    }

    foreach ($g in @($modules | Group-Object { $_.BaseName.ToLowerInvariant() } | Where-Object { $_.Count -gt 1 })) {
        Add-Failure 'duplicate-name' "More than one module named '$($g.Group[0].BaseName)': $(($g.Group | ForEach-Object { $_.Rel }) -join ', ')"
    }
    Write-Line "  Checked $($modules.Count) module(s)."

    # -- 5 ------------------------------------------------------------------
    Write-Line "`n[5/$total] Declarations above the first procedure" 'Cyan'

    foreach ($m in $modules) {
        foreach ($l in $m.Late) {
            $shown = $l.Text
            if ($shown.Length -gt 70) { $shown = $shown.Substring(0, 70) + '...' }
            Add-Failure 'decl-after-proc' "Module-level code after the first procedure (move it above): $($m.Rel):$($l.N)  $shown"
        }
    }
    Write-Line "  Checked $($modules.Count) module(s)."

    # -- 6 ------------------------------------------------------------------
    Write-Line "`n[6/$total] Parenthesised Sub calls" 'Cyan'

    $rxParenStart = New-Object System.Text.RegularExpressions.Regex('(?:^\s*|\bThen\s+|\bElse\s+|:\s+)((?:[A-Za-z_]\w*\.)?[A-Za-z_]\w*) \(', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $statements = 0
    foreach ($m in $modules) {
        foreach ($ln in $m.Body) {
            if (-not $ln.InProc) { continue }
            $statements++
            foreach ($hit in $rxParenStart.Matches($ln.Code)) {
                $parts = $hit.Groups[1].Value.Split('.')
                $name  = $parts[$parts.Length - 1]
                if ($parts.Length -eq 2 -and -not $moduleNames.Contains($parts[0])) { continue }
                if (-not $allCallables.Contains($name)) { continue }

                # Find the paren that closes the one just matched.
                $open  = $hit.Index + $hit.Length - 1
                $depth = 0
                $close = -1
                for ($k = $open; $k -lt $ln.Code.Length; $k++) {
                    $c = $ln.Code[$k]
                    if ($c -eq [char]'(') { $depth++ }
                    elseif ($c -eq [char]')') { $depth--; if ($depth -eq 0) { $close = $k; break } }
                }
                if ($close -lt 0) { continue }
                $rest = $ln.Code.Substring($close + 1).TrimStart()
                if ($rest.StartsWith('=') -or $rest.StartsWith('.')) { continue }

                Add-Failure 'paren-call' "'$name (...)' forces its argument by value - write '$name arg' or 'Call $name(arg)' (#29): $($m.Rel):$($ln.N)"
            }
        }
    }
    Write-Line "  Checked $statements statement(s) against $($allCallables.Count) Sub/Function name(s)."

    # -- 7 ------------------------------------------------------------------
    Write-Line "`n[7/$total] String-named procedures (Application.OnTime / Run)" 'Cyan'

    $rxStringCall = New-Object System.Text.RegularExpressions.Regex('\bApplication\.(OnTime|Run)\b\s*\(?(.*)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $literalRefs = 0
    $dynamicRefs = 0
    foreach ($m in $srcModules) {
        foreach ($ln in $m.Body) {
            $call = $rxStringCall.Match($ln.Raw)
            if (-not $call.Success) { continue }

            $method = $call.Groups[1].Value
            $callArgs = Split-TopLevelArgs $call.Groups[2].Value
            $target = $null
            foreach ($a in $callArgs) {
                if ($a -match '^(Procedure|Macro)\s*:=\s*(.+)$') { $target = $Matches[2].Trim() }
            }
            if ($null -eq $target) {
                $index = 0
                if ($method -ieq 'OnTime') { $index = 1 }
                if ($callArgs.Count -gt $index) { $target = $callArgs[$index] }
            }
            if ([string]::IsNullOrEmpty($target)) { continue }

            $lit = [regex]::Match($target, '^"((?:[^"]|"")*)"\)?$')
            if (-not $lit.Success) { $dynamicRefs++; continue }
            $literalRefs++

            $named = $lit.Groups[1].Value
            $bang  = $named.LastIndexOf('!')
            if ($bang -ge 0) { $named = $named.Substring($bang + 1) }
            $ident = [regex]::Match($named.Trim("'", ' '), '^[A-Za-z_][\w.]*')
            $segs  = $ident.Value.Split('.')
            $proc  = $segs[$segs.Length - 1]

            $found = $false
            if ($segs.Length -ge 2) {
                $owner = $srcModules | Where-Object { $_.IsStandard -and $_.BaseName -ieq $segs[$segs.Length - 2] } | Select-Object -First 1
                if ($owner) { $found = [bool]($owner.Procs | Where-Object { $_.Name -ieq $proc -and $_.Kind -ine 'Property' }) }
            } else {
                $found = $stdCallables.Contains($proc)
            }
            if (-not $found) {
                Add-Failure 'string-proc-ref' "Application.$method names ""$($lit.Groups[1].Value)"", which is not a Sub/Function in a standard module in /src: $($m.Rel):$($ln.N)"
            }
        }
    }
    Write-Line "  Checked $literalRefs string-named call(s)."
    if ($dynamicRefs -gt 0) {
        Write-Line "  Note: $dynamicRefs call(s) pass a name that is not a string literal - those cannot be checked."
    }

    # -- 8 ------------------------------------------------------------------
    Write-Line "`n[8/$total] WithPerformance wiring (PERF_* constants)" 'Cyan'

    $perfModule = $srcModules | Where-Object { $_.Procs | Where-Object { $_.Name -ieq 'WithPerformance' } } | Select-Object -First 1
    if (-not $perfModule) {
        Add-Failure 'perf-parity' "No 'Sub WithPerformance' found in /src - this check needs updating if it was renamed."
    } else {
        $perfConsts = [ordered]@{}
        $perfCases  = [ordered]@{}
        $bodyLines  = $perfModule.Body
        for ($i = 0; $i -lt $bodyLines.Count; $i++) {
            $ln = $bodyLines[$i]
            if (-not $ln.InProc) {
                $c = [regex]::Match($ln.Raw, '^\s*Public\s+Const\s+(PERF_\w+)\s+As\s+String\s*=\s*"(\w+)"')
                if ($c.Success) { $perfConsts[$c.Groups[1].Value] = $c.Groups[2].Value }
                continue
            }
            $k = [regex]::Match($ln.Code, '^\s*Case\s+(PERF_\w+)\s*$')
            if ($k.Success) {
                $calls = ''
                if ($i + 1 -lt $bodyLines.Count) {
                    $next = [regex]::Match($bodyLines[$i + 1].Code, '^\s*([A-Za-z_]\w*)')
                    if ($next.Success) { $calls = $next.Groups[1].Value }
                }
                $perfCases[$k.Groups[1].Value] = $calls
            }
        }

        if ($perfConsts.Count -eq 0) {
            Add-Failure 'perf-parity' "No 'Public Const PERF_* As String = ""...""' found in $($perfModule.Rel)."
        }
        foreach ($name in $perfConsts.Keys) {
            $value = $perfConsts[$name]
            if (-not $perfCases.Contains($name)) {
                Add-Failure 'perf-parity' "$name has no 'Case $name' in WithPerformance ($($perfModule.Rel))."
            } elseif ($perfCases[$name] -ine $value) {
                Add-Failure 'perf-parity' "'Case $name' calls '$($perfCases[$name])' but the constant's value names '$value' ($($perfModule.Rel))."
            }
            if (-not $srcCallables.Contains($value)) {
                Add-Failure 'perf-parity' "$name = ""$value"", but there is no Sub/Function '$value' in /src."
            }
            $nameRx = '(?<![\w.])' + [regex]::Escape($name) + '(?!\w)'
            $used = $false
            foreach ($m in $srcModules) {
                if ($m.Rel -eq $perfModule.Rel) { continue }
                foreach ($ln in $m.Body) { if ($ln.Code -cmatch $nameRx) { $used = $true; break } }
                if ($used) { break }
            }
            if (-not $used) {
                Add-Failure 'perf-parity' "$name is declared but nothing outside $($perfModule.BaseName) passes it to WithPerformance."
            }
        }
        foreach ($name in $perfCases.Keys) {
            if (-not $perfConsts.Contains($name)) {
                Add-Failure 'perf-parity' "'Case $name' in WithPerformance has no matching 'Public Const $name' ($($perfModule.Rel))."
            }
        }
        Write-Line "  Checked $($perfConsts.Count) constant(s) against $($perfCases.Count) Case(s)."
    }

    # -- 9 ------------------------------------------------------------------
    Write-Line "`n[9/$total] Constants declared (UPPER_SNAKE identifiers)" 'Cyan'

    $usesChecked = 0
    foreach ($m in $modules) {
        $reported = & $newSet
        foreach ($ln in $m.Body) {
            foreach ($hit in $script:RxUpper.Matches($ln.Code)) {
                $usesChecked++
                $n = $hit.Value
                if ($m.Declared.Contains($n) -or $publicUpper.Contains($n)) { continue }
                if (-not $reported.Add($n)) { continue }
                Add-Failure 'undeclared-const' "'$n' is not declared in this module and is not Public in any standard module: $($m.Rel):$($ln.N)"
            }
        }
    }
    Write-Line "  Checked $usesChecked use(s) against $($publicUpper.Count) Public constant(s) plus each module's own."

    # -- 10 -----------------------------------------------------------------
    Write-Line "`n[10/$total] Version sync (ADDIN_VERSION / README / CHANGELOG)" 'Cyan'

    $version = $null
    foreach ($m in $srcModules) {
        $v = [regex]::Match($m.Text, '(?m)^\s*Public\s+Const\s+ADDIN_VERSION\s+As\s+String\s*=\s*"(\d+\.\d+\.\d+)"')
        if ($v.Success) { $version = $v.Groups[1].Value; break }
    }
    if (-not $version) {
        Add-Failure 'version-sync' "No 'Public Const ADDIN_VERSION As String = ""X.Y.Z""' found in /src."
    } else {
        if (-not (Test-Path -LiteralPath $readmePath)) {
            Add-Failure 'version-sync' 'README.md not found.'
        } else {
            $readmeLines = [regex]::Split((Get-Content -LiteralPath $readmePath -Raw), "\r\n|\n|\r")
            $markerAt = -1
            for ($i = 0; $i -lt $readmeLines.Length; $i++) {
                if ($readmeLines[$i] -match '<!--\s*download-link:') { $markerAt = $i; break }
            }
            if ($markerAt -lt 0 -or $markerAt + 1 -ge $readmeLines.Length) {
                Add-Failure 'version-sync' "README.md has no '<!-- download-link: ... -->' marker line followed by the link (tools/sync-readme-version.sh needs it too)."
            } else {
                $linkVersions = @([regex]::Matches($readmeLines[$markerAt + 1], 'v(\d+\.\d+\.\d+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
                if ($linkVersions.Count -eq 0) {
                    Add-Failure 'version-sync' 'The README line after the download-link marker carries no vX.Y.Z version.'
                } elseif ($linkVersions.Count -gt 1 -or $linkVersions[0] -ne $version) {
                    Add-Failure 'version-sync' "README download link says v$($linkVersions -join ' / v') but ADDIN_VERSION is $version (run tools/sync-readme-version.sh)."
                }
            }
        }

        if (-not (Test-Path -LiteralPath $changelog)) {
            Add-Failure 'version-sync' 'CHANGELOG.md not found.'
        } else {
            $newest = [regex]::Match((Get-Content -LiteralPath $changelog -Raw), '(?m)^## \[(\d+\.\d+\.\d+)\]')
            if (-not $newest.Success) {
                Add-Failure 'version-sync' "CHANGELOG.md has no '## [X.Y.Z]' release heading."
            } elseif ($newest.Groups[1].Value -ne $version) {
                Add-Failure 'version-sync' "Newest CHANGELOG release heading is [$($newest.Groups[1].Value)] but ADDIN_VERSION is $version."
            }
        }
        Write-Line "  ADDIN_VERSION is $version."
    }

    # -- 11 -----------------------------------------------------------------
    Write-Line "`n[11/$total] PowerShell scripts parse (tools/*.ps1)" 'Cyan'

    $toolsRoot = Join-Path $Root 'tools'
    $scripts = @()
    if (Test-Path -LiteralPath $toolsRoot) { $scripts = @(Get-ChildItem -LiteralPath $toolsRoot -Filter *.ps1 | Sort-Object Name) }
    foreach ($s in $scripts) {
        $tokens = $null
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($s.FullName, [ref]$tokens, [ref]$parseErrors)
        foreach ($e in @($parseErrors)) {
            Add-Failure 'ps-syntax' "tools/$($s.Name) line $($e.Extent.StartLineNumber) - $($e.Message)"
        }
    }
    Write-Line "  Parsed $($scripts.Count) script(s)."

    return $modules.Count
}

# ---------------------------------------------------------------------------
# Self-test: a synthetic tree that passes, then one seeded break per case.
# ---------------------------------------------------------------------------
function Write-FixtureFile {
    param([string]$Root, [string]$RelPath, [string]$Content)
    $path = Join-Path $Root $RelPath
    $dir  = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    # Fixtures spell a Rubberduck annotation's leading '@ as ~@: at the start of
    # a line, '@ would end the here-string the fixture is written in.
    $text = ($Content -replace '(?m)^~@', "'@")
    $text = ($text -replace "`r?`n", "`r`n")
    if (-not $text.EndsWith("`r`n")) { $text += "`r`n" }
    [System.IO.File]::WriteAllText($path, $text, $script:Latin1)
}

function Edit-FixtureFile {
    param([string]$Root, [string]$RelPath, [string]$Find, [string]$Replace)
    $path    = Join-Path $Root $RelPath
    $text    = [System.IO.File]::ReadAllText($path, $script:Latin1)
    $Find    = ($Find    -replace "`r?`n", "`r`n")
    $Replace = ($Replace -replace "`r?`n", "`r`n")
    $at = $text.IndexOf($Find, [System.StringComparison]::Ordinal)
    if ($at -lt 0) { throw "Self-test bug: text to replace not found in ${RelPath}: $Find" }
    $text = $text.Substring(0, $at) + $Replace + $text.Substring($at + $Find.Length)
    [System.IO.File]::WriteAllText($path, $text, $script:Latin1)
}

function New-SelfTestFixture {
    param([string]$Root)

    Write-FixtureFile $Root 'src/Core/modStartup.bas' @'
Attribute VB_Name = "modStartup"
~@Folder("TPD_Addin.Core")

Option Explicit

Public Const ADDIN_VERSION As String = "1.2.3"

Public Sub InitializeAddIn()
End Sub
'@

    Write-FixtureFile $Root 'src/Core/modPerformance.bas' @'
Attribute VB_Name = "modPerformance"
~@Folder("TPD_Addin.Core")

Option Explicit

Public Const PERF_DEMO As String = "Demo_Internal"

Public Sub WithPerformance(workProcName As String)
    Select Case workProcName
        Case PERF_DEMO
            Demo_Internal
        Case Else
            MsgBox "Unknown action: " & workProcName
    End Select
End Sub
'@

    Write-FixtureFile $Root 'src/Ribbon/modRibbonCallbacks.bas' @'
Attribute VB_Name = "modRibbonCallbacks"
~@Folder("TPD_Addin.Ribbon")

Option Explicit

Public gRibbon As Object

Public Sub RunDemo(control As Object)
    WithPerformance PERF_DEMO
End Sub

Public Sub RibbonOnLoad(ribbon As Object)
    Set gRibbon = ribbon
    InitializeAddIn
End Sub
'@

    # The baseline deliberately carries look-alikes that must NOT be reported:
    # a paren call inside a comment and inside a string, Call Helper(ws), a
    # function call in an expression, a local Const, an Enum, a Declare inside
    # #If, a line continuation, and an Application.Run with a variable.
    Write-FixtureFile $Root 'src/Demo/modMain_Demo.bas' @'
Attribute VB_Name = "modMain_Demo"
~@Folder("TPD_Addin.Demo")

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#End If

Private Const ROW_LIMIT As Long = 3
Private Const PREF_DEMO_KEY As String = "Demo_Key"

Private Enum DemoKind
    DEMO_KIND_A = 1
    DEMO_KIND_B
End Enum

Public Sub Demo_Internal()
    Const LOCAL_STEP As Long = 1
    Dim ws As Object
    Dim total As Long
    Dim macroName As String

    ' Helper (ws) in a comment is not a call.
    Debug.Print "Helper (ws) PERF_FAKE"
    Helper ws
    Call Helper(ws)
    total = CountRows(ws) + ROW_LIMIT + _
            LOCAL_STEP + DEMO_KIND_A
    If total > 0 Then Helper ws
    Debug.Print PREF_DEMO_KEY, DEMO_KIND_B

    Application.OnTime Now + TimeSerial(0, 0, 5), "ClearStatus"
    Application.Run "modMain_Demo.ClearStatus"
    macroName = "ClearStatus"
    Application.Run macroName
End Sub

Private Sub Helper(ByVal ws As Object)
End Sub

Private Function CountRows(ByVal ws As Object) As Long
    CountRows = 1
End Function

Public Sub ClearStatus()
    Application.StatusBar = False
End Sub
'@

    Write-FixtureFile $Root 'src/Forms/frmDemo.frm' @'
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDemo
   Caption         =   "Demo"
   OleObjectBlob   =   "frmDemo.frx":0000
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
~@Folder("TPD_Addin.Forms")

Option Explicit

Private mCancelled As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Private Sub UserForm_Initialize()
    mCancelled = False
End Sub
'@
    [System.IO.File]::WriteAllBytes((Join-Path $Root 'src/Forms/frmDemo.frx'), [byte[]](1, 2, 3, 4))

    Write-FixtureFile $Root 'src/Document/ThisWorkbook.cls' @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
~@Folder("TPD_Addin.Document")
'@

    Write-FixtureFile $Root 'tests/unit/modTest_Demo.bas' @'
Attribute VB_Name = "modTest_Demo"
~@Folder("TPD_Addin.Tests")
~@TestModule

Option Explicit
Option Private Module

~@TestMethod
Public Sub Demo_Runs()
    Demo_Internal
End Sub
'@

    Write-FixtureFile $Root 'customUI/customUI14.xml' @'
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
  <ribbon><tabs><tab id="tabDemo" label="Demo"><group id="grpDemo" label="Demo">
    <button id="btnDemo" label="Demo" onAction="RunDemo" />
  </group></tab></tabs></ribbon>
</customUI>
'@

    Write-FixtureFile $Root 'README.md' @'
# Demo

<!-- download-link: version managed by tools/sync-readme-version.sh -->
**[Download `TPD_Addin.xlam` - v1.2.3](https://example.invalid/releases/download/v1.2.3/TPD_Addin.xlam)**
'@

    Write-FixtureFile $Root 'CHANGELOG.md' @'
# Changelog

## [Unreleased]

## [1.2.3] - 2030-01-01

## [1.2.2] - 2029-12-01
'@

    Write-FixtureFile $Root 'tools/Demo-Tool.ps1' @'
param([string]$Name = 'world')
$greeting = "Hello, ${Name}: welcome"
Write-Host $greeting
'@
}

function Invoke-SelfTest {
    $demo = 'src/Demo/modMain_Demo.bas'
    $perf = 'src/Core/modPerformance.bas'

    $cases = @(
        @{ Name = 'Baseline tree (look-alikes included) is clean'; Expect = @(); Mutate = { param($r) } }

        @{ Name = 'Ribbon onAction names a missing Sub'; Expect = @('ribbon-callback')
           Mutate = { param($r) Edit-FixtureFile $r 'customUI/customUI14.xml' 'onAction="RunDemo"' 'onAction="RunMissing"' } }

        @{ Name = 'Module without Option Explicit'; Expect = @('option-explicit')
           Mutate = { param($r) Edit-FixtureFile $r $demo "Option Explicit`n" '' } }
        @{ Name = 'Test module without Option Explicit'; Expect = @('option-explicit')
           Mutate = { param($r) Edit-FixtureFile $r 'tests/unit/modTest_Demo.bas' "Option Explicit`n" '' } }

        @{ Name = 'Form tagged with a feature folder'; Expect = @('folder-tag')
           Mutate = { param($r) Edit-FixtureFile $r 'src/Forms/frmDemo.frm' '"TPD_Addin.Forms"' '"TPD_Addin.Demo"' } }
        @{ Name = 'Module with no @Folder tag'; Expect = @('folder-tag')
           Mutate = { param($r) Edit-FixtureFile $r $demo "'@Folder(""TPD_Addin.Demo"")`n" '' } }

        @{ Name = 'LF-only line endings'; Expect = @('line-endings')
           Mutate = { param($r)
               $p = Join-Path $r 'src/Forms/frmDemo.frm'
               $t = [System.IO.File]::ReadAllText($p, $script:Latin1) -replace "`r`n", "`n"
               [System.IO.File]::WriteAllText($p, $t, $script:Latin1) } }
        @{ Name = '.frm header names another form'; Expect = @('file-shape')
           Mutate = { param($r) Edit-FixtureFile $r 'src/Forms/frmDemo.frm' '} frmDemo' '} frmOther' } }
        @{ Name = '.frm without its .frx'; Expect = @('file-shape')
           Mutate = { param($r) Remove-Item -LiteralPath (Join-Path $r 'src/Forms/frmDemo.frx') } }
        @{ Name = 'Attribute VB_Name differs from the file name'; Expect = @('file-shape')
           Mutate = { param($r) Edit-FixtureFile $r 'src/Core/modStartup.bas' 'VB_Name = "modStartup"' 'VB_Name = "modBoot"' } }
        @{ Name = 'File saved with a UTF-8 BOM'; Expect = @('file-shape')
           Mutate = { param($r)
               $p = Join-Path $r 'src/Core/modStartup.bas'
               $b = [System.IO.File]::ReadAllBytes($p)
               [System.IO.File]::WriteAllBytes($p, ([byte[]](0xEF, 0xBB, 0xBF) + $b)) } }
        @{ Name = 'Two modules with one name'; Expect = @('duplicate-name')
           Mutate = { param($r) Copy-Item -LiteralPath (Join-Path $r 'src/Core/modStartup.bas') -Destination (Join-Path $r 'src/Demo/modStartup.bas') } }

        @{ Name = 'Private Const below the first procedure (the 2026-09-08 bug)'; Expect = @('decl-after-proc')
           Mutate = { param($r) Edit-FixtureFile $r $demo "Private Sub Helper(" "Private Const LATE_LIMIT As Long = 9`n`nPrivate Sub Helper(" } }
        @{ Name = 'Module-level Dim at the end of a module'; Expect = @('decl-after-proc')
           Mutate = { param($r) Edit-FixtureFile $r $demo "    Application.StatusBar = False`nEnd Sub`n" "    Application.StatusBar = False`nEnd Sub`n`nDim mLate As Long`n" } }

        @{ Name = 'Sub called as Name (arg)'; Expect = @('paren-call')
           Mutate = { param($r) Edit-FixtureFile $r $demo "    Helper ws`n    Call" "    Helper (ws)`n    Call" } }
        @{ Name = 'Single-line If ... Then Name (arg)'; Expect = @('paren-call')
           Mutate = { param($r) Edit-FixtureFile $r $demo 'Then Helper ws' 'Then Helper (ws)' } }

        @{ Name = 'OnTime names a missing Sub'; Expect = @('string-proc-ref')
           Mutate = { param($r) Edit-FixtureFile $r $demo ', "ClearStatus"' ', "ClearStatusBar"' } }
        @{ Name = 'Application.Run names a missing Sub'; Expect = @('string-proc-ref')
           Mutate = { param($r) Edit-FixtureFile $r $demo '"modMain_Demo.ClearStatus"' '"modMain_Demo.ClearIt"' } }

        @{ Name = 'PERF_* constant with no Case'; Expect = @('perf-parity')
           Mutate = { param($r) Edit-FixtureFile $r $perf "        Case PERF_DEMO`n            Demo_Internal`n" '' } }
        @{ Name = 'PERF_* value names a missing Sub'; Expect = @('perf-parity')
           Mutate = { param($r) Edit-FixtureFile $r $perf '= "Demo_Internal"' '= "Demo_Intrnal"' } }
        @{ Name = 'Case runs a different Sub than its constant names'; Expect = @('perf-parity')
           Mutate = { param($r) Edit-FixtureFile $r $perf "            Demo_Internal`n" "            InitializeAddIn`n" } }
        @{ Name = 'PERF_* constant no caller passes'; Expect = @('perf-parity')
           Mutate = { param($r)
               Edit-FixtureFile $r $perf "Public Sub WithPerformance" "Public Const PERF_SPARE As String = ""ClearStatus""`n`nPublic Sub WithPerformance"
               Edit-FixtureFile $r $perf "        Case Else" "        Case PERF_SPARE`n            ClearStatus`n        Case Else" } }

        @{ Name = 'Mistyped constant'; Expect = @('undeclared-const')
           Mutate = { param($r) Edit-FixtureFile $r $demo 'PREF_DEMO_KEY, DEMO_KIND_B' 'PREF_DEMO_KEY, PREF_DEMO_KEEY' } }
        @{ Name = 'Private constant used from another module'; Expect = @('undeclared-const')
           Mutate = { param($r) Edit-FixtureFile $r 'src/Ribbon/modRibbonCallbacks.bas' '    InitializeAddIn' "    InitializeAddIn`n    Debug.Print ROW_LIMIT" } }

        @{ Name = 'README link left on the old version'; Expect = @('version-sync')
           Mutate = { param($r) Edit-FixtureFile $r 'README.md' 'download/v1.2.3/' 'download/v1.2.2/' } }
        @{ Name = 'CHANGELOG has no heading for the version'; Expect = @('version-sync')
           Mutate = { param($r) Edit-FixtureFile $r 'CHANGELOG.md' "## [1.2.3] - 2030-01-01`n" '' } }
        @{ Name = 'ADDIN_VERSION bumped on its own'; Expect = @('version-sync')
           Mutate = { param($r) Edit-FixtureFile $r 'src/Core/modStartup.bas' '"1.2.3"' '"1.2.4"' } }

        # The real one: "$label: text" reads as a drive-qualified variable.
        @{ Name = 'Tools script with a "$name:" syntax error'; Expect = @('ps-syntax')
           Mutate = { param($r) Edit-FixtureFile $r 'tools/Demo-Tool.ps1' '${Name}:' '$Name:' } }
    )

    $work = Join-Path ([System.IO.Path]::GetTempPath()) ('TPD_StaticSelfTest_' + [Guid]::NewGuid().ToString('N'))
    $base = Join-Path $work 'base'
    $bad  = 0

    try {
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        New-SelfTestFixture $base

        Write-Host "`nSelf-test: $($cases.Count) case(s), one seeded break each" -ForegroundColor Cyan
        $n = 0
        foreach ($case in $cases) {
            $n++
            $caseRoot = Join-Path $work ("case{0:d2}" -f $n)
            Copy-Item -LiteralPath $base -Destination $caseRoot -Recurse
            & $case.Mutate $caseRoot

            $script:Quiet = $true
            try { [void](Invoke-StaticChecks -Root $caseRoot) } finally { $script:Quiet = $false }

            $got  = @($script:Failures | ForEach-Object { $_.Id } | Sort-Object -Unique)
            $want = @($case.Expect | Sort-Object)
            $wantText = '(none)'
            if ($want.Count -gt 0) { $wantText = $want -join ', ' }
            $gotText = '(none)'
            if ($got.Count -gt 0) { $gotText = $got -join ', ' }

            if ($gotText -eq $wantText) {
                Write-Host ("  ok    {0,-62} -> {1}" -f $case.Name, $gotText)
            } else {
                $bad++
                Write-Host ("  WRONG {0,-62} -> expected {1}, got {2}" -f $case.Name, $wantText, $gotText) -ForegroundColor Red
                foreach ($f in $script:Failures) { Write-Host "          [$($f.Id)] $($f.Message)" -ForegroundColor DarkGray }
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    if ($bad -gt 0) {
        Write-Host "Self-test: $bad of $($cases.Count) case(s) did not behave as expected." -ForegroundColor Red
        return 1
    }
    Write-Host "Self-test passed: all $($cases.Count) case(s) reported exactly the expected check." -ForegroundColor Green
    return 0
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if ($SelfTest) {
    exit (Invoke-SelfTest)
}

if ([string]::IsNullOrEmpty($RepoRoot)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $RepoRoot = Split-Path -Parent $scriptDir
}

$moduleCount = Invoke-StaticChecks -Root $RepoRoot

Write-Host ''
if ($script:Failures.Count -gt 0) {
    Write-Host "$($script:Failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
} else {
    Write-Host "All static checks passed ($moduleCount modules)." -ForegroundColor Green
    exit 0
}
