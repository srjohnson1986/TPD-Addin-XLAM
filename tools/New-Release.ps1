<#
.SYNOPSIS
    Prepares a release of the add-in: checks the repo is ready, runs the local
    gate, and creates a DRAFT GitHub release with the built .xlam attached.

.DESCRIPTION
    Run this on main, after the release PR (version bump + CHANGELOG entry +
    user guide) has merged. It never publishes anything: the release it
    creates is a draft, which has no tag and is visible only to people with
    write access. Publishing it on GitHub is the human step - that is what
    creates the vX.Y.Z tag, makes the README's download link work, and
    triggers the wiki sync (.github/workflows/wiki-sync.yml).

    What it does, in order:

      1. Reads ADDIN_VERSION from src\Core\modStartup.bas.
      2. Checks the repo is ready to release (all problems are listed at once):
           - on main, nothing uncommitted, level with origin/main
           - no vX.Y.Z tag and no release for it already
           - CHANGELOG.md has a "## [X.Y.Z]" section with something in it
           - "## [Unreleased]" is empty (anything left there would be missed)
      3. Runs tools\Test-All.ps1 - static checks, build, drift, unit tests.
      4. Takes the SHA-256 of build\TPD_Addin.xlam.
      5. Uses that CHANGELOG section, as written, for the release notes, and
         adds an "Install / update" footer carrying the file's SHA-256.
      6. Creates the draft release vX.Y.Z on this commit with TPD_Addin.xlam
         attached, and prints its URL.

    The asset is always named TPD_Addin.xlam and the tag vX.Y.Z, because the
    README's pinned link is releases/download/vX.Y.Z/TPD_Addin.xlam.

    NOT DONE HERE, and still yours: Debug > Compile VBAProject on the build,
    a Test Explorer run, the ribbon flows against the fixtures, refreshing
    the base file if the ribbon or sheets changed, and clicking Publish.

.PARAMETER DryRun
    Do everything except create the release, and show what would be created.
    Unlike a real run it carries on past a failed readiness check, so one
    dry run shows every problem.

.PARAMETER SkipGate
    With -DryRun only: do not run Test-All.ps1 (so Excel is not started) and
    describe whatever build is already in build\. A real run cannot skip it.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\New-Release.ps1 -DryRun
    powershell -ExecutionPolicy Bypass -File tools\New-Release.ps1
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$SkipGate
)

$ErrorActionPreference = 'Stop'

$toolsDir  = $PSScriptRoot
$repoRoot  = Split-Path -Parent $toolsDir
$xlamPath  = Join-Path (Join-Path $repoRoot 'build') 'TPD_Addin.xlam'
$changelog = Join-Path $repoRoot 'CHANGELOG.md'
$startup   = Join-Path (Join-Path (Join-Path $repoRoot 'src') 'Core') 'modStartup.bas'
$hostExe   = (Get-Process -Id $PID).Path

if ($SkipGate -and -not $DryRun) {
    Write-Host '-SkipGate is only allowed with -DryRun. A real release always runs the gate.' -ForegroundColor Red
    exit 2
}

# Runs git / gh and returns the exit code and output. Standard error is dropped
# and never treated as a PowerShell error, whatever $ErrorActionPreference says.
function Invoke-Native {
    param([string]$Exe, [string[]]$Arguments)
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @Arguments 2>$null
        return @{ Code = $LASTEXITCODE; Out = @($out) }
    } finally {
        $ErrorActionPreference = $old
    }
}

function Invoke-Git { param([string[]]$Arguments) Invoke-Native -Exe 'git' -Arguments (@('-C', $repoRoot) + $Arguments) }

$gh = $null
$ghCommand = Get-Command gh -ErrorAction SilentlyContinue
if ($ghCommand) { $gh = $ghCommand.Source }
elseif (Test-Path (Join-Path $env:ProgramFiles 'GitHub CLI\gh.exe')) { $gh = Join-Path $env:ProgramFiles 'GitHub CLI\gh.exe' }

$problems = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Write-Step { param([string]$Text) Write-Host "`n$Text" -ForegroundColor Cyan }

# ---------------------------------------------------------------------------
# 1. Version
# ---------------------------------------------------------------------------

$versionMatch = [regex]::Match([System.IO.File]::ReadAllText($startup),
                               '(?m)^\s*Public\s+Const\s+ADDIN_VERSION\s+As\s+String\s*=\s*"(\d+\.\d+\.\d+)"')
if (-not $versionMatch.Success) {
    Write-Host "No ADDIN_VERSION found in $startup" -ForegroundColor Red
    exit 2
}
$version = $versionMatch.Groups[1].Value
$tag     = "v$version"

$mode = 'RELEASE'
if ($DryRun) { $mode = 'DRY RUN - nothing will be created' }
Write-Host "TPD add-in release $tag  ($mode)"

# ---------------------------------------------------------------------------
# 2. Is the repo ready?
# ---------------------------------------------------------------------------

Write-Step '[1/4] Is the repo ready to release?'

$slug = ''
$remote = Invoke-Git @('remote', 'get-url', 'origin')
if ($remote.Code -eq 0 -and "$($remote.Out[0])" -match 'github\.com[:/]([^/]+)/(.+?)(\.git)?$') {
    $slug = "$($Matches[1])/$($Matches[2])"
} else {
    $problems.Add('Could not work out the GitHub repository from the origin remote.')
}

[void](Invoke-Git @('fetch', 'origin', '--prune', '--tags', '--quiet'))

$branch = "$((Invoke-Git @('rev-parse', '--abbrev-ref', 'HEAD')).Out[0])"
$head   = "$((Invoke-Git @('rev-parse', 'HEAD')).Out[0])"
$remoteMain = "$((Invoke-Git @('rev-parse', 'origin/main')).Out[0])"
$dirty  = @((Invoke-Git @('status', '--porcelain')).Out | Where-Object { $_ })

if ($branch -ne 'main')   { $problems.Add("On branch '$branch', not main.") }
if ($dirty.Count -gt 0)   { $problems.Add("$($dirty.Count) uncommitted change(s) in the working tree.") }
if ($head -ne $remoteMain) { $problems.Add('HEAD is not the same commit as origin/main (pull, or merge the release PR first).') }

$tagOnRemote = Invoke-Git @('ls-remote', '--tags', 'origin', "refs/tags/$tag")
if (@($tagOnRemote.Out | Where-Object { $_ }).Count -gt 0) { $problems.Add("Tag $tag already exists on GitHub - $version has been released. Bump ADDIN_VERSION first.") }

if (-not $gh) {
    $problems.Add('GitHub CLI (gh) not found. Install it, or add it to PATH.')
} elseif ($slug) {
    $existing = Invoke-Native -Exe $gh -Arguments @('release', 'list', '--repo', $slug, '--limit', '100', '--json', 'tagName,isDraft')
    if ($existing.Code -ne 0) {
        $problems.Add('Could not list releases with gh (not signed in? run: gh auth status).')
    } else {
        # Assigned first, then looped: Windows PowerShell hands a JSON array down
        # the pipeline as one object, so filtering it directly matches nothing.
        $releases = ConvertFrom-Json ($existing.Out -join '')
        $match = $null
        foreach ($r in $releases) { if ($r.tagName -eq $tag) { $match = $r; break } }
        if ($match) {
            if ($match.isDraft) {
                $problems.Add("A draft release for $tag already exists on GitHub. Delete that draft first, or publish it.")
            } else {
                $problems.Add("A published release for $tag already exists on GitHub.")
            }
        }
    }
}

# CHANGELOG: this version's section is the release notes.
$logText  = [System.IO.File]::ReadAllText($changelog) -replace "`r`n", "`n"
$section  = [regex]::Match($logText, '(?ms)^## \[' + [regex]::Escape($version) + '\]([^\n]*)\n(.*?)(?=^## \[|\z)')
$notes    = ''
if (-not $section.Success) {
    $problems.Add("CHANGELOG.md has no '## [$version]' section.")
} else {
    $notes = $section.Groups[2].Value.Trim()
    if ($notes.Length -eq 0) { $problems.Add("The '## [$version]' section of CHANGELOG.md is empty.") }
    $dated = [regex]::Match($section.Groups[1].Value, '(\d{4}-\d{2}-\d{2})')
    $today = Get-Date -Format 'yyyy-MM-dd'
    if (-not $dated.Success) { $warnings.Add("The '## [$version]' heading has no date.") }
    elseif ($dated.Groups[1].Value -ne $today) { $warnings.Add("The '## [$version]' heading is dated $($dated.Groups[1].Value); today is $today.") }
}
$unreleased = [regex]::Match($logText, '(?ms)^## \[Unreleased\][^\n]*\n(.*?)(?=^## \[|\z)')
if ($unreleased.Success -and $unreleased.Groups[1].Value.Trim().Length -gt 0) {
    $problems.Add("CHANGELOG.md still has entries under '## [Unreleased]' - they would be left out of $tag.")
}

if ($problems.Count -eq 0) { Write-Host '      yes' }
foreach ($p in $problems) { Write-Host "      PROBLEM: $p" -ForegroundColor Red }
foreach ($w in $warnings) { Write-Host "      note: $w" -ForegroundColor Yellow }

if ($problems.Count -gt 0 -and -not $DryRun) {
    Write-Host "`nNot ready - nothing was built or created." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# 3. The gate
# ---------------------------------------------------------------------------

Write-Step '[2/4] Local gate (tools\Test-All.ps1)'

$builtFresh = $false
if ($SkipGate) {
    Write-Host '      skipped (-SkipGate). Excel was not started; the build below is whatever was already there.' -ForegroundColor Yellow
} else {
    $gate = Start-Process -FilePath $hostExe -Wait -PassThru -NoNewWindow -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $toolsDir 'Test-All.ps1')`"")
    if ($gate.ExitCode -ne 0) {
        $problems.Add('The local gate failed - see its output above.')
        if (-not $DryRun) {
            Write-Host "`nThe gate failed - no release was created." -ForegroundColor Red
            exit 1
        }
    } else {
        $builtFresh = $true
    }
}

# ---------------------------------------------------------------------------
# 4. The asset and the notes
# ---------------------------------------------------------------------------

Write-Step '[3/4] What goes into the release'

$sha256 = ''
$size   = 0
if (Test-Path -LiteralPath $xlamPath) {
    $sha256 = (Get-FileHash -LiteralPath $xlamPath -Algorithm SHA256).Hash
    $item   = Get-Item -LiteralPath $xlamPath
    $size   = $item.Length
    $origin = 'built just now by the gate'
    if (-not $builtFresh) { $origin = "NOT rebuilt - last written $($item.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))" }
    Write-Host "      tag and title : $tag  (the tag is created when the draft is published)"
    Write-Host "      on commit     : $head"
    Write-Host "      asset         : TPD_Addin.xlam, $size bytes, $origin"
    Write-Host "      SHA-256       : $sha256"
} else {
    $problems.Add("No build at $xlamPath.")
    Write-Host "      PROBLEM: no build at $xlamPath" -ForegroundColor Red
}

# The CHANGELOG section, then the same footer v2.4.5 was published with: how to
# install, and the hash of the attached file so a download can be checked.
$notes = $notes + "`n`n" + (@(
    '## Install / update',
    '',
    "Download ``TPD_Addin.xlam`` below and replace the old file in your ``AddIns`` folder, keeping the same filename. Your saved defaults and logo carry over. Full steps are in the [user guide](https://github.com/$slug/wiki).",
    '',
    "``TPD_Addin.xlam`` SHA-256: ``$sha256``"
) -join "`n")

$noteLines = @($notes -split "`n")
Write-Host "      release notes : $($noteLines.Count) line(s): the '## [$version]' section of CHANGELOG.md plus an install footer, starting:"
$noteLines | Select-Object -First 8 | ForEach-Object { Write-Host "        | $_" -ForegroundColor DarkGray }
if ($noteLines.Count -gt 8) { Write-Host "        | ... $($noteLines.Count - 8) more line(s)" -ForegroundColor DarkGray }

# ---------------------------------------------------------------------------
# 5. Create the draft
# ---------------------------------------------------------------------------

Write-Step '[4/4] Draft release'

if ($DryRun) {
    $preview = Join-Path ([System.IO.Path]::GetTempPath()) ("TPD_release_notes_$version.preview.md")
    [System.IO.File]::WriteAllText($preview, $notes + "`n", (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "      DRY RUN: nothing was created. A real run would now run:"
    Write-Host "        gh release create $tag build\TPD_Addin.xlam --repo $slug --draft --title $tag --notes-file <notes> --target $head"
    Write-Host "      The full release notes are in: $preview"
    Write-Host ''
    if ($problems.Count -gt 0) {
        Write-Host "RESULT: a real run would stop - $($problems.Count) problem(s):" -ForegroundColor Red
        $problems | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        exit 1
    }
    Write-Host 'RESULT: ready. Run again without -DryRun to create the draft.' -ForegroundColor Green
    exit 0
}

$notesFile = Join-Path ([System.IO.Path]::GetTempPath()) ("TPD_release_notes_$version.md")
[System.IO.File]::WriteAllText($notesFile, $notes + "`n", (New-Object System.Text.UTF8Encoding($false)))
try {
    $created = Invoke-Native -Exe $gh -Arguments @('release', 'create', $tag, $xlamPath, '--repo', $slug, '--draft',
                                                   '--title', $tag, '--notes-file', $notesFile, '--target', $head)
} finally {
    Remove-Item -LiteralPath $notesFile -Force -ErrorAction SilentlyContinue
}
if ($created.Code -ne 0) {
    Write-Host "      gh release create failed (exit $($created.Code)). Check https://github.com/$slug/releases for a half-made draft." -ForegroundColor Red
    exit 1
}

Write-Host "      created: $($created.Out | Select-Object -Last 1)" -ForegroundColor Green
Write-Host ''
Write-Host "Draft release $tag is ready for you. It is not public and has no tag yet."
Write-Host 'Still to do, by hand:'
Write-Host '  1. Debug > Compile VBAProject on build\TPD_Addin.xlam, and the fixtures smoke test, if not done already.'
Write-Host '  2. Read the draft on GitHub, then Publish it. That creates the tag and starts the wiki sync.'
Write-Host "  3. Check the README download link: https://github.com/$slug/releases/download/$tag/TPD_Addin.xlam"
Write-Host '  4. If the ribbon, sheets or logo changed, refresh build\_base\TPD_Addin_base.xlam (CONTRIBUTING.md).'
Write-Host "SHA-256 of the attached TPD_Addin.xlam: $sha256"
exit 0
