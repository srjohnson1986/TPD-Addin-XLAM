<#
.SYNOPSIS
    Regenerates the wiki's Home.md from docs/USER_GUIDE.md.

.DESCRIPTION
    Home.md is the user guide with one addition: the "Do not edit this page
    in the wiki editor" notice under the title. The new Home.md is built as

        the guide's own title line (its "# ..." heading)
        the notice, copied from the Home.md that is there now
        everything in the guide after its title

    Nothing else is changed: the guide's text goes across as written. Only
    Home.md is touched; _Sidebar.md and any other wiki page are hand-kept.

    The notice is taken from the existing Home.md so its wording has one
    home. If the existing page has no such notice the script stops and writes
    nothing, so it cannot quietly drop it.

    Pure text, no Excel: runs under Windows PowerShell or pwsh on any OS.
    .github/workflows/wiki-sync.yml runs it when a release is published; it
    can be run by hand against a clone of the wiki the same way.

.PARAMETER GuidePath
    Path to docs/USER_GUIDE.md.

.PARAMETER HomePath
    Path to Home.md in a checkout of the wiki repository. Read for the
    notice, then overwritten.

.PARAMETER Check
    Write nothing. Exit 0 if Home.md is already what would be generated, 1 if
    it differs.

.EXAMPLE
    git clone https://github.com/srjohnson1986/TPD-Addin-XLAM.wiki.git ..\wiki
    powershell -ExecutionPolicy Bypass -File tools\Sync-WikiHome.ps1 -GuidePath docs\USER_GUIDE.md -HomePath ..\wiki\Home.md
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$GuidePath,
    [Parameter(Mandatory = $true)][string]$HomePath,
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

function Read-Lines {
    param([string]$Path)
    $text = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path)
    $text = $text.TrimStart([char]0xFEFF) -replace "`r`n", "`n" -replace "`r", "`n"
    return ,@($text.TrimEnd("`n") -split "`n")
}

$guide = Read-Lines $GuidePath
if ($guide.Count -eq 0 -or $guide[0] -notmatch '^# \S') {
    throw "$GuidePath does not start with a '# Title' line."
}
$title = $guide[0]

# Everything after the title, minus the blank line(s) right under it.
$bodyStart = 1
while ($bodyStart -lt $guide.Count -and $guide[$bodyStart].Trim().Length -eq 0) { $bodyStart++ }
if ($bodyStart -ge $guide.Count) { throw "$GuidePath has a title and nothing else." }
$body = $guide[$bodyStart..($guide.Count - 1)]

# The notice: the first block of "> ..." lines near the top of the current page.
$current = Read-Lines $HomePath
$notice  = New-Object System.Collections.Generic.List[string]
$limit   = [Math]::Min($current.Count, 12)
for ($i = 0; $i -lt $limit; $i++) {
    if ($current[$i].StartsWith('>')) {
        while ($i -lt $current.Count -and $current[$i].StartsWith('>')) { $notice.Add($current[$i]); $i++ }
        break
    }
}
if ($notice.Count -eq 0 -or (($notice -join ' ') -notmatch 'Do not edit this page')) {
    throw "$HomePath has no 'Do not edit this page' notice near the top. Not writing anything - restore the notice first."
}

$newLines = @($title, '') + @($notice) + @('') + @($body)
$newText  = ($newLines -join "`n") + "`n"
$oldText  = ($current -join "`n") + "`n"

if ($newText -ceq $oldText) {
    Write-Host "Home.md is already up to date ($($newLines.Count) lines)."
    exit 0
}

if ($Check) {
    Write-Host "Home.md differs from what docs/USER_GUIDE.md would generate ($($current.Count) lines now, $($newLines.Count) generated)."
    exit 1
}

[System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $HomePath).Path, $newText, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Home.md rewritten from the user guide ($($current.Count) lines before, $($newLines.Count) now)."
exit 0
