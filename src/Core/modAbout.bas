Attribute VB_Name = "modAbout"
'@Folder("TPD_Addin.Core")

'===========================================================
'  Add-in identity: the outbound links surfaced on the Set
'  Defaults dialog ([#94]). The URLs live here, in one place,
'  and open through the shared ThisWorkbook.FollowHyperlink
'  wrapper - FollowHyperlink just hands the URL to the default
'  browser, so this is not a network call of our own (the
'  add-in never checks GitHub for a newer release - #94).
'
'  [#94]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/94
'===========================================================

Option Explicit

' The end-user guide - the wiki home, which is re-synced from
' docs/USER_GUIDE.md on each release ([#94], [#132]).
'
' [#132]: https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/132
Public Const ABOUT_USER_GUIDE_URL As String = _
    "https://github.com/srjohnson1986/TPD-Addin-XLAM/wiki"

Private Const ABOUT_RELEASES_URL As String = _
    "https://github.com/srjohnson1986/TPD-Addin-XLAM/releases"

' The GitHub release page for the running version, e.g.
' .../releases/tag/v2.4.0. A dev build whose tag is not
' published yet lands on GitHub's "tag not found" page, which
' links straight to the Releases list - acceptable, and it
' resolves the moment that version ships.
Public Function AboutReleaseUrl() As String
    AboutReleaseUrl = ABOUT_RELEASES_URL & "/tag/v" & ADDIN_VERSION
End Function

Public Sub OpenReleasePage()
    OpenAboutUrl AboutReleaseUrl()
End Sub

Public Sub OpenUserGuide()
    OpenAboutUrl ABOUT_USER_GUIDE_URL
End Sub

' One place for the "open a URL in the browser" call. Swallows
' errors (a missing browser association, an offline machine) -
' a dead link must not throw an unhandled error out of a form
' click handler.
Private Sub OpenAboutUrl(ByVal url As String)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink url
    On Error GoTo 0
End Sub
