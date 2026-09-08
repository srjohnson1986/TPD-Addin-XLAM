Attribute VB_Name = "modHelpers_DialogChrome"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  Shared run-time chrome for the TPD dialogs (#118).
'
'  The picker / split / export forms are brought in line with
'  frmSetTPDDefaults: a TPD-grey brand band, a muted helper
'  line, and blue underlined About links bottom-left. The
'  styling that frmSetTPDDefaults applies in code rather than
'  the .frx (so it stays in source review) is centralised here
'  so the four forms cannot drift apart.
'
'  Each form places the controls by hand in the VBE and calls
'  ApplyDialogChrome + InitAboutLinks from UserForm_Initialize;
'  the lblVersion / lblUserGuide Click handlers stay on the
'  form (three lines each) because a standard module can't wire
'  an event.
'
'  Colour note: VBA BackColor / ForeColor literals are BGR, and
'  RGB() already returns BGR, so RGB(r, g, b) is the readable
'  way to write these - same as frmSetTPDDefaults.
'===========================================================

Option Explicit

'-----------------------------------------------------------
' Brand band + help line. Pass Nothing for anything a given
' form doesn't have.
'   brandBar  - full-width 28 pt Label flush to the top
'   helpLabel - the one-line help text under the title
'                (lblHelp, matching frmSetTPDDefaults' lbl*Help)
'-----------------------------------------------------------
Public Sub ApplyDialogChrome(ByVal brandBar As MSForms.Label, _
                             ByVal helpLabel As MSForms.Label)
    If Not brandBar Is Nothing Then
        brandBar.BackColor = &H5D6664       ' TPD grey #64665D (BGR)
        brandBar.Caption = vbNullString
    End If

    If Not helpLabel Is Nothing Then
        helpLabel.ForeColor = RGB(74, 74, 74)   ' #4A4A4A
    End If
End Sub

'-----------------------------------------------------------
' The bottom-left About links, identical to frmSetTPDDefaults:
'   lblVersion   -> this version's GitHub release page
'   lblUserGuide -> the user guide (wiki)
' Captions, tooltips and the blue-underline link look are set
' here; the form wires the two Click handlers to modAbout.
'-----------------------------------------------------------
Public Sub InitAboutLinks(ByVal lblVersion As MSForms.Label, _
                          ByVal lblUserGuide As MSForms.Label)
    If Not lblVersion Is Nothing Then
        lblVersion.Caption = "v" & ADDIN_VERSION
        lblVersion.ControlTipText = AboutReleaseUrl()
        StyleAsLink lblVersion
    End If

    If Not lblUserGuide Is Nothing Then
        lblUserGuide.Caption = "User guide"
        lblUserGuide.ControlTipText = modAbout.ABOUT_USER_GUIDE_URL
        StyleAsLink lblUserGuide
    End If
End Sub

Private Sub StyleAsLink(ByVal lbl As MSForms.Label)
    lbl.ForeColor = RGB(0, 102, 204)   ' hyperlink blue
    lbl.Font.Underline = True
End Sub
