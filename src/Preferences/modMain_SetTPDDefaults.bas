Attribute VB_Name = "modMain_SetTPDDefaults"
'@Folder("TPD_Addin.Preferences")

'===========================================================
'  "Set TPD Defaults" ribbon entry point - opens the modal
'  frmSetTPDDefaults, the one place to configure the default
'  column lists the per-run pickers fall back to.
'
'  ClearTPDDefaultsStatusBar lives here (not on the form)
'  because Application.OnTime can only call a public sub in a
'  standard module.
'===========================================================

Option Explicit

' Ribbon callback
Public Sub RunSetTPDDefaults(control As IRibbonControl)
    frmSetTPDDefaults.Show vbModal
    Unload frmSetTPDDefaults
End Sub

' Scheduled by frmSetTPDDefaults after a successful save to wipe the
' "TPD defaults saved" status-bar confirmation.
Public Sub ClearTPDDefaultsStatusBar()
    Application.StatusBar = False
End Sub
