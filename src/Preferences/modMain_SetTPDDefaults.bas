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

' Set by frmSetTPDDefaults.cmdSaveRun_Click to a PERF_* constant naming the
' one-click flow to run once the dialog closes; "" otherwise (#100). Read and
' cleared by RunSetTPDDefaults after the form unloads - running it there keeps
' the perf wrapper's screen-updating / calc toggling clear of the modal form.
Public gPendingDefaultsFlow As String

' Ribbon callback
Public Sub RunSetTPDDefaults(control As IRibbonControl)
    Dim flow As String

    gPendingDefaultsFlow = vbNullString
    frmSetTPDDefaults.Show vbModal
    Unload frmSetTPDDefaults

    If Len(gPendingDefaultsFlow) = 0 Then Exit Sub

    flow = gPendingDefaultsFlow
    gPendingDefaultsFlow = vbNullString      ' clear before running
    WithPerformance flow
End Sub

' Scheduled by frmSetTPDDefaults after a successful save to wipe the
' "TPD defaults saved" status-bar confirmation.
Public Sub ClearTPDDefaultsStatusBar()
    Application.StatusBar = False
End Sub
