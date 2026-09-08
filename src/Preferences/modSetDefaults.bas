Attribute VB_Name = "modSetDefaults"
'@Folder("TPD_Addin.Preferences")

'===========================================================
'  The "Set Defaults" ribbon entry point - opens the modal
'  frmSetTPDDefaults, the one place to configure the default
'  column lists the per-run pickers fall back to. The only
'  onAction callback that isn't a one-line forwarder in
'  modRibbonCallbacks, because it does real work after the
'  dialog closes.
'
'  Was modMain_SetTPDDefaults: a 41-line entry-point shim is
'  not a "main" like modMain_CustEQList, and the ribbon has
'  said "Set Defaults" rather than "Set TPD Defaults" since
'  v2.4.1.
'
'  ClearTPDDefaultsStatusBar keeps its name deliberately:
'  frmSetTPDDefaults schedules it by STRING in its
'  Application.OnTime call, which no compiler checks, so a
'  rename would silently stop the status bar ever clearing. It
'  lives here rather than on the form because OnTime can only
'  call a public sub in a standard module.
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
