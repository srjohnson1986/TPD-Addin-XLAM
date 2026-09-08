VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilenameOptions 
   Caption         =   "Save Each Sheet to XLSX"
   ClientHeight    =   4836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8184
   OleObjectBlob   =   "frmFilenameOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilenameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("TPD_Addin.Forms")

Option Explicit

Public UserAppend As String
Public IncludeDate As Boolean
Public Cancelled As Boolean

' Seeded by the caller before .Show so the live preview reads like a real
' output name; falls back to "Sheet1" in PreviewExportFileName when unset.
Public SampleSheetName As String

'===========================================================
' Chrome + live preview (#118). Brand band / help line / About
' links are shared with the other dialogs (modHelpers_DialogChrome).
' lblFullFilename sits inside fraPreview and recomputes as the
' user types or toggles the date box - PreviewExportFileName is
' the same builder the export flow uses, so they can't disagree.
'===========================================================
Private Sub UserForm_Initialize()
    ApplyDialogChrome lblBrandBar, lblHelp
    InitAboutLinks lblVersion, lblUserGuide
End Sub

' The first preview is drawn in Activate, not Initialize (#154): assigning
' frm.SampleSheetName is the caller's first member access, so it runs
' Initialize before that assignment lands - Initialize would build the
' preview with an empty sample name. Activate fires on .Show, after the
' caller has set SampleSheetName.
Private Sub UserForm_Activate()
    RefreshPreview
End Sub

Private Sub RefreshPreview()
    lblFullFilename.Caption = PreviewExportFileName( _
        SampleSheetName, txtAppend.Text, CBool(chkIncludeDate.value))
End Sub

Private Sub txtAppend_Change()
    RefreshPreview
End Sub

Private Sub chkIncludeDate_Click()
    RefreshPreview
End Sub

Private Sub lblVersion_Click()
    modAbout.OpenReleasePage
End Sub

Private Sub lblUserGuide_Click()
    modAbout.OpenUserGuide
End Sub

Private Sub cmdOK_Click()
    ' Allow blank input
    UserAppend = txtAppend.Text
    IncludeDate = chkIncludeDate.value
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Me.Hide
        Cancel = True
    End If
End Sub

