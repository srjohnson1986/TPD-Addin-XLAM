VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilenameOptions 
   Caption         =   "Append Filename Text"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5448
   OleObjectBlob   =   "frmFilenameOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilenameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TPD_Addin.SplitExport")

Option Explicit

Public UserAppend As String
Public IncludeDate As Boolean
Public Cancelled As Boolean

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

