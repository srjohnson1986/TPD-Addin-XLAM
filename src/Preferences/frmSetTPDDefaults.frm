VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetTPDDefaults 
   Caption         =   "Set TPD Defaults"
   ClientHeight    =   7188
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13320
   OleObjectBlob   =   "frmSetTPDDefaults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetTPDDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("TPD_Addin.Preferences")

Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim tempPic As StdPicture

    Set ws = ThisWorkbook.Worksheets("_Resources")
    Set shp = ws.Shapes("DefaultLogo")

    ' Copy the shape to the clipboard
    shp.Copy

    ' Retrieve the clipboard image as a StdPicture
    Set tempPic = PastePicture()

    ' Assign to the Image control
    Me.imgDefaultLogo.Picture = tempPic
    Me.imgDefaultLogo.PictureSizeMode = fmPictureSizeModeZoom
End Sub

