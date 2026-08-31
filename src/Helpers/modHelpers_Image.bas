Attribute VB_Name = "modHelpers_Image"
'@Folder("TPD_Addin.Helpers")

Option Explicit

Public Sub SafeInsertLogoAtRight(ws As Worksheet, imgPath As String, headerRow As Long, dataHeaderRow As Long)
    On Error GoTo LogoErr

    If Len(Dir(imgPath)) = 0 Then Exit Sub

    InsertLogoAtRight ws, imgPath, headerRow, dataHeaderRow
    Exit Sub

LogoErr:
    MsgBox "Unable to insert logo image: " & imgPath, vbExclamation
End Sub

Public Function PastePicture() As StdPicture
    ' Returns a picture object from the clipboard
    Dim IData As Object
    Set IData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    IData.GetData 1
    Set PastePicture = IData.GetData(1)
End Function

